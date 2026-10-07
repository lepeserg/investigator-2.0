# -*- coding: utf-8 -*-
"""
av_ingest.py - prijom audio/video dlja investigator-sk (vsjo lokal'no/oflajn).

Podkomandy:
  info   <file>                          -> JSON metadannye (ffprobe): dlitel'nost', kodeki, dorozhki
  audio  <video> [out.wav]               -> izvlech' zvuk 16 kHz mono (ffmpeg) dlja raspoznavanija
  frames <video> <t1,t2,...> [outdir]    -> kadry po tajmkodam (ffmpeg) dlja osmotra video
  transcribe <file> [opcii]              -> stenogramma cherez WhisperX -> .txt / .srt / .json

Opcii transcribe:
  --model large-v3   --lang ru   --device auto|cuda|cpu   --batch 16   --out <dir|prefix>
  --diarize          (razmetka govorjashhih; trebuet HF token)
  --speakers N | --min N --max N          (chislo govorjashhih, esli izvestno)
  --hf-token <tok>   (ili peremennaja okruzhenija HF_TOKEN)

Trebuet ffmpeg na PATH. Dlja transcribe: whisperx + torch. Dlja --diarize: HF token
(modeli pyannote skachivajutsja odin raz, dal'she oflajn - zapis' ne pokidaet mashinu).
"""
import sys
import os
import json
import subprocess
import argparse


def _utf8():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


def _fmt_ts(sec, srt=False):
    sec = float(sec or 0.0)
    h = int(sec // 3600)
    m = int((sec % 3600) // 60)
    s = int(sec % 60)
    ms = int(round((sec - int(sec)) * 1000))
    if ms > 999:
        ms = 999
    if srt:
        return "%02d:%02d:%02d,%03d" % (h, m, s, ms)
    return "%02d:%02d:%02d" % (h, m, s)


def cmd_info(args):
    path = args[0]
    out = subprocess.run(
        ["ffprobe", "-v", "quiet", "-print_format", "json",
         "-show_format", "-show_streams", path],
        capture_output=True)  # БЕЗ text=True: ffprobe отдаёт UTF-8, а text декодирует в cp1251 и теряет кириллицу в пути
    raw = (out.stdout or b"").decode("utf-8", "replace")
    if not raw.strip():
        sys.stderr.write("ffprobe не вернул данные (код %s). Проверь путь к файлу и наличие ffprobe на PATH.\n" % out.returncode)
        sys.exit(2)
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        sys.stderr.write("ffprobe вернул нераспознаваемый вывод (не JSON).\n")
        sys.exit(2)
    fmt = data.get("format", {}) or {}
    streams = data.get("streams", []) or []
    a = next((s for s in streams if s.get("codec_type") == "audio"), None)
    v = next((s for s in streams if s.get("codec_type") == "video"), None)
    info = {
        "path": path,
        "duration_sec": round(float(fmt.get("duration", 0) or 0), 1),
        "container": fmt.get("format_name"),
        "size_mb": round(int(fmt.get("size", 0) or 0) / (1024 * 1024), 1),
        "creation_time": (fmt.get("tags", {}) or {}).get("creation_time"),
        "has_audio": a is not None,
        "audio_codec": (a or {}).get("codec_name"),
        "audio_channels": (a or {}).get("channels"),
        "sample_rate": (a or {}).get("sample_rate"),
        "has_video": v is not None,
        "video_codec": (v or {}).get("codec_name"),
        "width": (v or {}).get("width"),
        "height": (v or {}).get("height"),
    }
    print(json.dumps(info, ensure_ascii=True, indent=2))


def cmd_audio(args):
    video = args[0]
    if len(args) < 2:
        raise ValueError("Укажите путь выходного WAV в папке проекта")
    out = args[1]
    r = subprocess.run(["ffmpeg", "-y", "-i", video, "-vn", "-ac", "1", "-ar", "16000", out],
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    if r.returncode != 0:
        sys.stderr.write((r.stderr or "")[-2000:])
        sys.exit(r.returncode)
    print(out)


def cmd_frames(args):
    video = args[0]
    times = args[1].split(",")
    if len(args) < 3:
        raise ValueError("Укажите папку проекта для кадров")
    outdir = args[2]
    os.makedirs(outdir, exist_ok=True)
    for i, t in enumerate(times, 1):
        t = t.strip()
        fp = os.path.join(outdir, "frame_%02d_%s.jpg" % (i, t.replace(":", "-").replace(".", "_")))
        result = subprocess.run(["ffmpeg", "-y", "-ss", t, "-i", video, "-frames:v", "1", "-q:v", "3", fp],
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
        if result.returncode or not os.path.isfile(fp):
            raise RuntimeError("Кадр не получен: " + result.stderr[-1000:])
        print(fp)


def speaker_segments(segments, turns):
    """Split aligned text by word speakers; do not attribute mixed unaligned text to one voice."""
    output = []
    for segment in segments:
        words = segment.get('words') or []
        if words and all('start' in w and 'end' in w for w in words):
            groups = []
            for word in words:
                speaker = word.get('speaker')
                if not groups or groups[-1].get('speaker') != speaker:
                    groups.append({'start': word['start'], 'end': word['end'],
                                   'speaker': speaker, 'text': word.get('word', ''), 'words': [word]})
                else:
                    groups[-1]['end'] = word['end']
                    groups[-1]['text'] += ' ' + word.get('word', '')
                    groups[-1]['words'].append(word)
            output.extend(groups)
        else:
            speakers = {t['speaker'] for t in turns if min(t['end'], segment['end']) - max(t['start'], segment['start']) > 0.25}
            item = dict(segment)
            if len(speakers) > 1:
                item['speaker'] = None
                item['speaker_warning'] = 'Внутри фрагмента несколько голосов; принадлежность слов требует сверки.'
            output.append(item)
    return output


def cmd_transcribe(args):
    p = argparse.ArgumentParser(prog="av_ingest.py transcribe")
    p.add_argument("file")
    p.add_argument("--model", default="large-v3")
    p.add_argument("--lang", default="ru")
    p.add_argument("--device", default="auto")
    p.add_argument("--batch", type=int, default=16)
    p.add_argument("--diarize", action="store_true")
    p.add_argument("--speakers", type=int, default=None)
    p.add_argument("--min", type=int, default=None, dest="min_sp")
    p.add_argument("--max", type=int, default=None, dest="max_sp")
    p.add_argument("--out", required=True, help="Папка или префикс результата внутри проекта")
    p.add_argument("--hf-token", default=None, dest="hf_token")
    p.add_argument("--no-align", action="store_true", dest="no_align",
                   help="skip word-level alignment (faster, no ~1GB ru align model)")
    a = p.parse_args(args)
    token = a.hf_token or os.environ.get("HF_TOKEN") or os.environ.get("HUGGINGFACE_TOKEN")
    if a.diarize and not token and sys.platform == 'win32':
        from setup_hf_access import saved_token
        token = saved_token()
    if a.diarize and not token:
        p.error("Разделение говорящих требует HF_TOKEN и доступа к модели pyannote")

    # robust model downloads on slow/flaky links (HF default read timeout is only 10s)
    os.environ.setdefault("HF_HUB_DOWNLOAD_TIMEOUT", "120")
    os.environ.setdefault("HF_HUB_DISABLE_SYMLINKS_WARNING", "1")
    import torch
    # ensure ctranslate2 (faster-whisper GPU backend) finds torch's bundled cuDNN/cuBLAS (Windows)
    try:
        _libdir = os.path.join(os.path.dirname(torch.__file__), "lib")
        if os.path.isdir(_libdir):
            try:
                os.add_dll_directory(_libdir)
            except Exception:
                pass
            os.environ["PATH"] = _libdir + os.pathsep + os.environ.get("PATH", "")
    except Exception:
        pass

    import whisperx

    device = a.device
    if device == "auto":
        device = "cuda" if torch.cuda.is_available() else "cpu"
    compute_type = "float16" if device == "cuda" else "int8"
    sys.stderr.write("device=%s compute=%s model=%s lang=%s\n" % (device, compute_type, a.model, a.lang))

    try:
        model = whisperx.load_model(a.model, device, compute_type=compute_type, language=a.lang, vad_method="silero")
    except Exception as e:
        if device == "cuda":
            sys.stderr.write("GPU load failed (%s); falling back to CPU.\n" % str(e)[:140])
            device = "cpu"
            compute_type = "int8"
            model = whisperx.load_model(a.model, device, compute_type=compute_type, language=a.lang, vad_method="silero")
        else:
            raise
    audio = whisperx.load_audio(a.file)
    result = model.transcribe(audio, batch_size=a.batch, language=a.lang)

    # word-level alignment (improves timestamps); skip with --no-align
    if not a.no_align:
        try:
            amodel, meta = whisperx.load_align_model(language_code=a.lang, device=device)
            result = whisperx.align(result["segments"], amodel, meta, audio, device,
                                    return_char_alignments=False)
        except Exception as e:
            sys.stderr.write("align skipped: %s\n" % e)

    # speaker diarization
    if a.diarize:
        if not token:
            sys.stderr.write("ERROR: --diarize requires HF token (--hf-token or HF_TOKEN env var).\n")
            sys.exit(4)
        try:
            from whisperx.diarize import DiarizationPipeline
        except Exception:
            from whisperx import DiarizationPipeline
        dia = DiarizationPipeline(token=token, device=device)
        kw = {}
        if a.speakers:
            kw["num_speakers"] = a.speakers
        if a.min_sp:
            kw["min_speakers"] = a.min_sp
        if a.max_sp:
            kw["max_speakers"] = a.max_sp
        dseg = dia(audio, **kw)
        result = whisperx.assign_word_speakers(dseg, result)
        turns = dseg[['start', 'end', 'speaker']].to_dict(orient='records')
        result['speaker_turns'] = turns
        result['segments'] = speaker_segments(result.get('segments', []), turns)

    segs = result.get("segments", []) or []

    base = a.out or os.path.splitext(a.file)[0]
    if a.out and os.path.isdir(a.out):
        base = os.path.join(a.out, os.path.splitext(os.path.basename(a.file))[0])

    # map raw speaker ids (SPEAKER_00...) -> Г1, Г2 ... in order of appearance
    label = {}

    def spk(seg):
        sp = seg.get("speaker")
        if sp is None:
            return None
        if sp not in label:
            label[sp] = "Г%d" % (len(label) + 1)
        return label[sp]

    txt = base + ".stenogramma.txt"
    with open(txt, "w", encoding="utf-8") as f:
        for s in segs:
            who = spk(s)
            head = "[%s - %s]" % (_fmt_ts(s.get("start")), _fmt_ts(s.get("end")))
            line = head + ((" " + who + ":") if who else "") + " " + (s.get("text") or "").strip()
            if s.get('speaker_warning'):
                line += ' [Смена говорящих — сверить по записи]'
            f.write(line + "\n")

    srtp = base + ".srt"
    with open(srtp, "w", encoding="utf-8") as f:
        for i, s in enumerate(segs, 1):
            who = spk(s)
            text = (("[" + who + "] ") if who else "") + (s.get("text") or "").strip()
            if s.get('speaker_warning'):
                text += ' [Смена говорящих — сверить по записи]'
            f.write("%d\n%s --> %s\n%s\n\n" % (
                i, _fmt_ts(s.get("start"), True), _fmt_ts(s.get("end"), True), text))

    jsonp = base + ".json"
    with open(jsonp, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2, default=str)

    print(txt)
    print(srtp)
    print(jsonp)


def main():
    _utf8()
    if len(sys.argv) > 1 and sys.argv[1] in ("-h", "--help", "/?", "help"):
        print(__doc__)
        sys.exit(0)
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(2)
    cmd, rest = sys.argv[1], sys.argv[2:]
    table = {"info": cmd_info, "audio": cmd_audio, "frames": cmd_frames, "transcribe": cmd_transcribe}
    fn = table.get(cmd)
    if not fn:
        sys.stderr.write("Unknown command: %s\n" % cmd)
        print(__doc__)
        sys.exit(2)
    fn(rest)


if __name__ == "__main__":
    main()
