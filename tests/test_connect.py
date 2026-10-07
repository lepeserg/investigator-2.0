"""Validate local discovery without touching installed user skills."""
from pathlib import Path
import tempfile
import unittest
import sys
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
import connect_skills


class ConnectTests(unittest.TestCase):
    def setUp(self):
        (ROOT/'.test-tmp').mkdir(exist_ok=True)

    def prepare(self, root):
        for name in connect_skills.NAMES:
            folder=root/name;folder.mkdir()
            (folder/'SKILL.md').write_text(f'---\nname: {name}\ndescription: Synthetic skill\n---\n',encoding='utf-8')

    def test_all_entrypoints_resolve_and_repeat_is_safe(self):
        with tempfile.TemporaryDirectory(dir=ROOT/'.test-tmp') as temp:
            root=Path(temp);self.prepare(root)
            self.assertEqual(connect_skills.connect(root),8)
            for name in connect_skills.NAMES:
                entry=root/'.agents/skills'/name/'SKILL.md'
                self.assertTrue((entry.parent/'../../../'/name/'SKILL.md').resolve().is_file())
            before={p:p.read_bytes() for p in (root/'.agents').rglob('SKILL.md')}
            connect_skills.connect(root)
            self.assertTrue(all(p.read_bytes()==v for p,v in before.items()))

    def test_foreign_skill_blocks_all_writes(self):
        with tempfile.TemporaryDirectory(dir=ROOT/'.test-tmp') as temp:
            root=Path(temp);self.prepare(root)
            foreign=root/'.agents/skills'/connect_skills.NAMES[-1]/'SKILL.md'
            foreign.parent.mkdir(parents=True);foreign.write_text('Unrelated skill',encoding='utf-8')
            with self.assertRaises(FileExistsError):connect_skills.connect(root)
            self.assertEqual(list((root/'.agents').rglob('SKILL.md')),[foreign])


if __name__=='__main__':unittest.main()
