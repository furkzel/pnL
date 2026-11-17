"""Mutation tests ensure portable gates reject regressions."""
import tempfile
import unittest
from pathlib import Path
from source_checks import check_sources, check_contrast, SRC


class SourceGates(unittest.TestCase):
    def gate(self, source):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder)
            (path / 'M99_Test.bas').write_text(source, encoding='utf-8')
            return check_sources(path)

    def test_clean(self):
        self.assertEqual(self.gate('Option Explicit\nPublic Sub Test()\nEnd Sub\n'), [])

    def test_ascii(self):
        self.assertTrue(any('ASCII' in x for x in self.gate('Option Explicit\n\'Türkçe')))

    def test_comments_and_strings_are_not_code(self):
        self.assertEqual(self.gate('Option Explicit\n\'RGB(1,2,3)\nPublic Sub Test()\nx = "RGB(1,2,3)"\nEnd Sub'), [])

    def test_color_literal(self):
        self.assertTrue(any('palette' in x for x in self.gate('Option Explicit\nx = RGB(1,2,3)')))

    def test_explicit(self):
        self.assertTrue(any('Option Explicit' in x for x in self.gate("'Option Explicit")))

    def test_unlock(self):
        self.assertTrue(any('EkranAc' in x for x in self.gate('Option Explicit\nPublic Sub Test()\nEkranKilitle\nEnd Sub')))

    def test_forbidden_transpose(self):
        self.assertTrue(any('Transpose' in x for x in self.gate('Option Explicit\nx = Application.Transpose(a)')))

    def test_font_size_literal(self):
        self.assertTrue(any('font size' in x for x in self.gate('Option Explicit\nx.Font.Size = 9')))
        self.assertEqual(self.gate('Option Explicit\nPublic Sub T()\nx.Font.Size = YZ_GOVDE\nEnd Sub'), [])

    def test_alias_contrast_resolved(self):
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder)
            text = (SRC / 'M01_Ayarlar.bas').read_text()
            # A role alias pointing to a pale token must fail where it is used as text.
            text = text.replace('Case prTipUyusmazYazi: Palet = Palet(prEngel)', 'Case prTipUyusmazYazi: Palet = Palet(prCizgi)')
            (target / 'M01_Ayarlar.bas').write_text(text)
            self.assertTrue(any('prTipUyusmazYazi/' in x for x in check_contrast(target)))

    def test_continuations(self):
        self.assertTrue(any('24' in x for x in self.gate('Option Explicit\n' + 'x = x + _\n' * 25)))

    def test_contrast_mutation(self):
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder)
            text = (SRC / 'M01_Ayarlar.bas').read_text()
            # A pale success foreground must fail against the pale success fill.
            text = text.replace('Case prOk: Palet = RGB(21, 128, 61)', 'Case prOk: Palet = RGB(220, 252, 231)')
            (target / 'M01_Ayarlar.bas').write_text(text)
            self.assertTrue(any('prOk/' in x for x in check_contrast(target)))

    def test_current_sources(self):
        self.assertEqual(check_sources() + check_contrast(), [])


if __name__ == '__main__':
    unittest.main()
