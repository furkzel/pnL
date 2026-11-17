"""Mutation tests: vba_lint must reject what Excel's compiler rejects."""
import tempfile
import unittest
from pathlib import Path

from vba_lint import lint, SRC

A = '''Attribute VB_Name = "MA"
Option Explicit
Public Const K_SAYI As Long = 3
Public Sub Islem(ByVal a As Long, ByVal b As String, Optional ByVal c As Long = 0)
End Sub
Private Sub Gizli()
End Sub
Public Function Topla(ByVal x As Long) As Long
    Topla = x + K_SAYI
End Function
'''


class Lint(unittest.TestCase):
    def run_lint(self, body_b: str, extra_a: str = ""):
        with tempfile.TemporaryDirectory() as d:
            p = Path(d)
            (p / "MA.bas").write_text(A + extra_a)
            (p / "MB.bas").write_text('Attribute VB_Name = "MB"\nOption Explicit\n'
                                      "Public Sub Test()\n    Dim n As Long\n" + body_b + "\nEnd Sub\n")
            return lint(p, builtins={"msgbox", "len"})

    def test_clean(self):
        self.assertEqual(self.run_lint('    Islem 1, "x"\n    n = Topla(2) + Len("a")\n    Call Islem(1, "y", 3)'), [])

    def test_undeclared(self):
        self.assertTrue(any("undeclared identifier 'm'" in x for x in self.run_lint("    m = 1")))

    def test_too_many_args(self):
        self.assertTrue(any("takes 3" in x for x in self.run_lint('    Islem 1, "x", 2, 4')))

    def test_missing_arg(self):
        self.assertTrue(any("missing argument" in x for x in self.run_lint("    n = Topla()")))

    def test_paren_statement(self):
        self.assertTrue(any("statement" in x for x in self.run_lint('    Islem(1, "x")')))

    def test_private_cross_module(self):
        self.assertTrue(any("Private" in x for x in self.run_lint("    Gizli")))

    def test_duplicate_dim(self):
        self.assertTrue(any("duplicate declaration" in x for x in self.run_lint("    Dim n As String")))

    def test_ambiguous(self):
        self.assertTrue(any("ambiguous" in x for x in self.run_lint("", "Public Sub Test()\nEnd Sub\n")))

    def test_named_args(self):
        self.assertEqual(self.run_lint('    Islem b:="x", a:=1'), [])
        self.assertTrue(any("no parameter" in x for x in self.run_lint('    Islem 1, "x", z:=1')))

    def test_current_sources(self):
        self.assertEqual(lint(SRC), [])


if __name__ == "__main__":
    unittest.main()
