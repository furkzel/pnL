"""Additional real-Excel regression checks; run after build.py and test_e2e.py.
No original E2E assertions are replaced. Windows/Excel/pywin32 required.
"""
import argparse
import zlib
from pathlib import Path

HARNESS = '''Option Explicit
Public Function ParseNumber(ByVal text As String) As Variant
    Dim value As Double
    If SayiDegeri(text, value) Then
        ParseNumber = value
    Else
        ParseNumber = "REJECTED"
    End If
End Function
Public Function Fingerprint(ByVal ws As Worksheet, ByVal legacy As Boolean) As String
    Fingerprint = AralikImzasi(ws, 1, 1, 5, False, legacy)
End Function
'''


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--panel', required=True, type=Path)
    args = ap.parse_args()
    import xlauto
    with xlauto.ExcelSession() as session:
        wb = session.xl.Workbooks.Open(str(args.panel.resolve()))
        try:
            component = wb.VBProject.VBComponents.Add(1)
            component.Name = 'EnterpriseRegression'
            component.CodeModule.AddFromString(HARNESS)
            ok, messages = xlauto.compile_project(session, wb)
            assert ok, messages
            for raw, expected in [('1.5', 1.5), ('1,5', 1.5), ('1,234.56', 1234.56),
                                  ('1.234,56', 1234.56), ('-12,25', -12.25), ('0', 0),
                                  ('1,234,567', 1234567), ('1.234', 'REJECTED'),
                                  ('1,234', 'REJECTED'), ('1e5', 'REJECTED'),
                                  ('(12)', 'REJECTED'), ('12,34,567', 'REJECTED'),
                                  ('1.23.45', 'REJECTED'), ('', 'REJECTED')]:
                result = session.xl.Run(f"'{wb.Name}'!ParseNumber", raw)
                assert result == expected, (raw, result, expected)
            ws = wb.Worksheets.Add()
            ws.Range('A1:A5').Value = [['first'], [10], ['middle'], [20], ['last']]
            run = lambda old: session.xl.Run(f"'{wb.Name}'!Fingerprint", ws, old)
            original, legacy = run(False), run(True)
            canonical = '8:5:first;5:2:10;8:6:middle;5:2:20;8:4:last;'
            expected_crc = zlib.crc32(canonical.encode('utf-16le'))
            assert original == f'V;CRC32={expected_crc:08X}', (original, expected_crc)
            ws.Range('A3').Value = 'edited'
            assert run(False) != original, 'middle text edit must be detected'
            assert run(True) == legacy, 'legacy algorithm must remain compatible'
            ws.Range('A3').Value = 'middle'
            ws.Range('A2').Value = 11
            ws.Range('A4').Value = 19
            assert run(False) != original, 'sum-preserving edits must be detected'
            ws.Range('A2').Value = 10
            ws.Range('A4').Value = 20
            assert run(False) == original, 'restoring content restores fingerprint'
            ws.Range('A3').Value = 'MIDDLE'
            assert run(False) != original, 'case-sensitive edit must be detected'
            session.run(wb, 'DenetimSessiz')
            audit = wb.Worksheets('Denetim')
            assert isinstance(audit.Cells(6, 2).Value, (int, float)), 'numeric KPI'
            assert audit.PageSetup.FitToPagesWide == 1
            assert audit.ChartObjects.Count >= 1
            session.run(wb, 'KorumalariUygula')
            for name in ('Panel', 'Rapor', 'Gecmis', 'Denetim', 'Profiller', 'Setler'):
                assert wb.Worksheets(name).ProtectContents, name
            for name in ('Gelistirici', 'Listeler1'):
                assert wb.Worksheets(name).Visible == 2, name  # xlSheetVeryHidden
            panel = wb.Worksheets('Panel')
            assert not panel.Cells(10, 3).Locked, 'source picker stays editable'
            assert panel.Cells(10, 1).Locked, 'labels stay protected'
            buttons = [shape for shape in panel.Shapes if shape.Visible and str(shape.Name).startswith('AP_BTN_')]
            assert len(buttons) <= 7, len(buttons)
            assert not panel.Shapes('AP_MENU_SetSil').Visible
            session.run(wb, 'DigerMenu')
            assert panel.Shapes('AP_MENU_SetSil').Visible
            session.run(wb, 'DigerMenu')
            assert not panel.Shapes('AP_MENU_SetSil').Visible
            print('Enterprise Excel regressions passed')
        finally:
            wb.Close(False)


if __name__ == '__main__':
    main()
