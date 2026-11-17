"""Additional real-Excel regression checks; run after build.py and test_e2e.py.
No original E2E assertions are replaced. Windows/Excel/pywin32 required.
"""
import argparse
import struct
import zlib
from pathlib import Path


def typed_crc(values):
    """Independent reference for M12_Butunluk.KarmaDeger (v1.1 'C2' signatures):
    tag 0 empty | tag 5 + IEEE-754 double | tag 8 + length (2 x u16) + UTF-16LE."""
    out = bytearray()
    for v in values:
        if v is None:
            out += struct.pack('<H', 0)
        elif isinstance(v, str):
            out += struct.pack('<HHH', 8, len(v) & 0xFFFF, len(v) >> 16) + v.encode('utf-16-le')
        else:
            out += struct.pack('<H', 5) + struct.pack('<d', float(v))
    return f'{zlib.crc32(bytes(out)):08X}'

HARNESS = '''Option Explicit
Public Function ParseNumber(ByVal text As String) As Variant
    Dim value As Double
    If SayiDegeri(text, value) Then
        ParseNumber = value
    Else
        ParseNumber = "REJECTED"
    End If
End Function
Public Function Fingerprint(ByVal ws As Worksheet, ByVal surum As String) As String
    Fingerprint = AralikImzasi(ws, 1, 1, 5, False, surum)
End Function
Public Function Karma(ByVal ws As Worksheet) As String
    Karma = HizliKarma(ws.Range("A1:A5").Value2)
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
            run = lambda surum: session.xl.Run(f"'{wb.Name}'!Fingerprint", ws, surum)
            original, dev, legacy = run('C2'), run('CRC32'), run('n')
            expected = typed_crc(['first', 10, 'middle', 20, 'last'])
            assert original == f'V;C2={expected}', (original, expected)
            assert session.xl.Run(f"'{wb.Name}'!Karma", ws) == expected
            canonical = '8:5:first;5:2:10;8:6:middle;5:2:20;8:4:last;'
            expected_dev = zlib.crc32(canonical.encode('utf-16le'))
            assert dev == f'V;CRC32={expected_dev:08X}', (dev, expected_dev)
            assert legacy.startswith('V;n=5;'), legacy
            ws.Range('A3').Value = 'edited'
            assert run('C2') != original, 'middle text edit must be detected'
            assert run('n') == legacy, 'v1.0 algorithm must remain compatible'
            ws.Range('A3').Value = 'middle'
            ws.Range('A2').Value = 11
            ws.Range('A4').Value = 19
            assert run('C2') != original, 'sum-preserving edits must be detected'
            ws.Range('A2').Value = 10
            ws.Range('A4').Value = 20
            assert run('C2') == original, 'restoring content restores fingerprint'
            ws.Range('A3').Value = 'MIDDLE'
            assert run('C2') != original, 'case-sensitive edit must be detected'
            ws.Range('A3').Value = 'middle'
            ws.Range('A2').Value = '10'
            assert run('C2') != original, 'number stored as text must differ from the number'
            session.run(wb, 'DenetimSessiz')
            audit = wb.Worksheets('Denetim')
            assert isinstance(audit.Cells(6, 2).Value, (int, float)), 'numeric KPI'
            assert audit.PageSetup.FitToPagesWide == 1
            assert audit.ChartObjects.Count >= 1
            session.run(wb, 'KorumalariUygula')
            for name in ('Panel', 'Rapor', 'Gecmis', 'Denetim', 'Profiller', 'Setler', 'Ayarlar', 'Analiz'):
                assert wb.Worksheets(name).ProtectContents, name
            for name in ('Gelistirici', 'Listeler1', 'Bellek'):
                assert wb.Worksheets(name).Visible == 2, name  # xlSheetVeryHidden
            panel = wb.Worksheets('Panel')
            assert not panel.Cells(10, 3).Locked, 'source picker stays editable'
            assert panel.Cells(10, 1).Locked, 'labels stay protected'
            buttons = [shape for shape in panel.Shapes if shape.Visible and str(shape.Name).startswith('AP_BTN_')]
            assert len(buttons) <= 8, len(buttons)   # DoD B: <= 8 visible buttons
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
