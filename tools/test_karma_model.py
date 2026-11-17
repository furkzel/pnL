"""Model test for M12_Butunluk (v1.1 'C2' fingerprints), no Excel needed.

The functions below replay the VBA statements of Tablo16Kur / Birim /
KarmaDeger with VBA's signed 32-bit Long semantics (integer division truncates
toward zero, (x < 0) evaluates to -1). The result must equal zlib.crc32 of the
documented byte serialization; test_enterprise.py checks the real VBA against
the same reference in Excel.
"""
import struct
import unittest
import zlib


def s32(x):
    x &= 0xFFFFFFFF
    return x - (1 << 32) if x & 0x80000000 else x


def vdiv(a, b):
    q = abs(a) // abs(b)
    return q if (a >= 0) == (b > 0) else -q


def vbool(c):
    return -1 if c else 0


POLY = s32(0xEDB88320)
TABLE = []
for i in range(65536):
    x = i
    for _ in range(16):
        if x & 1:
            x = s32((vdiv(x & 0x7FFFFFFF, 2) | 0x40000000) ^ POLY) if x < 0 else s32(vdiv(x, 2) ^ POLY)
        else:
            x = s32(vdiv(x & 0x7FFFFFFF, 2) | 0x40000000) if x < 0 else vdiv(x, 2)
    TABLE.append(x)


def birim(crc, u):
    x = s32(crc ^ u)
    return s32((vdiv(x & 0x7FFFFFFF, 65536) | (vbool(x < 0) & 0x8000)) ^ TABLE[x & 0xFFFF])


def karma(values):
    crc = s32(0xFFFFFFFF)
    for v in values:
        if v is None:
            crc = birim(crc, 0)
        elif isinstance(v, bool):
            crc = birim(crc, 11)
            crc = birim(crc, 1 if v else 0)
        elif isinstance(v, str):
            n = len(v)
            crc = birim(crc, 8)
            crc = birim(crc, n & 0xFFFF)
            crc = birim(crc, vdiv(n, 65536) & 0xFFFF)
            b = v.encode('utf-16-le')
            for j in range(0, len(b), 2):
                crc = birim(crc, b[j] | (b[j + 1] * 256))
        else:
            crc = birim(crc, 5)
            for unit in struct.unpack('<4h', struct.pack('<d', float(v))):
                crc = birim(crc, unit & 0xFFFF)
    return f'{(~crc) & 0xFFFFFFFF:08X}'


def reference(values):
    out = bytearray()
    for v in values:
        if v is None:
            out += struct.pack('<H', 0)
        elif isinstance(v, bool):
            out += struct.pack('<HH', 11, 1 if v else 0)
        elif isinstance(v, str):
            out += struct.pack('<HHH', 8, len(v) & 0xFFFF, len(v) >> 16) + v.encode('utf-16-le')
        else:
            out += struct.pack('<H', 5) + struct.pack('<d', float(v))
    return f'{zlib.crc32(bytes(out)):08X}'


class KarmaModel(unittest.TestCase):
    def test_matches_zlib(self):
        cases = [[], [None], ['first', 10, 'middle', 20, 'last'], ['0004821', 1250.5, -3.25, 45678.0],
                 ['Şirket ı İ ğ', '=HACK', ''], [True, False, 1e300, -0.0], ['x' * 70000]]
        for values in cases:
            self.assertEqual(karma(values), reference(values), values)

    def test_detects_sum_preserving_and_type_changes(self):
        base = karma(['a', 10, 'b', 20])
        self.assertNotEqual(base, karma(['a', 11, 'b', 19]))
        self.assertNotEqual(base, karma(['a', '10', 'b', 20]))
        self.assertNotEqual(base, karma(['a', 10, 'B', 20]))


if __name__ == '__main__':
    unittest.main()
