"""Regression tests for the ADAM-3600 tag generator.

Run from the repository root:
    pip install -r requirements.txt
    python -m unittest discover tests -v
"""
import os
import subprocess
import sys
import tempfile
import unittest

import xlrd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SKILL = os.path.join(ROOT, 'adam3600-tag-import')
SCRIPT = os.path.join(SKILL, 'scripts', 'adam3600_tags.py')
TEMPLATE = os.path.join(SKILL, 'assets', 'adam-3600_tags_template.xls')
HEADER = 'Device,Tag,RegisterType,Address,Format,Length,Access,StartBit,ActiveLow,ScanRate,Description\n'


def run(*args):
    p = subprocess.run([sys.executable, '-I', SCRIPT, *args], capture_output=True, text=True)
    return p.returncode, p.stdout + p.stderr


def cells(path):
    wb = xlrd.open_workbook(path)
    return {s.name: [[(c.ctype, c.value) for c in s.row(r)] for r in range(s.nrows)] for s in wb.sheets()}, wb.sheet_names()


class GeneratorTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.d = self.tmp.name

    def tearDown(self):
        self.tmp.cleanup()

    def csv(self, name, body):
        path = os.path.join(self.d, name)
        with open(path, 'w', newline='') as f:
            f.write(HEADER + body)
        return path

    def test_selftest(self):
        code, out = run('selftest')
        self.assertEqual(code, 0, out)

    def test_template_round_trip_is_identical(self):
        canon = os.path.join(self.d, 'canon.xlsx')
        out = os.path.join(self.d, 'rebuilt.xls')
        self.assertEqual(run('decode', TEMPLATE, '-o', canon)[0], 0)
        code, log = run('build', canon, '--base', '1', '-o', out)
        self.assertEqual(code, 0, log)
        want, want_order = cells(TEMPLATE)
        got, got_order = cells(out)
        self.assertEqual(got_order, want_order)
        self.assertEqual(got, want)

    def test_base0_and_base1(self):
        src = self.csv('b.csv', 'Dev,Coil,CS,0,DISCRETE,,RW,,,1,\nDev,Val,HR,9997,INT32,,R,,,5,\n')
        out = os.path.join(self.d, 'b.xls')
        self.assertEqual(run('build', src, '--base', '0', '-o', out)[0], 0)
        s = xlrd.open_workbook(out).sheet_by_name('Dev')
        self.assertEqual(s.cell_value(1, 4), '00001')
        self.assertEqual(s.cell_value(2, 4), '49998')     # 32 bits -> registers 9998..9999
        self.assertEqual((s.cell_value(2, 19), s.cell_value(2, 20)), ('32', '22'))
        code, log = run('build', src, '--base', '1', '-o', os.path.join(self.d, 'b1.xls'))
        self.assertEqual(code, 1, log)            # address 0 is invalid in base-1
        self.assertIn('outside 1..9999', log)

    def test_errors_write_nothing(self):
        src = self.csv('e.csv', 'VeryLongDeviceName,AndAVeryLongTag,HR,1,INT,16,R,,,1,\n'
                                'Dev,IRrw,IR,1,INT,16,RW,,,1,\n'
                                'Dev,NoScan,HR,2,INT,16,R,,,,\n'
                                'Dev,Frac,HR,3,INT,16,R,,,0.5,\n'
                                'Dev,NoBit,HR,4,DISCRETE,,R,,,1,\n')
        out = os.path.join(self.d, 'e.xls')
        code, log = run('build', src, '--base', '1', '-o', out)
        self.assertEqual(code, 1, log)
        self.assertEqual(log.count('ERROR'), 5, log)
        self.assertFalse(os.path.exists(out))

    def test_confirmations_need_flag(self):
        src = self.csv('c.csv', 'Dev,Float16,HR,1,FLOAT,16,R,,,1,\nDev,Pump Speed,HR,2,UINT,16,RW,,,1,\n')
        out = os.path.join(self.d, 'c.xls')
        code, log = run('build', src, '--base', '1', '-o', out)
        self.assertEqual(code, 2, log)
        self.assertFalse(os.path.exists(out))
        code, log = run('build', src, '--base', '1', '-o', out, '--confirm')
        self.assertEqual(code, 0, log)

    def test_multiple_devices_sheet_order(self):
        src = self.csv('m.csv', 'Meter1,V,HR,1,FLOAT32,,R,,,1,\nFlow2,T,HR,1,DOUBLE_BE,64,R,,,1,\nMeter1,I,HR,3,FLOAT,32,R,,,1,\n')
        out = os.path.join(self.d, 'm.xls')
        self.assertEqual(run('build', src, '--base', '1', '-o', out)[0], 0)
        wb = xlrd.open_workbook(out)
        self.assertEqual(wb.sheet_names(), ['BoardIO', 'Meter1', 'Flow2', 'User Tag', 'Calc Tag'])
        self.assertEqual(wb.sheet_by_name('Meter1').nrows, 3)


if __name__ == '__main__':
    unittest.main()
