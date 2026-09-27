import unittest,io
from openpyxl import Workbook
from logic import compare_schedule_row,compare_material_row
from schedule_import import read_schedule,FIELDS

class ImportTests(unittest.TestCase):
    def test_unchanged_progress_keeps_original_date(self):
        old={'組立進度':'50%','_progress_changed_at':'2026-09-01T08:00:00+08:00'}
        result=compare_schedule_row({'組立進度':'50%'},old,'2026-09-27T08:00:00+08:00')
        self.assertEqual(result['_progress_changed_at'],old['_progress_changed_at'])
        self.assertFalse(result['_last_compare_changed'])
    def test_changed_material_resets_date(self):
        result=compare_material_row({'PU':'齊料'},{'PU':'缺料'},'2026-09-27T08:00:00+08:00')
        self.assertTrue(result['_last_compare_changed'])
        self.assertIn('缺料 → 齊料',result['_material_diff'])
    def test_excel_stable_ids(self):
        wb=Workbook();ws=wb.active;ws.title='2026排程';ws.append(FIELDS)
        ws.append(['26M001','客戶','PN-01','EFEM','A','竹東','王先生','組立中','','2026-09-01','2026-10-01','','2026-10-02'])
        buf=io.BytesIO();wb.save(buf);buf.seek(0)
        a,_=read_schedule(buf);buf.seek(0);b,_=read_schedule(buf)
        self.assertEqual(a[0]['編號'],b[0]['編號'])
        self.assertEqual(a[0]['狀態'],'組立中')
if __name__=='__main__':unittest.main()
