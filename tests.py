from ExcelFile import ExcelFile
from os.path import isfile
from os import remove
from datetime import datetime

import unittest


class MyTestCase(unittest.TestCase):
    def test_create_excel_file(self):
        file_path = 'test/test' + str(datetime.now().timestamp()) + '.xlsx'
        e1 = ExcelFile(file_path)
        e1.create_empty_excel_file()
        self.assertEqual(isfile(file_path), True)
        remove(file_path)

    def test_call(self):
        file_path = 'unexisting/test.xlsx'
        with (self.assertRaises(FileNotFoundError)):
            e1 = ExcelFile(file_path)

    def test_load_excel_file(self):
        e1 = ExcelFile('test/test.xlsx')
        e1.load_excel_file()
        my_cell_value = e1.excel_file['Feuil1']['A1'].value
        self.assertEqual(my_cell_value, 'test')

    def test_write_in_cell_of_sheet(self):
        file_path = 'test/test' + str(datetime.now().timestamp()) + '.xlsx'
        e1 = ExcelFile(file_path)
        e1.write_in_cell_of_sheet('Sheet', 1, 1, 'test')
        self.assertEqual(e1.excel_file['Sheet']['A1'].value, 'test')
        remove(file_path)

    def test_write_in_cell_of_sheet_already_existing_file(self):
        file_path = 'test/test_write_already_existing_file.xlsx'
        e1 = ExcelFile(file_path)
        e1.write_in_cell_of_sheet('Sheet', 1, 1, 'test')
        self.assertEqual(e1.excel_file['Sheet']['A1'].value, 'test')
        self.assertEqual(e1.excel_file['Sheet']['G8'].value, 'hello')
        e1.write_in_cell_of_sheet('Sheet', 1, 1, '')

    def test_write_list_in_line_of_sheet(self):
        file_path = 'test/test' + str(datetime.now().timestamp()) + '.xlsx'
        e1 = ExcelFile(file_path)
        e1.write_list_in_line_of_sheet('Sheet', 3, ['test3', 'test2', 'test3'])
        self.assertEqual(e1.excel_file['Sheet']['A3'].value, 'test3')
        self.assertEqual(e1.excel_file['Sheet']['B3'].value, 'test2')
        self.assertEqual(e1.excel_file['Sheet']['C3'].value, 'test3')

        e1.write_list_in_line_of_sheet('Sheet', 7, ['test3', 'test2', 'test3'], start_column=5)
        self.assertEqual(e1.excel_file['Sheet']['E7'].value, 'test3')
        self.assertEqual(e1.excel_file['Sheet']['F7'].value, 'test2')
        self.assertEqual(e1.excel_file['Sheet']['G7'].value, 'test3')

        remove(file_path)


if __name__ == '__main__':
    unittest.main()
