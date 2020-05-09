from unittest import TestCase

import xlrd2

from .base import from_this_dir


class TestIgnoreWorkbookCorruption(TestCase):

    def test_not_corrupted(self):
        with self.assertRaises(Exception) as context:
            xlrd2.open_workbook(from_this_dir('corrupted_error.xls'))
        self.assertTrue('Workbook corruption' in str(context.exception))

        xlrd2.open_workbook(from_this_dir('corrupted_error.xls'), ignore_workbook_corruption=True)
