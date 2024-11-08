# -*- coding: utf-8 -*-

from .context import excel_workbook

import unittest


class AdvancedTestSuite(unittest.TestCase):
    """Advanced test cases."""

    def test_thoughts(self):
        self.assertIsNone(excel_workbook.hmm())


if __name__ == '__main__':
    unittest.main()
