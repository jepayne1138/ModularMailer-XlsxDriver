import unittest
import unittest.mock as mock
from xlsx_driver import xlsx_driver


class MockTest(unittest.TestCase):

    @mock.patch(
        'xlsx_driver.xlsx_driver.openpyxl.Workbook.active',
        new_callable=mock.PropertyMock)
    def test__get_sheet_from_workbook_no_sheet_index(self, mock_active):
        """The active property is called if no sheet index provided.

        In all honestly I'm sure there's a better way of doing this, but I'm
        going to continue to try to muddle through better tests and mocking
        if needed. I wasn't sure how to assert that the active sheet was
        returned if no sheet_index parameter was provided, so I decided to
        mock it and assert that the active() property accessor was called
        exactly one. While this is in my mind a weak test, I think it's will
        do for now as I don't need to retest the actual openpyxl
        functionality, and as I want to by default return the active sheet,
        I think it is most likely that I will have to access the 'active'
        property to do this.
        """
        wb = xlsx_driver.openpyxl.Workbook()
        # A hacky attempt to best emulate opening as read-only
        wb.__read_only = True
        # Test the function and assert that active property was accessed
        xlsx_driver._get_sheet_from_workbook(wb)
        mock_active.assert_called_once_with()

    @mock.patch(
        'xlsx_driver.xlsx_driver.openpyxl.Workbook.worksheets',
        new_callable=mock.PropertyMock)
    def test__get_sheet_from_workbook_given_sheet_index(self, mock_worksheets):
        """Given a sheet index, select from the list of worksheets.

        Again, seems like a weak test but if we are given a sheet index we
        would have to get the instance from the list given by the
        worksheets() property, therefore even in the case of possible future
        refactoring, it should be safe to acknowledge the the worksheets()
        property should always be accessed over using the active() property.
        """
        wb = xlsx_driver.openpyxl.Workbook()
        # A hacky attempt to best emulate opening as read-only
        wb.__read_only = True
        # Test the function and assert that worksheets property was accessed
        xlsx_driver._get_sheet_from_workbook(wb, sheet_index=0)
        mock_worksheets.assert_called_once_with()
