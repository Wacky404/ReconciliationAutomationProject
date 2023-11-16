import unittest
from DataFile import DataFile


def test_reconcile_google():
    test_df = DataFile(
        r"C:\Users\Wayne\Work Stuff\Data Conversion\Copy Georgia Educational Institutions 2023-06-20.xlsx",
        "All Georgia Institutions", "GA")
    test_df.reconcile_google(test_df.ws_uasys, DataFile.null_values)


if __name__ == '__main__':
    test_reconcile_google()
    print("Everything passed!")
