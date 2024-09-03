from src.DataFile import DataFile, load_workbook
from pandas.testing import assert_frame_equal
from typing import Union

import unittest
import pandas as pd

from os import PathLike

# change these paths
path_to_data: str = r'C:\Users\Wayne\Work Stuff\Data Conversion\DAPIPData (1)\AccreditationData.xlsx'

path_accreditation_data: str = 'data/AccreditationData.xlsx'
path_nces_data: str = 'data/Data_3-14-2023---623.xlsx'
path_testfile: str = 'data/HawaiiTest.xlsx'


class TestTable(unittest.TestCase):
    # desired paths here
    table_path: Union[str, PathLike] = path_to_data
    sheet_name: str = 'InstituteCampuses'
    table_path_prev: Union[str, PathLike] = path_accreditation_data
    sheet_name_prev: str = 'InstituteCampuses'

    def test_tableCreation(self, table_path=table_path, table_path_prev=table_path_prev,
                           sheet_name=sheet_name, sheet_name_prev=sheet_name_prev):
        self.assertNotEqual(str(table_path) == str(table_path_prev),
                            f"Hey! The new table is the still the old one. {table_path}")

        table_new = pd.read_excel(io=table_path, sheet_name=sheet_name)
        table_old = pd.read_excel(io=table_path_prev, sheet_name=sheet_name_prev)

        self.assertIsInstance(table_new, pd.DataFrame)
        self.assertIsInstance(table_old, pd.DataFrame)

    def test_tableCheck(self, table_path=table_path, table_path_prev=table_path_prev,
                        sheet_name=sheet_name, sheet_name_prev=sheet_name_prev):
        table_new = pd.read_excel(io=table_path, sheet_name=sheet_name)
        table_old = pd.read_excel(io=table_path_prev, sheet_name=sheet_name_prev)
        assert_frame_equal(
            left=table_new,
            right=table_old,
            check_dtype=True,
            check_column_type=True,
        )

    def test_reconcileCheck(self, table_path=table_path):
        state: DataFile = DataFile(path_testfile, 'All Hawaii Institutions', 'HI')
        state.wb_data_grab = load_workbook(table_path)
        state.reconcile_institution(state.wb_uasys, state.ws_uasys, state.transf_file, DataFile.ws_data_grab,
                                    DataFile.ws_nces_grab)
        state.reconcile_governing(state.wb_uasys, state.ws_uasys, state.transf_file, state.abbrev,
                                  DataFile.ws_data_grab, DataFile.ws_nces_grab)
        state.reconcile_campuslocation(state.wb_uasys, state.ws_uasys, state.transf_file, state.abbrev,
                                       DataFile.ws_data_grab, DataFile.ws_nces_grab)
        self.assertTrue(True)


if __name__ == '__main__':
    unittest.main()
