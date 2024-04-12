""" This file dictionary houses all the placeIDs with place names """
from log_util import logger
import pandas as pd
import os.path

file = str('placeidData.csv')
if os.path.isfile(file):
    df_ids = pd.read_csv(file, index_col=0)
    logger.debug(f'{file} created dataframe')
else:
    df_ids = pd.DataFrame(columns=['Place ID', 'Place Name'])
    logger.debug(f'new dataframe was created, {file} did not exist')


def update_place_ids(lst: list) -> None:
    """ Adds the elements of the list into the existing DataFrame and saves """
    try:
        for index, value in enumerate(lst):
            df_ids.loc[len(df_ids.index)] = value
        df_ids.to_csv(file)
    except Exception as e:
        logger.exception(f"An exception of type {type(e).__name__} occurred. Details: {str(e)}")
