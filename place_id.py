""" This file dictionary houses all the placeIDs with place names """
import pandas as pd
import os.path

file = str('placeidData.csv')
if os.path.isfile(file):
    df_ids = pd.read_csv(file, index_col=0)
else:
    df_ids = pd.DataFrame(columns=['Place ID', 'Place Name'])


def update_place_ids(lst):
    """ Adds the elements of the list into the existing DataFrame and saves """
    try:
        for index, value in enumerate(lst):
            print(index, value)
            df_ids.loc[len(df_ids.index)] = value
        df_ids.to_csv(file)
        lst.clear()
    except Exception as e:
        print(f"An exception of type {type(e).__name__} occurred. Details: {str(e)}")
