""" This file dictionary houses all the placeIDs with place names """
import pandas as pd
import os.path
import random

file = str('placeidData.csv')
if os.path.isfile(file):
    df_ids = pd.read_csv(file)
else:
    df_ids = pd.DataFrame(columns=['Place ID', 'Place Name'])


def update_place_ids(lst):
    """ Adds the elements of the list into the existing DataFrame and saves """
    for index, value in enumerate(lst):
        print(index, value)
        df_ids.loc[len(df_ids.index)] = value
    df_ids.to_csv(file)
    lst.clear()


""" 
not using this saving here, for DataFile.py 
def store_place_id(key_id, name_place):
    lst.append([key_id, name_place])
"""

# Works
if __name__ == "__main__":
    random_places = ['UALR', 'Depo', 'Wally', 'UAMS', 'Stormlight']
    store_place_id = []
    i = 1
    while i < 20:
        key = random.randint(0, 100)
        name = random.choice(random_places)
        store_place_id.append([key, name])
        i += 1
    update_place_ids(store_place_id)
    print(df_ids)
    print(store_place_id)
