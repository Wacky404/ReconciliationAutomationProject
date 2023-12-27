# Left Off: Do return statement for method and go back to frontend and configure
# the arguments to function caller.
from pprint import pprint
import requests
import random
import time
import json


class NominatimIntegration:
    url = f"https://nominatim.openstreetmap.org/search?"

    @staticmethod
    def query_structured(amenity=None, street=None, city=None, county=None,
                         state=None, country='USA', postalcode=None, url=url):
        """ queries the Nominatim in a structured format to limit results """
        arguments = locals()
        params = {}
        for key, value in arguments.items():
            if value is not None and key != 'url':
                params[f'{key}'] = value
        # additional output params
        params['format'] = 'json'
        params['limit'] = 1
        params['addressdetails'] = 1
        print(params)
        r = random.randint(2, 5)
        time.sleep(r)
        query_result = requests.get(url=url, params=params)
        print(query_result.url)
        details = json.loads(query_result.text)
        print('Starts Here: ')
        pprint(details)


if __name__ == "__main__":
    NominatimIntegration.query_structured(amenity='UALR', street='2801 South University Avenue', city='Little Rock',
                                          county='Pulaski', state='Arkansas', postalcode='72204')
