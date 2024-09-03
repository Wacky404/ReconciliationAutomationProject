from src.utils.log_util import logger
import requests
import random
import time
import json


class NominatimIntegration:

    url: str = "https://nominatim.openstreetmap.org/search?"
    url_status: str = "https://nominatim.openstreetmap.org/status"
    s = requests.session()
    abbreviations: dict = {
        "Alabama": "AL",
        "Alaska": "AK",
        "Arizona": "AZ",
        "Arkansas": "AR",
        "California": "CA",
        "Colorado": "CO",
        "Connecticut": "CT",
        "Delaware": "DE",
        "Florida": "FL",
        "Georgia": "GA",
        "Hawaii": "HI",
        "Idaho": "ID",
        "Illinois": "IL",
        "Indiana": "IN",
        "Iowa": "IA",
        "Kansas": "KS",
        "Kentucky": "KY",
        "Louisiana": "LA",
        "Maine": "ME",
        "Maryland": "MD",
        "Massachusetts": "MA",
        "Michigan": "MI",
        "Minnesota": "MN",
        "Mississippi": "MS",
        "Missouri": "MO",
        "Montana": "MT",
        "Nebraska": "NE",
        "Nevada": "NV",
        "New Hampshire": "NH",
        "New Jersey": "NJ",
        "New Mexico": "NM",
        "New York": "NY",
        "North Carolina": "NC",
        "North Dakota": "ND",
        "Ohio": "OH",
        "Oklahoma": "OK",
        "Oregon": "OR",
        "Pennsylvania": "PA",
        "Rhode Island": "RI",
        "South Carolina": "SC",
        "South Dakota": "SD",
        "Tennessee": "TN",
        "Texas": "TX",
        "Utah": "UT",
        "Vermont": "VT",
        "Virginia": "VA",
        "Washington": "WA",
        "West Virginia": "WV",
        "Wisconsin": "WI",
        "Wyoming": "WY",
        "District of Columbia": "DC",
        "American Samoa": "AS",
        "Guam": "GU",
        "Northern Mariana Islands": "MP",
        "Puerto Rico": "PR",
        "United States Minor Outlying Islands": "UM",
        "U.S. Virgin Islands": "VI",
    }

    @staticmethod
    def query_structured(amenity=None, street=None, city=None, county=None,
                         state=None, country='USA', postalcode=None, url=url, s=s, server=url_status) -> dict:
        """ queries the Nominatim api in a structured format to limit results """
        arguments = locals()
        params: dict = {}
        for key, value in arguments.items():
            if value is not None and key != 'url':
                params[key] = value
        # additional output params
        params['format'] = 'json'
        params['limit'] = 1
        params['addressdetails'] = 1
        r = random.randint(2, 5)
        time.sleep(r)
        try:
            query_result = s.get(url=url, params=params, timeout=1.5)
            details = json.loads(query_result.text)
            # If no result is found from query then details is empty
            if len(details) >= 1:
                details = dict(details[0])
            else:
                return None

            address = str(details['address']['house_number'] +
                          ' ' + details['address']['road'])
            state_in_details: str = str(details['address']['state'])
            state_abbreviated: str = str()

            if len(state_in_details) > 2 and state_in_details in NominatimIntegration.abbreviations.keys():
                state_abbreviated = str(
                    NominatimIntegration.abbreviations[state_in_details])
            elif len(state_in_details) == 2:
                state_abbreviated = state_in_details

            return {
                'ID': details['place_id'],
                'Name': details['name'],
                'Address': address,
                'Municipality': details['address']['city'],
                'State': state_abbreviated if state_abbreviated != "" else details['address']['state'],
                'ZipCode': details['address']['postcode']
            }

        except requests.exceptions.RequestException as e:
            logger.exception(
                f"An exception of type {type(e).__name__} occurred. Details: {str(e)}")
            time.sleep(r)
            query = s.get(url=server, params={'format': 'json'}, timeout=0.5)
            response = json.loads(query.text)
            logger.debug(f'{response}')
            return None


if __name__ == "__main__":
    result = NominatimIntegration.query_structured(amenity='UALR', street='2801 South University Avenue',
                                                city='Little Rock', county='Pulaski', state='AR',
                                                postalcode='72204')
