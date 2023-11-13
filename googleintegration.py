# Might just be one function, going to placeholder it with a template for now.
import requests
import time


class GoogleIntegration:
    api_key = open(r"placeholder").read()
    url = str("https://places.googlemaps.com/v1/places:search_request")

    def __init__(self, query, result):
        self.query = query
        self.result = result

    @classmethod
    def placeholder(cls, url, query, api_key):
        # Template request to get place_id for location search. Subject to change.
        r = requests.get(url + 'query=' + query + '&key' + api_key )
        result = r.json()
        remove_id = result['results']

        place_id = []
        for i in range(len(remove_id)):
            print(remove_id[i]['names'])
            place_id.append(remove_id)
        return place_id

    @classmethod
    def placeholder_1(cls):

    @classmethod
    def placeholder_2(cls):
