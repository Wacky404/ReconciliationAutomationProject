import googlemaps
import time

API_KEY = open(r"placeholder").read()
gmaps = googlemaps.Client(key=API_KEY)


def get_place_info(location_name):
    try:
        response = gmaps.places()
        results = response.get('results')
        return results
    except Exception as e:
        print(e)
        return None

