# Don't forget to store place_ids and associate them with each location, maybe dictionary with key
# Saved website/github as bookmark
# Could store place_ids in DataFrame using pandas
from googleplaces import GooglePlaces

api_key = open(r"C:\Users\Wayne\Work Stuff\Data Conversion\google.txt").read()
google_places = GooglePlaces(api_key)

query_result = google_places.text_search(query='SOUTHERN UNION STATE COMMUNITY COLLEGE OPELIKA CAMPUS',
                                         type='University', location='OPELIKA, Alabama')
if query_result.html_attributions:
    print(query_result.html_attributions)
# this does everything !!!
for place in query_result.places:
    # Returned places from a query are place summaries.
    test_place_id = place.place_id
    print("This is the variable for place_ID " + test_place_id)
    # Must be called.
    place.get_details()
    formatted_address = place.formatted_address
    print(formatted_address)
    PhoneNumberFull = place.international_phone_number if place.international_phone_number else place.local_phone_number
    print(PhoneNumberFull)


