from googleplaces import GooglePlaces, types


class GoogleIntegration:
    api_key = open(r"C:\Users\Wayne\Work Stuff\Data Conversion\google.txt").read()
    google_places = GooglePlaces(api_key)

    def get_details(self, google_places=google_places, query=None, kind_of=types.TYPE_UNIVERSITY, location=None):
        # 10 QPS as of 11/15
        query_result = google_places.text_search(query=query,type=kind_of, location=location)
        for place in query_result.places:
            # Returned places from a query are place summaries
            query_place_id = place.place_id
            # Must be called
            place.get_details()
            formatted_address = place.formatted_address
            phone_number_full = place.international_phone_number if place.international_phone_number else place.local_phone_number

            return query_place_id, formatted_address, phone_number_full
