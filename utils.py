import requests
import json


class Util:
    
    def __init__(self, key):
        self.key = key
    
    
    def google_places_extract_query(self, query):
        """Convert to route format."""
        if isinstance(query, str):
            query = {'input': query,
                    'inputtype': 'textquery',
                    'fields': 'formatted_address,name,geometry',
                    'key': self.key}
            return query
        else:
            raise TypeError


    def google_places_search(self,arg):
        route = 'https://maps.googleapis.com/maps/api/place/findplacefromtext/json'

        return requests.get(route, params=self.google_places_extract_query(arg)).json()


    def extract_google_place(self, place):
        if place['status'] != 'OK':
            raise OSError(f'Google places request was bad.\n{json.dumps(place, indent=2)}')
        candidates = place['candidates']
        if len(candidates) == 0:
            raise ValueError(f'No candidate places found. Please refine address search.')
        candidate = candidates[0]
        extract = {
            'lattitude': candidate['geometry']['location']['lat'],
            'longtiude': candidate['geometry']['location']['lng'],
            'name': candidate['name'],
            'address': candidate['address']
        }
        return extract



def nominatim_extract_query(query):
    """Convert to /search route format."""
    if isinstance(query, str):
        query = {'q': query}
    else:
        assert isinstance(query, dict)
    query['format'] = 'json'
    return query
    
    
def nominatim_search(*args):
    """Query Open Street Map's Nominatim API's /search route."""
    route = 'https://nominatim.openstreetmap.org/search'
    responses = [requests.get(route, params=nominatim_extract_query(arg)).json() for arg in args]
    return responses
        