import requests
import json
import openpyxl as xl
import pandas as pd


from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from google.colab import drive
from google.colab import auth
from oauth2client.client import GoogleCredentials

BENEFICIARIES_XLSX = 'beneficiaries.xlsx'

class Util:
    
    def __init__(self, key, beneficiaries_file_id):
        self.key = key
        self.beneficiaries_file_id = beneficiaries_file_id
        
        auth.authenticate_user()
        gauth = GoogleAuth()
        gauth.credentials = GoogleCredentials.get_application_default()
        self.gdrive = GoogleDrive(gauth)
        
        beneficiaries_file = self.gdrive.CreateFile({'id':beneficiaries_file_id})
        beneficiaries_file.GetContentFile(BENEFICIARIES_XLSX)
    
    
    def update_coords(self):
        df = pd.read_excel(BENEFICIARIES_XLSX)
        wb = xl.load_workbook('beneficiaries.xlsx')
        ws = wb.worksheets[0]
        lon_c = get_xl_col('longitude', df)
        lat_c = get_xl_col('lattitude', df)
        for row in ws.rows:
            print(row[lon_c].value, row[lat_c].value)
    
    
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


def get_xl_col(key, df: pd.DataFrame):
    key = key.lower()
    return [i+1 for i,col in df.columns if key in col.lower()][0]


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
        