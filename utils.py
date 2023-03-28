import requests
import json
import pandas as pd
import numpy as np
import openpyxl as xl
import folium
from multiset import Multiset

from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from google.colab import drive
from google.colab import auth
from oauth2client.client import GoogleCredentials

BENEFICIARIES_XLSX = 'beneficiaries.xlsx'

class Util:
    
    def __init__(self,
                 google_key,
                 beneficiaries_file_id,
                 meal_options):
        self.gkey = google_key
        self.beneficiaries_file_id = beneficiaries_file_id
        self.meal_options = meal_options
        self.num_api_calls = 0
        
        auth.authenticate_user()
        gauth = GoogleAuth()
        gauth.credentials = GoogleCredentials.get_application_default()
        self.gdrive = GoogleDrive(gauth)
        
        print(f'Downloading beneficiaries file {beneficiaries_file_id}...')
        beneficiaries_file = self.gdrive.CreateFile({'id':beneficiaries_file_id})
        beneficiaries_file.GetContentFile(BENEFICIARIES_XLSX)
        
        print(f'Reading file format...')
        df = pd.read_excel(BENEFICIARIES_XLSX)
        self.nam_c = get_col(df, 'beneficiary name')
        self.adr_c = get_col(df, 'address')
        self.rmk_c = get_col(df, 'remarks')
        self.lon_c = get_col(df, 'longitude')
        self.lat_c = get_col(df, 'lattitude')
        self.gname_c = get_col(df, 'google maps name')
        self.gaddr_c = get_col(df, 'google maps address')
        self.meal_options_c = np.full(len(meal_options), -1, dtype=int)
        for i,meal in enumerate(meal_options):
            col = get_col(df, meal)
            if col is None:
                print(f'ERROR: No matching column for the meal "{meal}" was found. '
                      f'Did you mean {df.columns[get_col(df, meal, threshold=0)]}?')
                continue
            self.meal_options_c[i] = col
            print(f'Meal "{meal}": found in column {get_excel_column_letter(col)}.')
        self.valid = True
        self.check_format_valid()
    
    
    def check_format_valid(self):
        col_names = [
            'Beneficiary Name',
            'Address',
            'Remarks',
            'Longitude',
            'Lattitude',
            'Google Maps Name',
            'Google Maps Address',
        ]
        cols = [
            self.nam_c,
            self.adr_c,
            self.rmk_c,
            self.lon_c,
            self.lat_c,
            self.gname_c,
            self.gaddr_c,
        ]
        for col_name, col in zip(col_names, cols):
            if col is None:
                print(f'ERROR: No matching column for "{col_name}" was found. '
                      'Please check the format of the beneficiaries file.')
                self.valid = False
            else:
                print(f'"{col_name}": found in column {get_excel_column_letter(col)}.')
        if np.min(self.meal_options_c) < 0:
            self.valid = False
    
    
    def update_coords(self):
        if not self.valid:
            return
        
        before_api_calls = self.num_api_calls
        
        df = pd.read_excel(BENEFICIARIES_XLSX)
        wb = xl.load_workbook('beneficiaries.xlsx')
        ws = wb.worksheets[0]
        errors = []
        for row in ws.rows:
            if row[self.adr_c].value is not None and \
                (row[self.lon_c].value is None or row[self.lat_c].value is None):
                input_addr = row[self.adr_c].value
                response = self.google_places_search(input_addr)
                try:
                    gplace = extract_google_place(response)
                    row[self.lon_c].value = gplace['longitude']
                    row[self.lat_c].value = gplace['lattitude']
                    row[self.gname_c].value = gplace['name']
                    row[self.gaddr_c].value = gplace['address']
                except ValueError:
                    errors.append([row[self.nam_c].value, row[self.adr_c].value])
        
        num_api_calls = self.num_api_calls - before_api_calls
        if num_api_calls > 0:
            wb.save(BENEFICIARIES_XLSX)
            beneficiaries_file = self.gdrive.CreateFile({'id': self.beneficiaries_file_id})
            beneficiaries_file.SetContentFile(BENEFICIARIES_XLSX)
            beneficiaries_file.Upload()
                      
        print(f'{num_api_calls} Google Places API calls made.')
        
        if len(errors) > 0:
            print(f'The addresses for the following {len(errors)} beneficiaries '
                  'do not match any known places on Google Maps.\n'
                  'Please modify the addresses appropriately, '
                  'or manually input the coordinates of their addresses.')
            for i,err in enumerate(errors):
                print(f'\t{i+1:3d}.\t{err[0]}: {err[1]}')


    def google_places_extract_query(self, query):
        """Convert to route format."""
        if isinstance(query, str):
            query = {'input': query,
                    'inputtype': 'textquery',
                    'fields': 'formatted_address,name,geometry',
                    'key': self.gkey}
            return query
        else:
            raise TypeError


    def google_places_search(self,arg):
        route = 'https://maps.googleapis.com/maps/api/place/findplacefromtext/json'
        self.num_api_calls += 1
        return requests.get(route, params=self.google_places_extract_query(arg)).json()


def get_xl_col(df: pd.DataFrame, *keys):
    keys = [key.lower() for key in keys]
    for i,col in enumerate(df.columns):
        col = col.lower()
        found = True
        for key in keys:
            if key not in col:
                found = False
                break
        if not found:
            continue
        
        # found
        return i


def get_col(df: pd.DataFrame, key, threshold=0.9):
    sims = string_similarity(key, *df.columns)
    col = np.argmax(sims)
    if sims[col] < threshold:
        return None
    return col


def get_excel_column_letter(index):
    return 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[index]


def extract_google_place(place):
    candidates = place['candidates']
    if place['status'] == '' or len(candidates) == 0:
        raise ValueError(f'No candidate places found. Please refine address search.')
    if place['status'] != 'OK':
        raise OSError(f'Google places request was bad.\n{json.dumps(place, indent=2)}')
    candidate = candidates[0]
    extract = {
        'lattitude': candidate['geometry']['location']['lat'],
        'longitude': candidate['geometry']['location']['lng'],
        'name': candidate['name'],
        'address': candidate['formatted_address']
    }
    return extract


def k_shingles(string, k=3):
    string = string.lower()
    shingles = Multiset()
    for i in range(len(string-k)):
        shingles.add(tuple(string[i:i+k]))
    return shingles


def jaccard_similarity(a: Multiset, b: Multiset):
    return len(a & b) / len(a | b)


def string_similarity(test, *args):
    k = max(len(test), 3)
    test_shingles = k_shingles(test, k)
    sims = np.zeros(len(args))
    for i,arg in enumerate(args):
        sims[i] = jaccard_similarity(test_shingles, k_shingles(arg, k))
    return sims


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
        