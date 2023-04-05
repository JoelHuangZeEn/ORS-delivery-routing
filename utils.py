import requests
import json
import colorsys
import pandas as pd
import numpy as np
import openpyxl as xl
import folium
import openrouteservice as ors
from openrouteservice import optimization as opt
from multiset import Multiset

from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from google.colab import drive
from google.colab import auth
from oauth2client.client import GoogleCredentials

BENEFICIARIES_XLSX = 'beneficiaries.xlsx'
DEFAULT_START = {
    'location': [42.23545,-83.73750],
    'name': 'Ann Arbor Meals on Wheels',
}
DEFAULT_MAP = {
    'location': [42.27332,-83.73769],
    'zoom_start': 12
}
WIDTH = 400
HEIGHT = 400
PRECISION = 5
STYLE = \
'''
<head><style>
html * {
    font-family: Calibri;
}
h2 {
    margin-block-end: 0;
}
h3 {
    margin-block-end: 0;
}
</style></head>
'''

class Util:
    
    def __init__(self,
                 google_key,
                 ors_key,
                 beneficiaries_file_id,
                 number_of_vehicles,
                 capacities,
                 time_limit,
                 stop_time,
                 vehicle_start=None):
        self.gkey = google_key
        self.okey = ors_key
        self.beneficiaries_file_id = beneficiaries_file_id
        self.num_vehicles = number_of_vehicles
        self.capacities = capacities
        self.meal_options = capacities.keys()
        self.time_limit = time_limit
        self.stop_time = stop_time
        self.vehicle_start = vehicle_start if vehicle_start is not None else DEFAULT_START
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
        self.meal_options_c = np.full(len(self.meal_options), -1, dtype=int)
        for i,meal in enumerate(self.meal_options):
            col = get_col(df, meal)
            if col is None:
                print(f'{red("ERROR:")} No matching column for the meal option '
                      f'{red(meal)} was found. '
                      f'Did you mean {yellow(df.columns[get_col(df, meal, threshold=0)])}?')
                continue
            self.meal_options_c[i] = col
            print(f'Meal option {green(meal)} found in column '
                  f'{green(get_excel_column_letter(col))}.')
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
                print(f'{red("ERROR:")} No matching column for '
                      f'{red(col_name)} was found. '
                      'Please check the format of the beneficiaries file.')
                self.valid = False
            else:
                print(f'{green(col_name)} found in column '
                      f'{green(get_excel_column_letter(col))}.')
        if np.min(self.meal_options_c) < 0:
            self.valid = False
        if self.valid:
            print(f'Beneficiaries file format is {green("valid.")}')
        else:
            print(f'{red("ERROR:")} Beneficiaries file format is '
                  f'{red("invalid")}. Please correct it to proceed.')
    
    
    def update_coords(self):
        if not self.valid:
            return
        
        before_api_calls = self.num_api_calls
        
        df = pd.read_excel(BENEFICIARIES_XLSX)
        wb = xl.load_workbook(BENEFICIARIES_XLSX)
        ws = wb.worksheets[0]
        errors = []
        print('')
        for row in ws.rows:
            if row[self.adr_c].value is not None and \
                (row[self.lon_c].value is None or row[self.lat_c].value is None):
                input_addr = row[self.adr_c].value
                print(f'\r{yellow(f"Looking up the coordinates for {input_addr}...")}'
                      '                                                 '
                      '                                                 ',
                      end='')
                response = self.google_places_search(input_addr)
                try:
                    gplace = extract_google_place(response)
                    row[self.lon_c].value = gplace['longitude']
                    row[self.lat_c].value = gplace['lattitude']
                    row[self.gname_c].value = gplace['name']
                    row[self.gaddr_c].value = gplace['address']
                except ValueError:
                    errors.append([row[self.nam_c].value, row[self.adr_c].value])
        print('\r                                                               '
              '                                                                 ',
              end='')
        num_api_calls = self.num_api_calls - before_api_calls
        if num_api_calls > 0:
            wb.save(BENEFICIARIES_XLSX)
            beneficiaries_file = self.gdrive.CreateFile({'id': self.beneficiaries_file_id})
            beneficiaries_file.SetContentFile(BENEFICIARIES_XLSX)
            beneficiaries_file.Upload()
                      
        print(f'\r{yellow(num_api_calls)} Google Places API calls made.')
        
        if len(errors) > 0:
            print(f'{red("ERROR:")} The addresses for the following {len(errors)} beneficiaries '
                  'do not match any known places on Google Maps.\n'
                  'Please modify the addresses appropriately, '
                  'or manually input the coordinates of their addresses.')
            for i,err in enumerate(errors):
                print(f'\t{i+1:3d}.\t{err[0]}: {err[1]}')
        
        print('')
        return self.display_beneficiaries()


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
    
    
    def marker_from_row(self, row):
        name = row[self.nam_c]
        addr = row[self.adr_c]
        
        try:
            lat = round(float(row[self.lat_c]),PRECISION)
        except ValueError:
            print(f'{red("ERROR:")} The lattitude of {name}\'s address has a formatting error. '
                  f'[{red(row[self.lat_c])}]')
        
        try:
            lon = round(float(row[self.lon_c]),PRECISION)
        except ValueError:
            print(f'{red("ERROR:")} The longitude of {name}\'s address has a formatting error. '
                  f'[{red(row[self.lon_c])}]')
            
        meals = ''
        for mo,mc in zip(self.meal_options, self.meal_options_c):
            try:
                meals += f'{mo}: {int(row[mc])}<br/>'
            except ValueError:
                print(f'{red("ERROR:")} The number of {mo}s for {name}'
                      ' has a formatting error. '
                      f'[{red(row[mc])}]')
        
        g_name = row[self.gname_c]
        g_addr = row[self.gaddr_c]
        iframe = folium.IFrame(f'{STYLE}'
                               f'<h3>{name}</h3>{addr}'
                               f'<h3>Meals</h3>{meals}'
                               f'<h3>Location Data</h3>'
                               f'Google Maps Name: {g_name}<br/>'
                               f'Google Maps Address: {g_addr}<br/>'
                               f'{lat} °N, {lon} °E',
                               width=WIDTH,
                               height=HEIGHT)
        popup = folium.Popup(iframe, max_width=1000)
        return folium.Marker(location=[lat, lon],
                             popup=popup)
    
    
    def job_from_row(self, row):
        name = row[self.nam_c]
        
        try:
            lat = round(float(row[self.lat_c]),PRECISION)
        except ValueError:
            print(f'{red("ERROR:")} The lattitude of {name}\'s address has a formatting error. '
                  f'[{red(row[self.lat_c])}]')
        
        try:
            lon = round(float(row[self.lon_c]),PRECISION)
        except ValueError:
            print(f'{red("ERROR:")} The longitude of {name}\'s address has a formatting error. '
                  f'[{red(row[self.lon_c])}]')
            
        meals = []
        for mo,mc in zip(self.meal_options, self.meal_options_c):
            try:
                num_meal = int(row[mc])
                meals.append(num_meal)
            except ValueError:
                print(f'{red("ERROR:")} The number of {mo}s for {name}'
                      ' has a formatting error. '
                      f'[{red(row[mc])}]')
                self.valid = False
        return {
            'location': [lon, lat],
            'amount': meals,
            'service': self.stop_time,
        }
    
    
    def display_beneficiaries(self):
        m = folium.Map(**DEFAULT_MAP)
        s_iframe = folium.IFrame(f'{STYLE}'
                                 f'<h3>Vehicle Start</h3>{self.vehicle_start["name"]}<br/>'
                                 f'{self.vehicle_start["location"][0]} °N, '
                                 f'{self.vehicle_start["location"][1]} °E',
                                 width=300,
                                 height=100)
        folium.Marker(location=self.vehicle_start['location'],
                      popup=folium.Popup(s_iframe, max_width=1000),
                      icon=folium.Icon(color='black',
                                       icon='home')).add_to(m)
        df = pd.read_excel(BENEFICIARIES_XLSX)
        for _,row in df.iterrows():
            self.marker_from_row(row).add_to(m)
        return m
    
    
    def route(self):
        if not self.valid:
            return
        
        client = ors.Client(key=self.okey)
        vehicles = [
            opt.Vehicle(id=i, profile='driving-car',
                        start=self.vehicle_start['location'],
                        end=self.vehicle_start['location'],
                        time_window=[0, self.time_limit],
                        capacity=list(self.capacities.values())
                        )
            for i in range(self.num_vehicles)
            ]
        df = pd.read_excel(BENEFICIARIES_XLSX)
        jobs = [opt.Job(id=i, **self.job_from_row(row)) for i, row in df.iterrows() ]
        
        if not self.valid:
            return
        optimized = client.optimization(jobs=jobs, vehicles=vehicles, geometry=True)
        colors = rainbow(self.num_vehicles)
        m = folium.Map(**DEFAULT_START)
        for i,route in enumerate(optimized['routes']):
            folium.PolyLine(locations=[
                list(reversed(coord)) for coord in
                ors.convert.decode_polyline(route['geometry'])['coordinates']],
                color=colors[i]).add_to(m)
            for step in route['steps']:
                if not step['type'] == 'job':
                    continue
                lat = round(float(step['location'][0]),PRECISION)
                lon = round(float(step['location'][1]),PRECISION)
                arrival = int(step['arrival'])
                mask = aeq(df.iloc[:,self.lat_c],lat) & \
                    aeq(df.iloc[:,self.lon_c], lon)
                row = df.loc[mask][0]
                name = row[self.nam_c]
                addr = row[self.adr_c]
                load = step['load']
                meals = ''
                carry = ''
                for mo,mc,lo in zip(self.meal_options, self.meal_options_c, load):
                    meals += f'{mo}: {int(row[mc])}<br/>'
                    carry += f'{mo}: lo<br/>'
                
                hrs = int(arrival//3600)
                mins = int(arrival//60)
                time = ''
                if hrs > 0:
                    time += f'{hrs} hr'
                    time = f'{time}s ' if hrs > 1 else f'{time} '
                if mins > 0:
                    time += f'{mins} min'
                    time = f'{time}s ' if mins > 1 else f'{time} '
                
                iframe = folium.IFrame(f'{STYLE}'
                                       f'<h2>Vehicle {i}</h2>'
                                       f'<h3>{name}</h3>{addr}'
                                       f'<h3>Meals</h3>{meals}<br/>'
                                       f'Arrive in {time}<br/>'
                                       f'Leave carrying<br/>{carry}<br/>'
                                       f'{lat} °N, {lon} °E',
                                       width=WIDTH,
                                       height=HEIGHT)
                popup = folium.Popup(iframe, max_width=1000)
                folium.Marker(location=[lat, lon],
                              popup=popup,
                              icon=folium.Icon(color=colors[i])).add_to(m)
        return m


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
    for i in range(len(string)-k+1):
        shingles.add(tuple(string[i:i+k]))
    return shingles


def jaccard_similarity(a: Multiset, b: Multiset):
    if len(a | b) == 0:
        print(a)
        print(b)
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


def red(string):
    return f'\u001b[31m{string}\u001b[0m'

def green(string):
    return f'\u001b[32m{string}\u001b[0m'

def yellow(string):
    return f'\u001b[33m{string}\u001b[0m'

def blue(string):
    return f'\u001b[34m{string}\u001b[0m'

def rainbow(num):
    hues = np.linspace(0, 0.8, num=num)
    cols = []
    for hue in hues:
        r,g,b = colorsys.hsv_to_rgb(hue, 1, 1)
        cols.append(f'#{r:x}{g:x}{b:x}')
    return cols

def aeq(a,b):
    return (np.abs(a-b) < 10**(-PRECISION) * 0.6)