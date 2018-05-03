"""
Uses part of nest-thermostat api to read in Nest data, 
Uses OpenWeatherMap API to gather outside weather information,
then uses the Google API for spreadsheets to
write that data to a spreadsheet of your choice

Nest Thermostat Source: https://github.com/FiloSottile/nest_thermostat
Open Wearther Map API: https://openweathermap.org/
Google API: https://developers.google.com/sheets/api/v3/
"""
#Google API Imports
from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

# Nest imports
import sys
from optparse import OptionParser
from nest_thermostat import Nest

# For TimeStamp
import time

# For api.openweathermap.org
import urllib2 # OWM responds to HTTP GET
import json # OWM gives data as text, used to transform to json

# Config
import nestconfig as cfg

# Needed for Google API
try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
        flags = None

# Global variables for google api
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheets API Nest RPi Logging'      

# Parser for nest
def create_parser():
   parser = OptionParser(usage="nest [options] command [command_options] [command_args]",
        description="Commands: fan temp",
        version="unknown")
   parser.add_option("-u", "--user", dest="user",
                     help="username for nest.com", metavar="USER", default=None)
   parser.add_option("-p", "--password", dest="password",
                     help="password for nest.com", metavar="PASSWORD", default=None)
   parser.add_option("-c", "--celsius", dest="celsius", action="store_true", default=False,
                     help="use celsius instead of farenheit")
   parser.add_option("-s", "--serial", dest="serial", default=None,
                     help="optional, specify serial number of nest thermostat to talk to")
   parser.add_option("-i", "--index", dest="index", default=0, type="int",
                     help="optional, specify index number of nest to talk to")
   return parser

# Get Google API Credentials
def get_credentials():
    # Gets valid user credentials from storage
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join('/home/pi/Nest', '.credentials')
    store = Storage(credential_dir)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibilit with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def main():
    
    ####################################################################
    ## Init Variables
    
    # Nest Variables
    target_temperature = '--'
    target_humidity = '--'
    leaf_status = '--'
    current_humidity = '--'
    current_temperature = '--'
    current_mode = '--'
    temperature_scale = '--'
    away_status = '--'
    away_temperature_low = '--'
    away_temperature_high = '--'
    hvac_heater_state = '--'
    hvac_ac_state = '--'
    hvac_fan_state = '--'
    cur_time = '--'
    cur_date = '--'
    cur_time2 = '--'
    
    # Open Weather Map Variables
    OWM_city = '--'
    OWM_cityid = '--'
    OWM_curTemp = '--'
    OWM_pressure = '--'
    OWM_humidity = '--'
    OWM_windSpd = '--'
    OWM_windDir = '--'
    OWM_weatherType = '--'
    OWM_weatherDesc = '--'
    OWM_clouds = '--'
    OWM_rain = '--'
    OWM_snow = '--'
    
    
    ####################################################################
    ## Gather Nest Data
    
    # Nest Login Data Parser
    parser = create_parser()
    (opts, args) = parser.parse_args() # must keep this unless you want to specify all options from the parser (celsius, serial etc
    opts.user = cfg.NEST_USER # nest email
    opts.password = cfg.NEST_PASS # nest password
    opts.serial = None
    opts.index = 0
    args = []
    units = "F"

    # Nest login, and get info
    n = Nest(opts.user, opts.password, opts.serial, opts.index, units=units)
    n.login()
    n.get_status()
    
    shared = n.status["shared"][n.serial]
    device = n.status["device"][n.serial]
    allvars = shared
    allvars.update(device)

    # Get all variables that you want
    # Data is guarenteed assuming Nest does not update format
    target_temperature = allvars['target_temperature'] #num
    target_humidity = allvars['target_humidity'] #num
    leaf_status = allvars['leaf'] #True/False
    current_humidity = allvars['current_humidity'] #num
    current_temperature = allvars['current_temperature'] #num
    current_mode = allvars['current_schedule_mode'] #str
    #temperature_scale = allvars['temperature_scale'] #str/char  # Temperature scale displayed on Nest device - not data temp scale
    away_status = allvars['auto_away'] #boolean
    away_temperature_low = allvars['away_temperature_low'] #num
    away_temperature_high = allvars['away_temperature_high'] #num
    hvac_heater_state = allvars['hvac_heater_state'] # #True/False
    hvac_ac_state = allvars['hvac_ac_state'] #True/False
    hvac_fan_state = allvars['hvac_fan_state'] #True/False
    cur_time = time.strftime("%c") # Date and time
    cur_date = time.strftime("%a, %b %d, %Y") # Just date
    cur_time2 = time.strftime("%H:%M:%S") # Just time
    
    ####################################################################
    ## Get Open Weather Map Data (outside weather info)
    
    # Setup GET(URL)
    url = "http://api.openweathermap.org/data/2.5/weather?id="
    url2 = "&appid="
    requesturl = ''.join([url, cfg.OWM_cityid, url2, cfg.OWM_ApiKey_Free])
    resp = urllib2.urlopen(requesturl).read()
    resp = json.loads(resp) # Transform to json format
    
    # Parse data into variables
    # Data is not guarenteed (for example: rain and snow)
    # Used try/catch to allow errors or missing data
    try:
        OWM_city = resp['name']
    except:
        print('No City Name Data')
    try:
        OWM_cityid = resp['id']
    except:
        print('No City ID Data')
    try:
        OWM_curTemp = resp['main']['temp'] - 273.15 # Kelvin to Celcius
    except:
        print('No Current Temperature Data')
    try:
        OWM_pressure = resp['main']['pressure'] # hPa
    except:
        print('No Pressure Data')
    try:
        OWM_humidity = resp['main']['humidity'] # Pct
    except:
        print('No Humidity Data')
    try:
        OWM_windSpd = resp['wind']['speed'] * 2.23694 # m/s to mph
    except:
        print('No Wind Speed Data')
    try:
        OWM_windDir = resp['wind']['deg'] # deg
    except:
        print('No Wind Speed Direction Data;')
    try:
        OWM_weatherType = resp['weather'][0]['main']
    except:
        print('No Weather Type Data')
    try:
        OWM_weatherDesc = resp['weather'][0]['description']
    except:
        print('No Weather Description Data')
    try:
        OWM_clouds = resp['clouds']['all']
    except:
        print('No clouds data')
    try:
        OWM_rain = resp['rain']['3h']
    except:
        print('No Rain data')
    try:
        OWM_snow = resp['snow']['3h']
    except:
        print('No Snow Data')
    
    ####################################################################
    ## Write to Google Sheets
    
    # Setup Spreadsheet Info
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http, discoveryServiceUrl = discoveryUrl)
    
    # Spreadsheet ID
    spreadsheetId = cfg.GOOGLE_SHEETS_SPREADSHEETID
    # Sheet to modify
    range = 'Sheet1' # Name of sheet to append rows to; table needs to be formated
    
    # formatting will be RAW
    value_input_option = 'RAW'
    # insert rows rather then overwrite
    insert_data_option = 'INSERT_ROWS'
    
    # Create Array of values
    values = [ [ cur_time, current_temperature, current_humidity, target_temperature, target_humidity,
                  away_temperature_low, away_temperature_high, away_status, current_mode, leaf_status,
                 hvac_heater_state, hvac_ac_state, hvac_fan_state, OWM_city, OWM_cityid, OWM_curTemp,
                 OWM_humidity, OWM_pressure, OWM_windSpd, OWM_windDir, OWM_clouds, OWM_rain, OWM_snow, 
                 OWM_weatherType, OWM_weatherDesc] ]
    body = {
        'values' : values
        }
    
    # Write data to the spreadsheet with specified options
    request = service.spreadsheets().values().append(spreadsheetId=spreadsheetId,
                    range=range, valueInputOption=value_input_option,
                    insertDataOption=insert_data_option, body=body)
    response = request.execute()

if __name__=="__main__":
   main()
