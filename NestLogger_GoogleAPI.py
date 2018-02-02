#!/usr/bin/python

"""
Uses part of pynest api to read in Nest data, then uses the Google API for spreadsheets to
write that data to a spreadsheet of your choice

pynest Source: https://github.com/smbaker/pynest
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
    #Gets valid user credentials from storage
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join('/home/pi/Nest', '.credentials')
    store = Storage(credential_dir)
    #print(credential_dir)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: #Needed only for compatibilit with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def main():
    
    # Nest Login Data Parser
    parser = create_parser()
    (opts, args) = parser.parse_args() #must keep this unless you want to specify all options from the parser (celsius, serial etc
    opts.user = cfg.NEST_USER #nest email
    opts.password = cfg.NEST_PASS #nest password
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
    # taget_temperature target_humidity leaf current_humidity current_temperature current_schedule_mode
    # temperature_scale auto_away away_temperature_low away_temperature_high
    target_temperature = allvars['target_temperature'] #num
    target_humidity = allvars['target_humidity'] #num
    leaf_status = allvars['leaf'] #True/False
    current_humidity = allvars['current_humidity'] #num
    current_temperature = allvars['current_temperature'] #num
    current_mode = allvars['current_schedule_mode'] #str
    #temperature_scale = allvars['temperature_scale'] #str/char  # Temperature scale displayed on Nest device
    away_status = allvars['auto_away'] #boolean
    away_temperature_low = allvars['away_temperature_low'] #num
    away_temperature_high = allvars['away_temperature_high'] #num
    hvac_heater_state = allvars['hvac_heater_state'] # #True/False
    hvac_ac_state = allvars['hvac_ac_state'] #True/False
    hvac_fan_state = allvars['hvac_fan_state'] #True/False
    cur_time = time.strftime("%c") # Date and time
    cur_date = time.strftime("%a, %b %d, %Y") # Just date
    cur_time2 = time.strftime("%H:%M:%S") # Just time
    
    # Print variables to console
    """print( "time ... ", time)
    print( "target temperature ... ", target_temperature)
    print( "target humidity ... ", target_humidity)
    print( "leaf_status ... ", leaf_status)
    print( "current_humidity ... ", current_humidity)
    print( "current temperature ... ", current_temperature)
    print( "current mode ... ", current_mode)
    #print( "temperature_scale ... ", temperature_scale)
    print( "away_status ... ", away_status)
    print( "away_temperature_low ... ", away_temperature_low)
    print( "away_temperature_high ... ", away_temperature_high)
    print( "hvac_heater_state ... ", hvac_heater_state)
    print( "hvac_ac_state ... ", hvac_ac_state)
    print( "hvac_fan_state ... ", hvac_fan_state)"""
    
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
                 hvac_heater_state, hvac_ac_state, hvac_fan_state] ]
    body = {
        'values' : values
        }
    
    # Write data to the spreadsheet with specified options
    request = service.spreadsheets().values().append(spreadsheetId=spreadsheetId,
                    range=range, valueInputOption=value_input_option,
                    insertDataOption=insert_data_option, body=body)
    response = request.execute()
    #print response

if __name__=="__main__":
   main()
