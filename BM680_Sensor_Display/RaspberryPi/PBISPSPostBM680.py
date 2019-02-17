#!/usr/bin/env python
# -*- coding: utf-8 -*-

import bme680
import time
import requests
import json

import Adafruit_GPIO.SPI as SPI
import Adafruit_SSD1306

from datetime import datetime, timedelta
from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont

###############
# Variables
###############

# Power BI (Streamimg Data Set)
PBI_URL = ''
PBI_headers = {
            'Content-Type' :'application/json'
}

# Office365 Token
O365token_url     = ''
O365client_id     = ''
O365client_secret = ''
O365graph_url     = 'https://graph.microsoft.com/v1.0/'
O365headers = {
            'Content-Type': 'application/x-www-form-urlencoded'
}
O365Body = 'client_id=' + \
        O365client_id + '&' + \
        'client_secret=' + \
        O365client_secret + '&' + \
        'grant_type=client_credentials&resource=https://graph.microsoft.com/'

# SharePoint
SPS_SiteName        = ''
SPS_ListName        = ''
SPS_SearchColumn    = ''
SPS_SearchRecord    = ''

# PostTiming
Post_sec            = 6

###############
# IoT Device BME680
###############
Temp_offcet     = -10
try:
    sensor = bme680.BME680(bme680.I2C_ADDR_PRIMARY)
except IOError:
    sensor = bme680.BME680(bme680.I2C_ADDR_SECONDARY)

sensor.set_humidity_oversample(bme680.OS_2X)
sensor.set_pressure_oversample(bme680.OS_4X)
sensor.set_temperature_oversample(bme680.OS_8X)
sensor.set_filter(bme680.FILTER_SIZE_3)
sensor.set_gas_status(bme680.ENABLE_GAS_MEAS)

print('\n\nInitial reading:')
for name in dir(sensor.data):
    value = getattr(sensor.data, name)

    if not name.startswith('_'):
        print('{}: {}'.format(name, value))

sensor.set_gas_heater_temperature(320)
sensor.set_gas_heater_duration(150)
sensor.select_gas_heater_profile(0)

###############
# IoT Device SSD1306
###############
# Raspberry Pi pin configuration:
RST = None     # on the PiOLED this pin isnt used
# Note the following are only used with SPI:
DC = 23
SPI_PORT = 0
SPI_DEVICE = 0
disp = Adafruit_SSD1306.SSD1306_128_32(rst=RST)
# Initialize library.
disp.begin()
# Clear display.
disp.clear()
disp.display()
# Create blank image for drawing.
# Make sure to create image with mode '1' for 1-bit color.
width = disp.width
height = disp.height
image = Image.new('1', (width, height))
# Get drawing object to draw on image.
draw = ImageDraw.Draw(image)
# Draw a black filled box to clear the image.
draw.rectangle((0,0,width,height), outline=0, fill=0)
# Draw some shapes.
# First define some constants to allow easy resizing of shapes.
padding = -2
top = padding
bottom = height-padding
# Move left to right keeping track of the current x position for drawing shapes.
x = 0
# Load default font.
font = ImageFont.load_default()

class O365:
    def __init__(self):
        # get datetime
        self.jst_now = datetime.now().strftime('%Y/%m/%d %H:%M:%S')
        self.utc_now = datetime.now()+timedelta(hours=-9)

    def timeset(self):
        prm_year = '{:0=4}'.format(self.utc_now.year)
        prm_month = '{:0=2}'.format(self.utc_now.month)
        prm_day = '{:0=2}'.format(self.utc_now.day)
        prm_hour = '{:0=2}'.format(self.utc_now.hour)
        prm_minute = '{:0=2}'.format(self.utc_now.minute)
        prm_second = '{:0=2}'.format(self.utc_now.second)

        self.datetime =  prm_year + "-" \
                +   prm_month + "-" \
                +   prm_day + "T" \
                +   prm_hour + ":" \
                +   prm_minute + ":" \
                +   prm_second + "Z"

    def bm680sensor(self):
        sensor.get_sensor_data()
        sensor.data.heat_stable

        self.temp = sensor.data.temperature + Temp_offcet
        self.pres = sensor.data.pressure
        self.humi = sensor.data.humidity
        self.gas  = sensor.data.gas_resistance

    def dispSSD1306(self):
        draw.rectangle((0,0,width,height), outline=0, fill=0)
        dsipsensor1 = str(self.temp) + "C / " + str(self.humi) + "%"
        dispsensor2 = str(self.pres) + "hPa"
        dispsensor3 = str(self.gas) + "Ohms"
        draw.text((x, top),       str(self.jst_now),  font=font, fill=255)
        draw.text((x, top+8),     str(dsipsensor1), font=font, fill=255)
        draw.text((x, top+16),    str(dispsensor2),  font=font, fill=255)
        draw.text((x, top+24),    str(dispsensor3),  font=font, fill=255)
        disp.image(image)
        disp.display()

    def PostPowerBI(self):
        body = [
            {
                "datetime" : self.datetime,
                "temp"  : float(self.temp),
                "humi"  : float(self.humi),
                "press" : float(self.pres),
                "gas"   : float(self.gas)
            }
        ]
        body_json = json.dumps(body).encode("utf-8")

        try:
            res = requests.post(
                    PBI_URL,
                    data=body_json,
                    headers=PBI_headers
                )
            res.close
        except:
            print('PBI request error')

    def GraphTokenGet(self):
        headers = {
            'Content-Type': 'application/x-www-form-urlencoded'
        }

        body = 'client_id=' + \
            O365client_id + '&' + \
            'client_secret=' + \
            O365client_secret + '&' + \
            'grant_type=client_credentials&resource=https://graph.microsoft.com/'

        try:
            res = requests.post(
                O365token_url,
                data=body,
                headers=headers
            )
            res.close
        except:
            print('Graph Token request error')
        
        resjson             = res.json()
        token_type          = resjson['token_type']
        token               = resjson['access_token']
        self.O365TokenKey   = token_type + ' ' + token

    def GetSharePointListID(self):
        headers = {
                    'Authorization':self.O365TokenKey,
                    'Content-Type' :'application/json'
        }

        # Get Site ID
        SiteGet_URL = 'https://graph.microsoft.com/v1.0/sites?search=' + SPS_SiteName

        try:
            res1 = requests.get(
                SiteGet_URL,
                headers=headers
            )
            res1.close
        except:
            print('Site ID Request error')

        res1json    = res1.json()
        

        self.SiteID = res1json['value'][0]['id']
        print(self.SiteID)

        # Get List ID
        ListGet_URL = "https://graph.microsoft.com/v1.0/sites/" + \
                    self.SiteID + "/lists?$filter=displayName eq '" + SPS_ListName + "'"

        print(ListGet_URL)

        try:
            res2 = requests.get(
                ListGet_URL,
                headers=headers
            )
            res2.close
        except:
            print('List ID Request error')

        res2json    = res2.json()
        print(res2json)
        self.ListID = res2json['value'][0]['id']
        print(self.ListID)


        # Get Record ID
        RecordGet_URL   = "https://graph.microsoft.com/v1.0/sites/" + \
                        self.SiteID + "/lists/" + self.ListID + \
                        "/items?expand=fields(select=Id," + SPS_SearchColumn + \
                        ")&filter=fields/" + SPS_SearchColumn + " eq '" + \
                        SPS_SearchRecord + "'"

        try:
            res3 = requests.get(
                RecordGet_URL,
                headers=headers
            )
            res3.close
        except:
            print('Record ID Request error')

        res3json        = res3.json()
        self.RecordID   = res3json['value'][0]['id']
        print(self.RecordID)
    
    def PatchSharePointValue(self):
        DataPatch_URL   = 'https://graph.microsoft.com/v1.0/sites/' + \
                        self.SiteID + '/lists/' + self.ListID + \
                        '/items/' + self.RecordID
        
        headers = {
                    'Authorization':self.O365TokenKey,
                    'Content-Type' :'application/json'
        }

        body = {
            "fields": {
                "Value1": self.temp,
                "Value2": self.humi,
                "Value3": self.pres,
                "Value4": self.gas
            }
        }

        body_json = json.dumps(body).encode("utf-8")

        try:
            res = requests.patch(
                DataPatch_URL,
                data=body_json,
                headers=headers
            )
            res.close
        except:
            print('Patch Request error')

o365    = O365()
i = 0
post_timing = Post_sec * 10

while True:
    o365.timeset()
    o365.bm680sensor()
    o365.dispSSD1306()
    if i == Post_sec :
        o365.PostPowerBI()
        o365.GraphTokenGet()
        o365.GetSharePointListID()
        o365.PatchSharePointValue()
        
        i = 0
    else:
        i += 1
    
    time.sleep(.1)
