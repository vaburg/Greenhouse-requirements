'''
@@@ Heating and lighting integrated model @@@

This script calculates the amount of "Hourly Heating Demand" and "Hourly Electricity Demand(for artificial lighting)"
in a Venlo greenhouse with 1-hectare (10000 m2) area.
This model input automaticly 'hourly temperature' and 'solar radiation' from web site :(https://www.renewables.ninja/).
Outputs of this model are hourly heating demand and hourly electricity demand for lighting in "Excel file".
--> excel file name:"Greenhouse Energy Demand_2021.xlsx"

This model can currently calculate greenhouse energy demand for: Tomatoes, Cucumbers, Lettuce, 
Bell peppers, and Strawberries.
'''


# Import necessary libraries
import requests
import pandas as pd
import json
import numpy as np
import math as math
from openpyxl import Workbook
from io import StringIO


# Create outpput file (Hourly energy demand: Greenhouse Energy Demand_2021.xlsx)
output_file = 'C:/Import VS Local/Results/Greenhouse Energy Demand_2021.xlsx' 
workbook = Workbook()
workbook.save(output_file) 
print("File Created Successfully")


# Input automaticly hourly temperature and hourly solar radiation from (https://www.renewables.ninja/).

# Define API token and request parameters
token = 'fdaf296c3422bf9eb5066ef43d054407027c2ff6'
# Function to fetch and process data from the API
def request_data(token, lat, lon,  date_from, date_to):
    # Base URL for the API
    url = 'https://www.renewables.ninja/api/data/weather?'
    # Initialize a session and set authorization header
    s = requests.session()
    s.headers = {'Authorization': 'Token ' + token}
    # Endpoint for photovoltaic data
    # url = api_base + 'data/pv'

    # Parameters for the API call
    args = {
        'lat': lat,
        'lon': lon,
        'date_from': date_from,
        'date_to': date_to,
        'dataset': 'merra2',
        'format': 'json',
        'local_time': True,
        'var_t2m':True,
        'var_swgdn':True,
    }

    # Make the API request
    response = s.get(url, params=args)
    # Parse the JSON response
    parsed_response = json.loads(response.text)
    # Convert JSON data into a pandas DataFrame
    data = pd.read_json(StringIO(json.dumps(parsed_response['data'])), orient='index')
    # Convert DataFrame to numpy array for processing
    s = np.array(data)
    # Calculate irradiance by summing specific columns
    irradiance = s[:, 2]
    # # Extract temperature data from a specific column
    temperature = s[:, 1]
    # Return processed data
    return temperature, irradiance, len(data)


loction={    # country : [lat, lon]
    'NL':[52.0034, 4.2608],'BE':[50.938, 3.0912],'FR':[48.8341, -3.1331],'CH':[47.283369, 7.710124] }

crop={       # Country : [Crops]
   'NL':['Tomato', 'Cucumber', 'Bell pepper', 'Lettuce', 'Strawberry'],  # Import
   'BE':['Tomato', 'Lettuce', 'Strawberry'],                             # Import
   'FR':['Tomato', 'Lettuce',],                                          # Import
   'CH':['Tomato', 'Cucumber', 'Bell pepper', 'Lettuce', 'Strawberry']   # Local
}


# Lighting plan requrments
PPFD={       # Photosynthetic Photon Flux Density (micromol/m2.s)
    'Tomato': 150, 'Cucumber':200, 'Bell pepper':175, 'Lettuce':100, 'Strawberry':130}

DLI={        # Daily Light Integral (mol/day)
    'Tomato': 30, 'Cucumber':30, 'Bell pepper':30, 'Lettuce':20, 'Strawberry':20}

Min_Naturallight={     # Min natural daily light integral or min solar radiation (W/m2)
    'Tomato': 100, 'Cucumber':130, 'Bell pepper':110, 'Lettuce':100, 'Strawberry':100}


temperature={    # Tempereture setpoints inside greenhouse Day/Night (oC)
    'Tomato':[19,16], 'Cucumber':[20,13], 'Bell pepper':[19,16], 'Lettuce':[14,10], 'Strawberry':[15,5]}

year=['2021']

country=['NL','BE','FR','CH']

for C in country:
    for y in year:
        lat,lon = loction[C]
        date_from = y+'-01-01'
        date_to = y+'-12-31'
        temp, rad, data_length = request_data(token, lat, lon, date_from, date_to)
        for p in crop[C]:
            T_day,T_night=temperature[p]            
            Pmax =PPFD[p]
            Light_Demand=DLI[p] 
            NL_min = Min_Naturallight[p]                               # W/m2, Minimum natural light limit in an hour
            # Greenhouse parameters 
            L=100                                                      # lenght of greenhouse
            W=100                                                      # width of greenhouse
            H=6                                                        # heigh of greenhouse
            G=0.76                                                     # geigh of arch
            teta=23                                                    # angle of arch
            U=3.7                                                      # U-value of cover   - W*m^-2 K , for double glass cover
            Area=L*W                                                   # greenhouse area
            taw = 0.8                                                  # Transparency Coefficient (light)
            Taw_heat=0.75                                              # Transparency Coefficient of covering materials (0.75-0.9)
            Area_cover=(2*(L+W)*H)+(L*W)/math.cos(math.radians(teta))  # greeanhouse cover area
            Volume=(H+(G/2))*L*W                                       # greenhouse Volume
            ACPH=0.5                                                   # air change per hour   1/h
            shading=1                                                  # shading rate

            # Constants
            ro=1.2                         # density of Air    kg/m3
            Cp=1024                        # air heat capacity, J/kg c

            # lighting  Parameters
            s_hour = 5                      # light start time
            e_hour = 19                     # light end time
            PAR = 0.45                      # PAR to Global Horizontal Irradiance (0.35-55)
            EF_LED = 10.8                   # mol/kWh, LED Efficacy (3-7 mol/kWh for HPS and 7-10.8 mol/kWh for LED)


            # Light demand model

            # Reshape the data for hourly radiation to daily

            if data_length==8784:
                Rad = rad.reshape(366 , 24)     # W/m2 for each hour of the year
            else:
                Rad = rad.reshape(365 , 24)     # W/m2 for each hour of the year
            # Calculate ALED_LED per hour for the entire year
            Pmax_1 = Pmax * 3600 / 1000000  # mol/m2.h

            # Initialize ALED_LED_per_hour array
            ALED_LED_per_hour = list()

            # Initialize daily lighting integral variables
            Natural_DLI = np.zeros(np.shape(Rad)[0])
            Artificial_DLI = np.zeros(np.shape(Rad)[0])
            Total_DLI = np.zeros(np.shape(Rad)[0])       # Natural_DLI + Artificial_DLI

            for i in range(np.shape(Rad)[0]):
                for j in range(24):
                    Natural_DLI[i] += Rad[i, j] * 4.6 * PAR * 3600 * taw / 1e6  # mol/day
                    if s_hour <= j <= e_hour and Rad[i, j] <= NL_min and Total_DLI[i] < Light_Demand:
                        AL = Pmax_1         # mol/h
                    else:
                        AL =   0            # mol/h
                    Artificial_DLI[i] += AL
                    Total_DLI[i] = Natural_DLI[i] + Artificial_DLI[i]

                    # Calculate ALED_LED for the hour based on AL and EF_LED  for writing to excel
                    ALED_LED_per_hour.append(AL / EF_LED * Area)   # kWh/hour for the total area of greenhouse
            
            # mol of artificial lighting / mol of natural lighting (for calculation of yield for Strawberry)
            # if p=="Strawberry":                
            #     print('Tota AL/Total NL', C+'_'+p)
            #     print(sum(Artificial_DLI[:])/sum(Natural_DLI[:]))
                

            # Heat demand model
            Qh=[]
            
            for t in range(data_length):

                    if rad[t]==0:
                        Tair=T_night  # Temp set point for night
                        # Tomato: 16, Cucumber:13, Bell pepper:16, Lettuce:10, Strawberry:8 oC
                    else:
                        Tair=T_day  #temp set point for day
                        # Tomato: 19, Cucumber:20, Bell pepper:19, Lettuce:14, Strawberry:17 oC
                    heat_demand= max( 0,(((U*(Tair-temp[t])*Area_cover) +((ACPH*ro*Cp*(Tair-temp[t])*Volume)/3600) - (Taw_heat*(shading)*rad[t]*Area))))/1000
                    Qh.append(float(heat_demand))

            w=pd.DataFrame(np.transpose(np.arange(1,data_length+1,1)),index=None, columns= ['hour'])
            w.insert(1,'Heat demand (kWh/ha)',Qh)
            w.insert(1,'Light demand (kWh/ha)',ALED_LED_per_hour)
            w.insert(1,'Radiation',rad)
            w.insert(1,'temperature',temp)
            # DataFrames w and w2 should already be defined in your code
            
            with pd.ExcelWriter(output_file, engine = 'openpyxl',mode='a') as writer:
                w.to_excel(writer,sheet_name=C+'_'+y+'_'+p)
                # writer.close()
            
            print('done', C+'_'+y+'_'+p)

print('All done')
