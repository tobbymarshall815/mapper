import requests
from pygeodesy.ellipsoidalVincenty import LatLon
import xlwt
import os
import platform

# ---- Created by Michael Bishop -----
# ---- LinkedIn: Michael Bishop -----
# ---- GitHub: https://github.com/tobbymarshall815 -----
# ---- v1.3 BETA----

# Prepare link for GET() request
# i used opencagedata.com API
api_url_1 = 'https://api.opencagedata.com/geocode/v1/json?q='
api_url_2 = '&key=55f3c34bb9a3424d96a72154deca11ea&no_annotations=1&language=en'

# HERE YOU CAN INPUT YOUR CITIES
citylist = ["moscow", 'kiev', 'tel-aviv', 'london', 'tokyo']

# Creating Workbook and new Sheet with name 'city_rows' using xlwt module
wb = xlwt.Workbook()
ws = wb.add_sheet('city_rows', cell_overwrite_ok=True)

# checking the os you ranning, because the terminal commands we use are not same for Windows and MacOs
os_run = platform.system()
print("You run: " + os_run + " OS\n")

# Function to find city coordinates
# i used opencagedata.com API

def get_coordinates(city):
    # Creating URL with city name inside
    request_url = api_url_1 + city + api_url_2
    # Sending GET() to our URL and saving into response
    response = requests.get(request_url)
    # converting response extension to usual text
    info = response.text
    # Here we will save our coordinates
    coordinates_list = []
    # serching the "lat" and "lng" in json response using find method of list
    lat = info.find('"lat":', 1)
    lng = info.find(',"lng":', 2)
    # adding coordinates to coordinates_list choosing only coordinates(numbers) indexes
    coordinates_list.append(info[lat+6:lat+14])
    coordinates_list.append(info[lng+7:lng+15])
    return coordinates_list

# calculating distance with pygeodesy module
def distance(city1, city2):
    print('processing distance for: ', city1, city2, '\n')
    # saving coordinates to new variables
    coordinates_city1 = get_coordinates(city1)
    coordinates_city2 = get_coordinates(city2)
    # using LatLon to find distance, convert our list of coordinated to float
    start_city = LatLon(float(coordinates_city1[0]), float(coordinates_city1[1]))
    end_city = LatLon(float(coordinates_city2[0]), float(coordinates_city2[1]))
    # using distanceTo method of LatLon to find distance in meters
    dist = start_city.distanceTo(end_city)
    # converting number to float, divide to 1000 because we got number in meters,
    # rounding it and converting to int
    newdist = int(round(float(dist))/1000)
    return newdist

# create our base of excel sheet. Input names of cities on first vertical rows and first horizontal rows
def create_excel_sheet(excelfile, city_list):
    # First we need to create excel_files folder for saving
    # checking if folder already exist and creating if it is not using os module and terminal of MacOs
    if os.path.isdir('../mapper/excel_files'):
        print("Folder already exists")
    else:
        name = os.getcwd()
        os.system("cd " + name)
        os.system('mkdir "excel_files"')
    # write names of cities first in vertical column using for avoid first row
    for i in range(len(city_list)):
        excelfile.write(i + 1, 0, city_list[i])
    for i in range(len(city_list)):
    # write names of cities in first horizontal line using for avoid first row
        excelfile.write(0, i + 1, city_list[i])
    #  saving our result to xl_rec.xls file
    wb.save('../mapper/excel_files/xl_rec.xls')
    print("Table created successfully!\n")

# inserting the distances into table
def insert_items_exl(city_list):
    for i in range(len(city_list)):
        # the script will write distances in first vertical column
        for y in range(i, len(city_list)):
            # goes down every column. every new column for goes from i to len(city_list) to avoid recount
            if city_list[i] == city_list[y]:
                #  if cities have same name, we don'' need to calculate their distance because it is 0
                print("SKIPED: ", city_list[i], city_list[y], '\n')
                # here we input zero to those rows
                ws.write(i + 1, y + 1, 0)
            else:
                # saving distance to distance to dist
                dist = distance(city_list[i], city_list[y])
                # in same time inserting distances in both sides of table, because part of the already calculated
                ws.write(i + 1, y + 1, dist)
                ws.write(y+1, i+1, dist)
                print("i:y progress - ", i+1, y+1)
                print('INSERTED CITIES: ', city_list[i], city_list[y], '\n')
    # saving the table
    wb.save('../mapper/excel_files/xl_rec.xls')
    print("XLS FILE SAVED!")

def finish():
    # finish def to reveal file in Finder using os module and terminal
    # using different commands for Windows and MacOS
    print("Reveal in Finder? y/n")
    keyword = input()
    if keyword == 'y':
        if os_run == "Darwin":
            folder = os.getcwd()
            commnd = 'open ' + folder + '/excel_files'
            os.system(commnd)
        if os_run == "Windows":
            folder = os.getcwd()
            commnd = 'cd ' + folder + '/excel_files'
            os.system(commnd)
            os.system("start .")
    if keyword == 'n':
        print("OKEY")



# def start to start our program, makes user experiense more comfortable :)
def start():
    x = 1
    while x>0:
        # some features that you can choose
        print("Choose the NUMBER")
        print("1. Create XLS file with city list")
        print("2. Find coordinates of city")
        print("3. Find distance between the cities")
        print("4. Exit\n")
        key = int(input())
        if key == 1:
            print("***You can change the cities in 'citylist' on the top***\n")
            # using our defs to create base of sheet and inserting there distances. Using finish() in the end
            create_excel_sheet(ws, citylist)
            insert_items_exl(citylist)
            finish()
        if key == 2:
            # choose to find coordinates of city, you will get list if coordinates
            city = input("Input city name ")
            print(get_coordinates(city))
        if key == 3:
            print("ATTENTION: the algorithm considers with an error of 0.5%")
            # yeap... the algoritm that uses LatLon may cause errors because of radius of Earth planet
            # the radius not same in different places, so they use arithmetic mean
            city1 = input("Input city 1 ")
            city2 = input("Input city 2 ")
            print(distance(city1, city2), "KM\n")
        if key == 4:
            x = 0

start()


# ---- Thanks for reading ----
# ---- I will be glad to talk and listen to comments ---
# ---- Michael Bishop ----
# ---- september 2019 ----