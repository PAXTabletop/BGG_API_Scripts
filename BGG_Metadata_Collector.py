import csv
import requests   #support for pulling contents of webpages
from bs4 import BeautifulSoup  #function for reading the page's XML returned by requests library
from time import sleep  #sleep function allows pausing of script, to avoid getting rate-limited by BGG
from random import randint  #generate random integers, used in randomizing wait time
import itertools #uses zip function to iterate over two lists concurrently
from pathlib import Path #used in handling Path objects
import html


def data_collect():
    ###### LOAD PAX TITLES ######
    # Read in elements of .csv to different lists, iterating over every row in the PAX Titles csv
    PAX_Titles_path = 'PAXcorrections.csv'
    if Path(PAX_Titles_path).is_file():
        print('Loading PAXcorrections.csv...')
        try:
            PAXgames = open(PAX_Titles_path, 'r', newline='', encoding='utf-8')
        except:
            print('Error loading file. Please load into current working directory and re-run script. Potential .csv type error - ensure UTF-16 encoding')
    else:
        PAX_Titles_path = input('No Game Title Correction export (PAXcorrections.csv) found. Please manually input filename: ')
        if Path(PAX_Titles_path).is_file():
            PAXgames = open(PAX_Titles_path, 'r', newline='', encoding='utf-16')
        else:
            print('No such filename found. Exiting to main menu...')
            sleep(2)
            return
        
    reader = csv.reader(PAXgames)

    #Initialize lists
    PAXnames = []
    PAXids  = []
    BGGids = []
    
    #use next() function to clear the first row in CSV reader, but replace header value with new list of column names for export
    next(reader)
    #header = ['PAX ID', 'Min Player', 'Max Player', 'Year Published', 'Playtime', 'Minimum Age', 'Avg Rating', 'Weight', 'Families','Mechanics','Categories','Description']
    #print(header)
    
    for rows in reader:
        PAXnames.append(rows[0])
        PAXids.append(rows[2])
        BGGids.append(rows[3])


    ###### OPEN NEW CSV FOR WRITING ######
    #Open file for writing, set the writer object, and write the header
    BGGmetadata = open('BGGmetadata.csv', 'w', newline='', encoding='utf-8')
    DataWriter = csv.writer(BGGmetadata, delimiter=',', escapechar='\\', quotechar='"', quoting=csv.QUOTE_NONNUMERIC)
    #DataWriter.writerow(header)

    ###### SET MASTER KEYS FOR TAG DICTIONARIES ######
    categories_full = {}
    families_full = {}
    mechanics_full = {}


    ###### PARSE METADATA FROM BGG API ######

    base_url = 'https://www.boardgamegeek.com/xmlapi2/thing?id='

    #Collect metadata one game at a time, necessary because family/category/mechanics <link> tags appear multiple times for each game and would require regex to parse from larger batch
    for IDs in BGGids:
        if IDs != 0:
            url = base_url + IDs  + '&stats=1'
            print(url)
                                    
            #Use requests and BeautifulSoup to extract and read XML. 
            response = requests.get(url)
            soup = BeautifulSoup(response.text, 'lxml')
            soup_min = soup.find('minplayers') 
            soup_max = soup.find('maxplayers')
            soup_year = soup.find('yearpublished')
            soup_time = soup.find('playingtime')
            soup_age = soup.find('minage')
            soup_rating = soup.find('average')
            soup_weight = soup.find('averageweight')
            soup_desc = soup.find('description').get_text()
            soup_links = soup.find_all('link')


            #Build dictionaries that contain master lists of the three meta tags collected. Key = tag's ID#, value = full text name of that key
            families = []
            mechanics = []
            categories = []

            for links in soup_links:
                if links.attrs['type'] == 'boardgamecategory':
                    categories.append(links.attrs['id'])
                    if links.attrs['id'] not in categories_full.keys():
                        categories_full[links.attrs['id']] = links.attrs['value']
                if links.attrs['type'] == 'boardgamemechanic':
                    mechanics.append(links.attrs['id'])
                    if links.attrs['id'] not in mechanics_full.keys():
                        mechanics_full[links.attrs['id']] = links.attrs['value']
                if links.attrs['type'] == 'boardgamefamily':
                    families.append(links.attrs['id'])
                    if links.attrs['id'] not in families_full.keys():
                        families_full[links.attrs['id']] = links.attrs['value']
                        
            #Extract value from XML tags. When values do not exist, set to 0
            try:
                game_min_player = soup_min.attrs['value']
            except:
                game_min_player = 0
            try:
                game_max_player = soup_max.attrs['value']
            except:
                game_max_player = 0
            try:
                year_published = soup_year.attrs['value']
            except: 
                year_published = 0
            try:
                play_time = soup_time.attrs['value']
            except:
                play_time = 0
            try:
                min_age = soup_age.attrs['value']
            except:
                min_age = 0
            try:
                avg_rating = soup_rating.attrs['value']
            except:
                avg_rating = 0
            try:
                avg_weight = soup_weight.attrs['value']
            except:
                avg_weight = 0

            #Text formatting for the Description paragraph
            desc = BeautifulSoup(soup_desc, 'html5lib')
            desc = str(desc).replace('<html><head></head><body>','')
            desc = str(desc).replace('</body></html>','')
                 
            #Write row to CSV only if game has a BGG ID#. Behavior dependent on PAX_Title_Corrector.py behavior that writes zeros to blank BGG ID# fields
            DataWriter.writerow([PAXids[BGGids.index(IDs)], game_min_player, game_max_player, year_published, play_time, min_age, avg_rating, avg_weight, families, mechanics, categories, desc])
            print(PAXnames[BGGids.index(IDs)])
        
        print('\n' + 'Attempting to load next batch of BGG IDs. Will take 10-15 seconds...' '\n')   
        sleep(randint(10,15))  #sleep to prevent rate-limit

    BGGmetadata.close()

    FamilyWriter = open('families_master.csv', 'w', newline='', encoding='utf-8')
    DataWriter = csv.writer(FamilyWriter, delimiter=',', escapechar='\\', quotechar='"', quoting=csv.QUOTE_NONNUMERIC)
    for key in families_full.keys():
        DataWriter.writerow([key,families_full[key]])
    FamilyWriter.close()

    MechanicsWriter = open('mechanics_master.csv', 'w', newline='', encoding='utf-8')
    DataWriter = csv.writer(MechanicsWriter, delimiter=',', escapechar='\\', quotechar='"', quoting=csv.QUOTE_NONNUMERIC)
    for key in mechanics_full.keys():
        DataWriter.writerow([key,mechanics_full[key]])
    MechanicsWriter.close()

    CategoriesWriter = open('categories_master.csv', 'w', newline='', encoding='utf-8')
    DataWriter = csv.writer(CategoriesWriter, delimiter=',', escapechar='\\', quotechar='"', quoting=csv.QUOTE_NONNUMERIC)
    for key in categories_full.keys():
        DataWriter.writerow([key,categories_full[key]])
    FamilyWriter.close()

    return

if __name__ == "__main__":
    data_collect()