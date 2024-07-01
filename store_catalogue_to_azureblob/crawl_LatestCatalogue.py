import requests
from bs4 import BeautifulSoup
import json
import os
from datetime import date
import img2pdf
import re
import datetime 
from os.path import exists, join, abspath
from concurrent.futures import ThreadPoolExecutor, as_completed
import logging


cata_path = r'D:\OneDrive - Profectus Group\Ongoing Catalogue\Auto\Tool\catalogues'

# cata_path = r'C:\Users\duongnguyen\OneDrive - Profectus Group\Ongoing Catalogue\Auto\Tool\catalogues'
logging.basicConfig(filename="catalogue_logging.log",
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                    filemode='a')
 
# Creating an object
logger = logging.getLogger()
 
# Setting the threshold of logger to DEBUG
logger.setLevel(logging.INFO)


def find_key(dictionary, value):
    for key in dictionary:
        if dictionary[key].lower() == value:
            return key
    return "Value not found in the dictionary"

def create_folder_if_not_exists(folder_path):
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        print("Folder created: " + folder_path)
    else:
        print("Folder already exists: ", folder_path)

def getLatestCatalogue (url):
    supplier = url.split("/")[3]
    response = requests.get(url)
    soup = BeautifulSoup(response.content, "html.parser")
    links = []
    for link in soup.find_all("a"):
        href = link.get("href")
        if href and supplier in href:
            links.append(href)
    allSuplierLinks = []
    for link in links:
        parts = link.split("/")
        if len(parts) == 4:  # Check if the URL has two parts
            # break 
            allSuplierLinks.append("https://www.latestcatalogues.com" + str(link))
            
    uniqueSuplierLinks = list(set(allSuplierLinks))   
    return uniqueSuplierLinks

def get_image_url(url):
    print("Downloading", url)
    logging.info(f"Downloading: {url} ")
    with requests.Session() as session:
        response = session.get(url)
        if response.status_code == 200:
            soup = BeautifulSoup(response.content, 'html.parser')
            image_element = soup.find(id="pageImage")
            if image_element and 'src' in image_element.attrs:
                return image_element['src']
    return None


def download_image(url, save_path):
    with requests.Session() as session:
        response = session.get(url)
        if response.status_code == 200:
            with open(save_path, 'wb') as file:
                file.write(response.content)
        else:
            print("Failed to download the image: ",url)
            logging.info(f"Failed to download the image: {url} ")

def delete_jpg_files(folder_path1):
    for filename in os.listdir(folder_path1):
        if filename.endswith(".jpg"):
            file_path = os.path.join(folder_path1, filename)
            os.remove(file_path)

def downloadListCataloguesImg(list_download_catalogues,latestcataloguesKeyValue,folder_path,record_list_final):
    with ThreadPoolExecutor() as executor:
        futures = []
        # Save to record file
        for i in range(len(list_download_catalogues)):
            progressStatus = str(int((i + 1) / len(list_download_catalogues) * 100))
            print("Progress status: ", progressStatus + "%")
            print(f"Start catalogue {list_download_catalogues[i]}")
            logging.info(f"Start catalogue {list_download_catalogues[i]}")
            numberPage = getPaginationNumber(list_download_catalogues[i])
            extracted_url = '/'.join(list_download_catalogues[i].split('/')[:4])
            supplierName = find_key(latestcataloguesKeyValue,extracted_url)
            for j in range(1, numberPage):
                imgName = str(j)
                image_url = get_image_url(list_download_catalogues[i] + "?page=" + imgName)
                if image_url:
                    save_location = folder_path + '\\' + imgName + '.jpg'
                    future = executor.submit(download_image, image_url, save_location)
                    futures.append(future)
                else:
                    print("Failed to extract the image URL: ",list_download_catalogues[i] + "?page=" + imgName)

            # Wait for all the submitted tasks to complete
            for future in as_completed(futures):
                futureResult = future.result()

            # List all files in the input directory and filter only JPG images
            image_files = [k for k in os.listdir(folder_path) if k.endswith(".jpg")]
            # Sort image files based on modification time to maintain order
            image_files = sorted(image_files, key=lambda x: int(x.split(".")[0]))
            # Convert the list of images to a single PDF
            pdf_filename = os.path.join(folder_path, supplierName + "_" + formatted_dates + "_" + catalogue + ".pdf")
            if not image_files:
                print("No image files found in the specified directory.")
            else:
                with open(pdf_filename, "wb") as pdf_file:
                    pdf_file.write(img2pdf.convert([os.path.join(folder_path, img) for img in image_files]))
                delete_jpg_files(folder_path)
                print(f"Done catalogue {list_download_catalogues[i]}")
                logging.info(f"Done catalogue {list_download_catalogues[i]}")
                record_list_final.append(list_download_catalogues[i])
                with open("record_list.json", 'w') as f:
                # indent=2 is not needed but makes the file human-readable 
                # if the data is nested
                    json.dump(record_list_final, f, indent=2 , sort_keys=True) 

def getPaginationNumber (url):
    response = requests.get(url)
    if response.status_code == 200:
        soup = BeautifulSoup(response.content, 'html.parser')
        pagination_numbers = soup.find_all("a", class_="page-num")
        last_pagination_number = int(pagination_numbers[-1].text)
        subtitles = soup.find_all(class_="subtitle")
            
        # Extract text content from each subtitle element
        for subtitle in subtitles:
            # Use regular expression to find all occurrences of date in format DD/MM/YYYY
            dates = re.findall(r'\d{2}/\d{2}/\d{4}', subtitle.get_text())

            # Replace "/" with " " in each date and concatenate them into a single string
            global formatted_dates
            formatted_dates = ' '.join(date.replace('/', ' ') for date in dates)

            supplierDate = url.split("/")[4].replace("-", " ")
            split_string = re.split(r'\sfrom\s', supplierDate, maxsplit=1)
        
            # Take the part before "from"
            result = split_string[0]
            
            # Split the result by space to handle the "catalogue" special case
            words = result.split()
            
            # Process words to handle the "catalogue" special case
            processed_words = []
            upper_case_flag = False
            
            for word in words:
                if word.lower() == "catalogue":
                    upper_case_flag = True
                    processed_words.append(word.capitalize())
                elif upper_case_flag:
                    processed_words.append(word.upper())
                else:
                    processed_words.append(word.capitalize())
            
            # Join the processed words back into a single string
            global catalogue
            catalogue = ' '.join(processed_words)
        return last_pagination_number
    else:
        return 0

def crawl_LatestCatalogue():
    # Open the JSON file
    with open('catalogues.json') as file:
        # Load the contents of the file
        catalogues = json.load(file)
    with open('record_list.json') as file:
    # Load the contents of the file
        record_list = json.load(file)
    # Access the data as a Python dictionary
    latestcatalogues = catalogues["latestcatalogues"].values()
    latestcataloguesKeyValue = catalogues["latestcatalogues"]
    
    print("Getting latest catalogues...")
    duplicate_download_catalogues = []    
    for latestcatalogue in latestcatalogues:
        duplicate_download_catalogues = duplicate_download_catalogues + getLatestCatalogue(latestcatalogue)

    list_current_catalogues = [x for x in duplicate_download_catalogues if x not in record_list]

    print("Get latest catalogues successfully!")


    # Generate date string
    today = datetime.date.today()
    monday = today + datetime.timedelta( (7-today.weekday()) % 7)
    monday_str = monday.strftime("%y%m%d")
    monday_month_str = monday.strftime("%m")
    monday_year_str = monday.strftime("%Y")

    # Create date folder
    folder_path = join(cata_path, monday_year_str, monday_month_str, f'{monday_str}')
    create_folder_if_not_exists(folder_path)

    # Open record file
    with open("record_list.json", 'r') as f:
        record_list_final = json.load(f)

    list_download_catalogues = []
    for cata in list_current_catalogues:
        if cata not in record_list_final:
            list_download_catalogues.append(cata)

    downloadListCataloguesImg(list_download_catalogues,latestcataloguesKeyValue,folder_path,record_list_final)

    return None

crawl_LatestCatalogue()