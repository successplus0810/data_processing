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


def saleFinderCatalogues(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.content, "html.parser")

    # Find the first element with class="catalogue-image"
    element = soup.find(class_="catalogue-image")

    # Check if the element exists before accessing its href attribute
    if element is not None:
        href_link = element.get("href")
        return href_link
    else:
        return None

def create_folder_if_not_exists(folder_path):
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        print("Folder created: " + folder_path)
    else:
        print("Folder already exists: ", folder_path)

# folder_path = os.path.join(os.path.expanduser("~"), "Downloads", "salefinder-" + str(current_date))
# create_folder_if_not_exists(folder_path)

def get_image_name(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')
    list_img_name = re.findall(r'"imagefile":"([^"]+)"', str(soup))
    global catalogue_dates
    catalogue_dates = soup.find(class_="sale-dates").get_text()
    return list_img_name

def download_image(url, save_path):
    print("Downloading image: ", url)
    logging.info(f"Downloading image: {url} ")
    with requests.Session() as session:
        response = session.get(url)
        if response.status_code == 200:
            with open(save_path, 'wb') as file:
                file.write(response.content)
        else:
            print("Failed to download the image: ",url)
            # logging.info("Failed to download the image: ",url)
            logging.info(f"Failed to download the image: {url} ")

def delete_jpg_files(folder_path1):
    for filename in os.listdir(folder_path1):
        if filename.endswith(".jpg"):
            file_path = os.path.join(folder_path1, filename)
            os.remove(file_path)

def download_list_salefinder_img(list_download_catalogues,saleFinderKeyValue,folder_path,record_list_final):
    with ThreadPoolExecutor() as executor:
        futures = []
        for catalogues in list_download_catalogues:
            print(f"Start catalogue {catalogues}")
            logging.info(f"Start catalogue {catalogues}")
            extracted_url = '/'.join(catalogues.split('/')[:4])
            supplierName = find_key(saleFinderKeyValue,extracted_url)
            list_img_name = get_image_name(catalogues)
            for j in range(0, len(list_img_name)):
                imgName = str(j)
                typeCatalogue = catalogues.split("/")[4].replace("-", " ").title()
                if supplierName == "IGA Liquor" or supplierName == "Amcal+":
                    image_url = "https://dduhxx0oznf63.cloudfront.net/images/salepages/" + list_img_name[j] 
                else:
                    image_url = "https://d7ldxuhwrlsh5.cloudfront.net/images/salepages/" + list_img_name[j] 
                save_location = folder_path + '\\' + imgName + '.jpg'
                future = executor.submit(download_image, image_url, save_location)
                futures.append(future)
            for future in as_completed(futures):
                futureResult = future.result()
            # List all files in the input directory and filter only JPG images
            image_files = [k for k in os.listdir(folder_path) if k.endswith(".jpg")]
            # Sort image files based on modification time to maintain order
            image_files = sorted(image_files, key=lambda x: int(x.split(".")[0]))
            # Convert the list of images to a single PDF
            pdf_filename = os.path.join(folder_path, supplierName + "_" + catalogue_dates + "_" + typeCatalogue + ".pdf")
            if not image_files:
                print("No image files found in the specified directory.")
                logging.info("No image files found in the specified directory.")
            else:
                with open(pdf_filename, "wb") as pdf_file:
                    pdf_file.write(img2pdf.convert([os.path.join(folder_path, img) for img in image_files]))
                delete_jpg_files(folder_path)
                print(f"Done catalogue {catalogues}")
                logging.info(f"Done catalogue {catalogues}")
                record_list_final.append(catalogues)


def crawl_SaleFinder():
    # Open the JSON file
    with open('catalogues.json') as file:
        # Load the contents of the file
        catalogues = json.load(file)
    with open('record_list.json') as file:
        # Load the contents of the file
        record_list = json.load(file)

    # Access the data as a Python dictionary
    latestcatalogues = catalogues["salefinder"].values()
    saleFinderKeyValue = catalogues["salefinder"]

    print("Getting salefinder catalogues...")
    logging.info("Getting salefinder catalogues...")
    list_current_catalogues = []    
    for latestcatalogue in latestcatalogues:
        checkLink = saleFinderCatalogues(latestcatalogue)
        if checkLink is not None:
            parts = checkLink.split("/")
            first_part = parts[1]
            if parts[1] == "iga-liquor-catalogue" or parts[1] == "amcal-catalogue":
                checkLink = "https://salefinder.com.au" + str(checkLink)
            else:
                checkLink = "https://salefinder.co.nz" + str(checkLink)
            if checkLink[-1] == "2":
                checkLink = checkLink[:-1]
            if checkLink not in record_list:
                list_current_catalogues.append(checkLink)
    print("Get salefinder catalogues successfully!")
    logging.info("Get salefinder catalogues successfully!")
    # Generate date string
    today = datetime.date.today()
    monday = today + datetime.timedelta( (7-today.weekday()) % 7 )
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

    download_list_salefinder_img(list_download_catalogues,saleFinderKeyValue,folder_path,record_list_final)

    # Save to record file
    with open("record_list.json", 'w') as f:
        # indent=2 is not needed but makes the file human-readable 
        # if the data is nested
        json.dump(record_list_final, f, indent=2 , sort_keys=True) 
    return None

crawl_SaleFinder()