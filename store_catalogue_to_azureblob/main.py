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
import crawl_LatestCatalogue
import crawl_SaleFinder
import logging
logging.basicConfig(filename="catalogue_logging.log",
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                    filemode='a')
 
# Creating an object
logger = logging.getLogger()
 
# Setting the threshold of logger to DEBUG
logger.setLevel(logging.INFO)


def main():
    crawl_LatestCatalogue.crawl_LatestCatalogue()
    crawl_SaleFinder.crawl_SaleFinder()
    return None

main()