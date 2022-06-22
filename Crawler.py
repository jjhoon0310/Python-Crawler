from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from urllib.parse import quote_plus
from urllib.request import urlopen
from openpyxl import Workbook
import time
import os
import datetime

# When there are no folder 
def folder_prob(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print('Error: Creating directory. ' + directory)

# Saving images
def save_imgs(images, save_path):
    for index, image in enumerate(images[:num]):
        src = image.get_attribute('src')
        t = urlopen(src).read()
        file = open(os.path.join(save_path, str(index + 1) + ".jpg"), "wb")
        file.write(t)
        print("img save " + save_path + str(index + 1) + ".jpg")

# User input
site, topic, type, num, save_path = input().split()
num = int(num)

if site == "google":
    site = "http://www.google.com"
    name = "q"
    if type == "img":
        site ="https://images.google.com/"
        img = "rg_i"

elif site == "naver":
    site = "https://www.naver.com/"
    name = "query"
    if type == "img":
        site ="https://search.naver.com/search.naver?where=image&section=image&query="
        img = "_image"

else:
    site = "https://www.daum.net/"
    name = "q"

# Opening Browser
driver = webdriver.Chrome("C:\Python\selenium\chromedriver.exe")
driver.implicitly_wait(3)
driver.get(site)
elem = driver.find_element_by_name(name)
elem.send_keys(topic)
elem.send_keys(Keys.RETURN)

# Input is Doc
if type == "doc":
    elem = driver.find_elements_by_class_name(".LC20lb.MBeuO.DKV0Md")


# Input is excel
elif type == "excel":
    elem = driver.find_elements_by_class_name()

# Input is image
else:
    imgs = driver.find_elements_by_class_name(img)
    folder_prob(save_path)
    save_imgs(imgs, save_path)

driver.close()
# assert "No results found." not in driver.page_source