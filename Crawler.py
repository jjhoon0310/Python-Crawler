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
    url = "http://www.google.com/search?q=" + topic + "&source=lmns&bih=722&biw=1536&hl=ko&sa=X&ved=2ahUKEwj8zsz-rsH4AhUE26QKHQExCioQ_AUoAHoECAEQAA"
    if type == "img":
        url ="https://images.google.com/search?q=" + topic + "&tbm=isch&sxsrf=ALiCzsafEkKE1WFqp9nWWIXLSpccWSoB0Q%3A1655910957657&source=hp&biw=1536&bih=722&ei=LTKzYoiRJZDrsAfxzJaABA&iflsig=AJiK0e8AAAAAYrNAPTnG88ayH60tfMErSTNGSwCHvcn2&oq=두햐ㅜㄷ&gs_lcp=CgNpbWcQAxgCMgUIABCABDIFCAAQgAQyBQgAEIAEMgUIABCABDIFCAAQgAQyBQgAEIAEMgUIABCABDIFCAAQgAQyBQgAEIAEMgUIABCABFAAWNoIYManCmgBcAB4AYABuwKIAZcNkgEFMi01LjGYAQCgAQGqAQtnd3Mtd2l6LWltZw&sclient=img"
        img = "rg_i"

elif site == "naver":
    url = "https://www.naver.com/search.naver?where=nexearch&sm=tab_jum&query=" + topic
    if type == "img":
        url ="https://search.naver.com/search.naver?where=image&section=image&query=" + topic
        img = "_image"

else:
    url = "https://www.daum.net/search?w=tot&DA=YZR&t__nil_searchbox=btn&sug=&sugo=&sq=&o=&q=" + topic
    if type == "img":
        url ="https://search.daum.net/search?w=img&nil_search=btn&DA=NTB&enc=utf8&q=" + topic
        img = "thumb_img"

# Opening Browser
driver = webdriver.Chrome("C:\Python\selenium\chromedriver.exe")
driver.implicitly_wait(3)
driver.get(url)

# Input is Doc
if type == "doc":
    elem = driver.find_elements_by_class_name(".LC20lb.MBeuO.DKV0Md")


# Input is excel
elif type == "excel":
    wb = Workbook()
    wb.create_sheet(topic)
    sheet = wb[topic]
    sheet.append(["Title", "URL"])
    wb.save(save_path)

    elem = driver.find_elements_by_class_name()

# Input is image
else:
    imgs = driver.find_elements_by_class_name(img)
    folder_prob(save_path)
    save_imgs(imgs, save_path)

driver.close()
# assert "No results found." not in driver.page_source