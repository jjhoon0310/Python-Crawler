from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from urllib.parse import quote_plus
from urllib.request import urlopen
from openpyxl import Workbook
from docx import Document
from docx.shared import Inches
import test1
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
    url = "http://www.google.com/search?q=" + topic + "&biw=746&bih=722&tbm=nws&sxsrf=ALiCzsYTcS6ApgPEDHY1n8eHDVcx5eLRvQ%3A1655968306018&source=hp&ei=MRK0YoihPOTHmAXl36SwBA&iflsig=AJiK0e8AAAAAYrQgQkDBh1WBzdR_C9dNxcyVMUqDkUI5&ved=0ahUKEwjI_YzVgsP4AhXkI6YKHeUvCUYQ4dUDCAk&uact=5&oq=ㄱㄱ&gs_lcp=Cgxnd3Mtd2l6LW5ld3MQAzIICAAQgAQQsQMyCwgAEIAEELEDEIMBMggIABCABBCxAzILCAAQgAQQsQMQgwEyCwgAEIAEELEDEIMBMgsIABCABBCxAxCDATIECAAQAzIFCAAQgAQyBQgAEIAEMgUIABCABFAAWKECYKcFaABwAHgAgAGwAYgBvgKSAQMwLjKYAQCgAQE&sclient=gws-wiz-news"
    headline = "WlydOe"
    if type == "img":
        url ="https://images.google.com/search?q=" + topic + "&tbm=isch&sxsrf=ALiCzsafEkKE1WFqp9nWWIXLSpccWSoB0Q%3A1655910957657&source=hp&biw=1536&bih=722&ei=LTKzYoiRJZDrsAfxzJaABA&iflsig=AJiK0e8AAAAAYrNAPTnG88ayH60tfMErSTNGSwCHvcn2&oq=두햐ㅜㄷ&gs_lcp=CgNpbWcQAxgCMgUIABCABDIFCAAQgAQyBQgAEIAEMgUIABCABDIFCAAQgAQyBQgAEIAEMgUIABCABDIFCAAQgAQyBQgAEIAEMgUIABCABFAAWNoIYManCmgBcAB4AYABuwKIAZcNkgEFMi01LjGYAQCgAQGqAQtnd3Mtd2l6LWltZw&sclient=img"
        img = "rg_i"

elif site == "naver":
    url = "https://search.naver.com/search.naver?where=news&sm=tab_jum&query=" + topic
    headline ="news_tit"
    if type == "img":
        url ="https://search.naver.com/search.naver?where=image&section=image&query=" + topic
        img = "_image"

else:
    url = "https://search.daum.net/search?w=news&nil_search=btn&DA=NTB&enc=utf8&cluster=y&cluster_page=1&q=" + topic
    headline = "tit_main"
    if type == "img":
        url ="https://search.daum.net/search?w=img&nil_search=btn&DA=NTB&enc=utf8&q=" + topic
        img = "thumb_img"

# Opening Browser
driver = webdriver.Chrome("C:\Python\selenium\chromedriver.exe")
driver.implicitly_wait(3)
driver.get(url)

# Input is Doc
if type == "doc":
    headlines = driver.find_elements_by_class_name(headline)

    document = Document()
    document.add_heading(topic, 0)
    if site == "google":
        
        for i in headlines[:num]:
            text = driver.find_element_by_css_selector('div[role="heading"][aria-level="3"]').get_attribute('innerText')
            href = i.get_attribute('href')
            document.add_paragraph(text)
            document.add_paragraph(href)   
    else:
        for i in headlines[:num]:
            document.add_paragraph(i.text)
            document.add_paragraph(i.get_attribute('href'))
    document.save(save_path + topic + ".docx")

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

# driver.close()
# google 이준석 doc 3 C:/Python/selenium
# naver 이준석 doc 3 C:/Python/selenium