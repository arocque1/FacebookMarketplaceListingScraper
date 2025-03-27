import requests
from PIL import Image
from io import BytesIO
import time
import sys
import os
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import random
###################################################
################### User Input ####################

searchName = "335d"
keywords = ["335d", "335 d", "diesel"]
locations = ["locations",]

# Facebook account email and password
email = "fbemail"
password = "fbPassword"

# Be sure that 2FA is enabled on the sender email so you
# can get an app password. Paste that app password here
senderEmail = "senderEmail"
appPassword = "yourEmailAppPassword"

# Email you want notifications to get sent to
yourEmail = "recieverEmail"
numListings = 20

populateExcel = True

###################################################
###################################################

def clickButton(driver, path):
    # Clicks button with xpath input, path, using input driver
    button = WebDriverWait(driver, 2).until(
        EC.element_to_be_clickable((By.XPATH, path)))

    driver.execute_script("arguments[0].scrollIntoView();", button)

    WebDriverWait(driver, 2).until(EC.visibility_of(button))

    driver.execute_script("arguments[0].click();", button)

def sendText(driver, path, text):
    # Types input text to xpath, path, using input river
    elem = driver.find_element(By.XPATH, path)
    elem.clear()
    elem.send_keys(text)


class listing:
    # Listing class that stores relevant data on listings
    def __init__(self, listingName = "N/A", price = -1, location = "N/A", miles = "N/A", coverpic = "N/A", link = "N/A"):
        self.listingName = listingName
        self.price = price
        self.location = location
        self.miles = miles
        self.coverpic = coverpic
        self.link = link

    def printInfo(self):
        print("Listing Name: " + self.listingName)
        print("Price: " + str(self.price))
        print("Location: " + self.location)
        print("Miles: " + str(self.miles))
        print("CoverPic: " + self.coverpic)
        print("Link: " + self.link)
        print()

    def toArray(self):
        return [str(self.listingName), str(self.price), str(self.location), str(self.miles), str(self.coverpic), str(self.link)]


driver = webdriver.Firefox()
firstRun = True
combos = pd.read_excel("loginXpaths.xlsx", engine='openpyxl')

for location in locations:
    url = "https://www.facebook.com/marketplace/" + (location.lower()).replace(" ","") + "/search?sortBy=creation_time_descend&query=" + searchName + "&exact=false"
    driver.get(url)

    # Extract info from each listing
    letter = -1
    emailXpath = ""
    if firstRun == True:
            # Enter email
        try:
            sendText(driver, "//*[@id='email']", email)
        except:
            for i in range(len(combos["Combos"])):
                try:
                    sendText(driver,combos.at[i,"Combos"],email)
                    emailEntered = True
                    emailXpath = i
                    break
                except:
                    pass

        try:
            # Enter password
            sendText(driver, "//*[@id='pass']", password)
        except:
            for i in range(len(combos["Combos"])):
                if i != emailXpath:
                    try:
                        sendText(driver,combos.at[i,"Combos"],password)
                        break
                    except:
                        pass

        for i in range(len(combos["LoginButton"])):
            try:
                clickButton(driver, combos.at[i,"LoginButton"])
                break
            except:
                pass
        print("done")

    # Click location
    clickButton(driver,"/html/body/div[1]/div/div[1]/div/div[3]/div/div/div[1]/div[1]/div[1]/div/div[3]/div[1]/div[2]/div[3]/div[2]/div[1]/div[1]/div/span")

    # Click radius
    clickButton(driver,"/html/body/div[1]/div/div[1]/div/div[4]/div/div/div[1]/div/div[2]/div/div/div/div[3]/div/div[1]/div[3]/div/div/div/label/div[1]/div/div")

    # Click 500 miles
    clickButton(driver,"/html/body/div[1]/div/div[1]/div/div[4]/div/div/div[1]/div/div[3]/div/div/div[1]/div[1]/div/div/div/div/div[1]/div/div[11]/div[1]")

    # Click Apply
    clickButton(driver,"/html/body/div[1]/div/div[1]/div/div[4]/div/div/div[1]/div/div[2]/div/div/div/div[4]/div/div[2]/div/div/div/div/div/div/div[1]/div/span/span")

    page = driver.page_source

    tempListings = []
    # Finds price of listings
    for i in range(1,numListings):
        indexStart = page.find(
            '"__isMarketplaceListingRenderable":"GroupCommerceProductItem"')
        indexEnd = page.find(
            ',"origin_group":null,"listing_video"')
        temp = page[indexStart:indexEnd]

        tempPrice = temp[temp.find('"formatted_amount":"') + 20: temp.find('","amount_with_offset_in_currency"')]

        tempLocation = temp[temp.find(':{"city":"') + 10: temp.find('","city_page":{"display_name"')]
        tempLocation = tempLocation[0:tempLocation.find('",')] + ", " + tempLocation[len(tempLocation) - 2:]

        tempName = temp[temp.find(',"custom_title":"') + 17:temp.find('","custom_sub_titles_with_rendering_flags":')]

        page = page[indexEnd + 108:]
        tempListing = listing(tempName, tempPrice, tempLocation)
        tempListings.append(tempListing)
        #tempListing.printInfo()

    listings = []
    # Finds and appends listing classes with relevant data webscraped off facebook marketplace
    for i in range(1, numListings):
        titleXpath = "/html/body/div[1]/div/div[1]/div/div[3]/div/div/div[1]/div[1]/div[2]/div/div/div[3]/div/div[2]/div[" + str(i) + "]/div/div/span/div/div/div/div/a/div/div[2]/div[2]/span/div/span/span"
        priceXpath = "/html/body/div[1]/div/div[1]/div/div[3]/div/div/div[1]/div[1]/div[2]/div/div/div[3]/div/div[2]/div[" + str(i) + "]/div/div/span/div/div/div/div/a/div/div[2]/div[1]/span/div/span"
        locationXpath = "/html/body/div[1]/div/div[1]/div/div[3]/div/div/div[1]/div[1]/div[2]/div/div/div[3]/div/div[2]/div[" + str(i) + "]/div/div/span/div/div/div/div/a/div/div[2]/div[3]/span/div/span/span"
        milesXpath = "/html/body/div[1]/div/div[1]/div/div[3]/div/div/div[1]/div[1]/div[2]/div/div/div[3]/div/div[2]/div[" + str(i) + "]/div/div/span/div/div/div/div/a/div/div[2]/div[4]/div/span/span"
        imgXpath = "/html/body/div[1]/div/div[1]/div/div[3]/div/div/div[1]/div[1]/div[2]/div/div/div[3]/div/div[2]/div[" + str(i) + "]/div/div/span/div/div/div/div/a/div/div[1]/div/div/div/div/div/img"
        linkXpath = "/html/body/div[1]/div/div[1]/div/div[3]/div/div/div[1]/div[1]/div[2]/div/div/div[3]/div/div[2]/div[" + str(i) + "]/div/div/span/div/div/div/div/a"

        listingName = driver.find_element(By.XPATH, titleXpath)
        price = driver.find_element(By.XPATH, priceXpath)
        location = driver.find_element(By.XPATH, locationXpath)
        miles = driver.find_element(By.XPATH, milesXpath)
        imgLink = driver.find_element(By.XPATH, imgXpath)
        link = driver.find_element(By.XPATH, linkXpath)

        listingName = listingName.get_attribute("outerHTML")
        listingName = listingName[listingName.find('">')+2:listingName.find("</span>")]
        #print(listingName)

        location = location.get_attribute("outerHTML")
        location = location[location.find('">')+2:location.find("</span>")]
        #print(location)

        miles = miles.get_attribute("outerHTML")
        miles = miles[miles.find('">')+2:miles.find("</span>")]
        #print(miles)

        img = imgLink.get_attribute("src")

        link = link.get_attribute("href")
        #print(link)

        tempListing = listing(listingName, -1, location, miles, img, link)
        listings.append(tempListing)
        #tempListing.printInfo()

    finalListings = []
    # Matches and combines listings with correct price data
    for seleniumListing in listings:
        for scrapeListing in tempListings:
            if scrapeListing.listingName == seleniumListing.listingName and scrapeListing.location == seleniumListing.location:
                tempListing = listing(seleniumListing.listingName, scrapeListing.price, seleniumListing.location, seleniumListing.miles, seleniumListing.coverpic, seleniumListing.link)

                if any(s in tempListing.listingName for s in keywords) or ("i" not in tempListing.listingName and "BMW" in tempListing.listingName):
                    finalListings.append(tempListing)
                break


    # Initial save to excel if running with empty excel spreadsheet
    if populateExcel == True:
        data = []
        for tempListing in finalListings:
            data.append([str(tempListing.listingName), str(tempListing.price), str(tempListing.location), str(tempListing.miles), str(tempListing.coverpic), str(tempListing.link)])

        columns = ['Listing Name', 'Price', 'Location', 'Miles', 'Coverpic', 'link']

        df = pd.DataFrame(columns=columns)

        for row in data:
            df.loc[len(df)] = row

        excel_path = "cars.xlsx"
        df.to_excel(excel_path, index=False, engine='openpyxl')

        print(f"Data saved to " + excel_path + "!")
        sys.exit()



    df = pd.read_excel("cars.xlsx", engine='openpyxl')
    columns = ['Listing Name', 'Price', 'Location', 'Miles', 'Coverpic', 'link']

    # Saves listing data to excel sheet
    if firstRun == True:
        newListings = []
    for tempListing in finalListings:
        for i in range(1,numListings):
            try:
                if tempListing.listingName != df.at[i, 'Listing Name'] and tempListing.price != df.at[i, 'Price'] and tempListing.miles != df.at[i, 'Miles'] and populateExcel == False:
                    newDf = pd.DataFrame(columns=columns)

                    row = tempListing.toArray()
                    newDf.loc[len(newDf)] = row

                    df = pd.concat([newDf, df], ignore_index=True)

                    newListings.append(tempListing)
            except:
                if len(newListings) > 0:
                    newListings.append(tempListing)
    if len(df) > numListings:
        df = df.drop(range(len(df)-numListings,len(df)))
    df.to_excel("cars.xlsx", index=False, engine='openpyxl')
    firstRun = False

# Sends email to user to notify that new listing is posted
if populateExcel == False:
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    message = MIMEMultipart()
    message['From'] = senderEmail
    message['To'] = yourEmail
    message['Subject'] = 'New Listings!'
    message.attach(MIMEText("""
        <h1><span style="text-decoration: underline;"><strong>New Listings!!!</strong></span></h1>""", 'html'))

    imgNames = []
    for tempListing in newListings:
        response = requests.get(tempListing.coverpic)
        if response.status_code == 200:
            img = Image.open(BytesIO(response.content))
            imgPath = "pics/" + tempListing.listingName + ".jpg"
            imgNames.append(imgPath)
            Image.open(imgPath)
            img.save(imgPath)
        else:
            print("Failed to retrieve the image.")

        with open(imgPath, 'rb') as imgFile:
            img = MIMEImage(imgFile.read())
            img.add_header('Content-ID', tempListing.listingName)
            message.attach(img)

        body = f"""
           <p><span style="text-decoration: underline;"><strong><img src="cid:""" + tempListing.listingName + """ alt="" /></strong></span></p>
           <p>""" + tempListing.price + """</p>
           <p>""" + tempListing.listingName + """</p>
           <p>""" + tempListing.location + """</p>
           <p>""" + tempListing.link + """</p>
           <p>&nbsp;</p>
           <p>&nbsp;</p>"""
        message.attach(MIMEText(body, 'html'))

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(email, appPassword)
        server.sendmail(email, yourEmail, message.as_string())
        print("Email sent successfully!")
    except Exception as e:
        print(f"Failed to send email: {e}")
    finally:
        server.quit()


    for imgName in imgNames:
        if os.path.exists(imgName):
            os.remove(imgName)

