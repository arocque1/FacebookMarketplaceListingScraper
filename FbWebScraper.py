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

searchName = "search name"
# Required keywords are keywords that MUST be in the listing title
requiredKeywords = ["keywords"]

# Optional keywords are used to select listings that may not have a
# required keyword, but do have an optional keyword and DON'T have an exclusion keyword
optionalKeywords = ["optional keywords"]

# Exclusion keywords are keywords that listing titles must NOT have
exclusionKeywords = ["exclusion keywords"]

locations = ["locations"]

# Facebook account email and password
email = "email"
password = "pass"

# Email that will send you emails notifying you of new listings.
# Be sure that 2FA is enabled on the sender email so you
# can get an app password. Paste that app password here
senderEmail = "email"
appPassword = "app pass"

# Email you want notifications to get sent to
yourEmail = "email"
numListings = 10

###################################################

def clickButton(driver, path):
    button = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, path)))

    driver.execute_script("arguments[0].scrollIntoView();", button)

    WebDriverWait(driver, 2).until(EC.visibility_of(button))

    driver.execute_script("arguments[0].click();", button)

def sendText(driver, path, text):
    elem = driver.find_element(By.XPATH, path)
    # Simulates human typing
    for c in text:
        mistake = random.randint(0, 100)
        elem.send_keys(c)
        time.sleep(random.uniform(0.05,0.2))
        if mistake < 10:
            mistakeNum = random.randint(1, 3)
            for i in range(mistakeNum):
                elem.send_keys(chr(random.randint(97,122)))
                time.sleep(random.uniform(0.05,0.1))
            time.sleep(random.uniform(0.15,0.5))
            for i in range(mistakeNum):
                elem.send_keys(Keys.BACKSPACE)
                time.sleep(random.uniform(0.01,0.1))




class listing:
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

    def __eq__(self, other):
        return self.listingName == other.listingName and self.price == other.price and self.location == other.location and self.miles == other.miles and self.coverpic == other.coverpic and self.link == other.link


driver = webdriver.Firefox()
firstRun = True
combos = pd.read_excel("loginXpaths.xlsx", engine='openpyxl')
for location in locations:
    url = "https://www.facebook.com/marketplace/" + (location.lower()).replace(" ","") + "/search?sortBy=creation_time_descend&query=" + searchName + "&exact=false"
    driver.get(url)

    #Extract info from each listing
    letter = -1
    emailXpath = ""
    if firstRun == True:
            #Enter email
        try:
            sendText(driver, "//*[@id='email']", email)
        except:
            for i in range(len(combos["LoginFields"])):
                try:
                    sendText(driver,combos.at[i,"LoginFields"],email)
                    emailEntered = True
                    emailXpath = i
                    break
                except:
                    pass

        try:
            #Enter password
            sendText(driver, "//*[@id='pass']", password)
        except:
            for i in range(len(combos["LoginFields"])):
                if i != emailXpath:
                    try:
                        sendText(driver,combos.at[i,"LoginFields"],password)
                        break
                    except:
                        pass

        for i in range(len(combos["LoginButton"])):
            try:
                clickButton(driver, combos.at[i,"LoginButton"])
                break
            except:
                pass

    #Click location
    for i in range(len(combos["RangeButton"])):
        try:
            clickButton(driver, combos.at[i,"RangeButton"])
        except:
            pass
    #clickButton(driver,"/html/body/div[1]/div/div[1]/div/div[3]/div/div/div[1]/div[1]/div[1]/div/div[3]/div[1]/div[2]/div[3]/div[2]/div[1]/div[1]/div/span")

    #Click radius
    clickButton(driver,"/html/body/div[1]/div/div[1]/div/div[4]/div/div/div[1]/div/div[2]/div/div/div/div[3]/div/div[1]/div[3]/div/div/div/label/div[1]/div/div")

    #Click 500 miles
    clickButton(driver,"/html/body/div[1]/div/div[1]/div/div[4]/div/div/div[1]/div/div[3]/div/div/div[1]/div[1]/div/div/div/div/div[1]/div/div[11]/div[1]")

    #Click Apply
    clickButton(driver,"/html/body/div[1]/div/div[1]/div/div[4]/div/div/div[1]/div/div[2]/div/div/div/div[4]/div/div[2]/div/div/div/div/div/div/div[1]/div/span/span")

    page = driver.page_source

    tempListings = []
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
    for i in range(1, numListings):
        titleXpath = "/html/body/div[1]/div/div[1]/div/div[3]/div/div/div[1]/div[1]/div[2]/div/div/div[3]/div/div[2]/div[" + str(i) + "]/div/div/span/div/div/div/div/a/div/div[2]/div[2]/span/div/span/span"
        priceXpath = "/html/body/div[1]/div/div[1]/div/div[3]/div/div/div[1]/div[1]/div[2]/div/div/div[3]/div/div[2]/div[" + str(i) + "]/div/div/span/div/div/div/div/a/div/div[2]/div[1]/span/div/span"
        locationXpath = "/html/body/div[1]/div/div[1]/div/div[3]/div/div/div[1]/div[1]/div[2]/div/div/div[3]/div/div[2]/div[" + str(i) + "]/div/div/span/div/div/div/div/a/div/div[2]/div[3]/span/div/span/span"
        milesXpath = "/html/body/div[1]/div/div[1]/div/div[3]/div/div/div[1]/div[1]/div[2]/div/div/div[3]/div/div[2]/div[" + str(i) + "]/div/div/span/div/div/div/div/a/div/div[2]/div[4]/div/span/span"
        imgXpath = "/html/body/div[1]/div/div[1]/div/div[3]/div/div/div[1]/div[1]/div[2]/div/div/div[3]/div/div[2]/div[" + str(i) + "]/div/div/span/div/div/div/div/a/div/div[1]/div/div/div/div/div/img"
        linkXpath = "/html/body/div[1]/div/div[1]/div/div[3]/div/div/div[1]/div[1]/div[2]/div/div/div[3]/div/div[2]/div[" + str(i) + "]/div/div/span/div/div/div/div/a"

        listingName = driver.find_element(By.XPATH, titleXpath)
        price = driver.find_element(By.XPATH, priceXpath)
        try:
            location = driver.find_element(By.XPATH, locationXpath)
        except:
            continue
        try:
            miles = driver.find_element(By.XPATH, milesXpath)
        except:
            continue
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

        wait = WebDriverWait(driver, 2)
        imgLink = wait.until(EC.presence_of_element_located((By.XPATH, imgXpath)))
        img = imgLink.get_attribute("src")
        #print(img)

        wait = WebDriverWait(driver, 2)
        link = wait.until(EC.presence_of_element_located((By.XPATH, linkXpath)))
        link = link.get_attribute("href")
        #print(link)

        tempListing = listing(listingName, -1, location, miles, img, link)
        listings.append(tempListing)
        #tempListing.printInfo()

    finalListings = []
    for seleniumListing in listings:
        for scrapeListing in tempListings:
            if scrapeListing.listingName == seleniumListing.listingName and scrapeListing.location == seleniumListing.location:
                tempListing = listing(seleniumListing.listingName, scrapeListing.price, seleniumListing.location, seleniumListing.miles, seleniumListing.coverpic, seleniumListing.link)

                if any(s in tempListing.listingName for s in requiredKeywords) or (not any(t in tempListing.listingName for t in exclusionKeywords) and any(u in tempListing.listingName for u in optionalKeywords)):
                    finalListings.append(tempListing)
                break

    dfXls = pd.read_excel("cars.xlsx", engine='openpyxl')
    i = 0
    columns = ['Listing Name', 'Price', 'Location', 'Miles', 'Coverpic', 'link']
    #print("finalListing len: " + str(len(finalListings)))
    if firstRun == True:
        newListings = []
        newDf = pd.DataFrame(columns=columns)
        firstRun = False
    for tempListing in finalListings:
        append = False
        for i in range(0,len(dfXls)-1):
            if tempListing.listingName != dfXls.at[i, 'Listing Name'] and tempListing.price != dfXls.at[i, 'Price'] and tempListing.miles != dfXls.at[i, 'Miles']:
                append = True

        if len(dfXls) == 0:
            append = True
        for temp in newListings:
            if tempListing == temp:
                append = False

        if append:
            df = pd.DataFrame(columns=columns)
            row = tempListing.toArray()
            #print(row)
            df.loc[len(df)] = row
            #print(newDf)
            newDf = pd.concat([df, newDf], ignore_index=True)
            newListings.append(tempListing)

    #if len(newDf) > numListings:
        #newDf = newDf.drop(range(len(newDf)-numListings,len(newDf)))

newDf = pd.concat([dfXls, newDf], ignore_index=True)
newDf.to_excel("cars.xlsx", index=False, engine='openpyxl')


#print("newListings length: " + str(len(newListings)))

smtp_server = 'smtp.gmail.com'
smtp_port = 587
message = MIMEMultipart()
message['From'] = email
message['To'] = yourEmail
message['Subject'] = 'New Listings!'
message.attach(MIMEText("""
    <h1><span style="text-decoration: underline;"><strong>New Listings!!!</strong></span></h1>""", 'html'))


if len(newListings) > 0:
    imgNames = []
    for tempListing in newListings:
        response = requests.get(tempListing.coverpic)
        if response.status_code == 200:
            img = Image.open(BytesIO(response.content))
            imgPath = tempListing.listingName + ".jpg"
            imgNames.append(imgPath)
            #Image.open(imgPath)
            img.save(imgPath)

        with open(imgPath, 'rb') as imgFile:
            img = MIMEImage(imgFile.read())
            img.add_header('Content-ID', tempListing.listingName)
            message.attach(img)

        body = f"""
           <p><span style="text-decoration: underline;"><strong><img src="cid:""" + tempListing.listingName + """ alt="" /></strong></span></p>
           <p>""" + str(tempListing.price) + """</p>
           <p>""" + tempListing.listingName + """</p>
           <p>""" + tempListing.location + """</p>
           <p>""" + tempListing.miles + """</p>
           <p>""" + tempListing.link + """</p>
           <p>&nbsp;</p>
           <p>&nbsp;</p>"""
        message.attach(MIMEText(body, 'html'))

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(senderEmail, appPassword)
        server.sendmail(senderEmail, yourEmail, message.as_string())
        print("Email sent successfully!")
    except Exception as e:
        print(f"Failed to send email: {e}")
    finally:
        server.quit()

    for imgName in imgNames:
        if os.path.exists(imgName):
            os.remove(imgName)

else:
    print("No new listings!")

driver.close()
