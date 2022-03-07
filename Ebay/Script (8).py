import time
from tqdm import tqdm
import sys
import random
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import *
from selenium.webdriver.support.ui import *
import re
import pandas as pd
from openpyxl import load_workbook
import math
from datetime import datetime,timedelta
import dateparser
from bs4 import BeautifulSoup as bs4
import datetime as dit
from selenium.webdriver.common.action_chains import ActionChains
import os
from openpyxl.utils.cell import coordinate_from_string


class EbayScanner():
   ## Defining options for chrome browser
   options = webdriver.ChromeOptions()
   options.add_argument("--ignore-certificate-errors")
   Browser = webdriver.Chrome(executable_path = "chromedriver",options = options)

   #Global variables
   MainUrl = ''
   #time sleep for ItemSold variable
   TimeItemSold = 0.3

   #time for all the web driver wait elements
   webdriverwait_timesleep=  10


   FolderName = ""
   FileName = ""

   ExcelFile = ""

   AllRecords = []

   def ReadingExcelData(self):
      #Open Excel File
      UserData = pd.read_excel('Ebay_scanner.xlsx',header = None)
      #Creating folder
      self.CreatingFolder()
      len1 =  len(UserData.values)
      print("Folder Name: "+self.FolderName)
      i = 0
      #Extracting Urls from the excel and pass it to the function
      for data in UserData.values:
         self.MainUrl = data[0]
         self.ScrapingProductURLs()
         len1 -= 1
         i+=1 
         print(str(i)+"/"+str(len(UserData.values))+" User(s) Completed remaming:  "+str(len1))
         AllRecords = []


   #Create Folder
   def CreatingFolder(self):
      #If folder name already exists duplicate the name with Nth number "folder (N)"
      self.FolderName = "Ebay "+str(dit.date.today())
      if os.path.isdir(self.FolderName):
         i = 0
         while True:
            self.FolderName = "Ebay "+str(dit.date.today())+" ("+str(i)+")"
            if os.path.isdir(self.FolderName):
               pass
            else:
               os.mkdir(self.FolderName)
               return
            i+=1
      else:
         os.mkdir(self.FolderName)



   def ScrapingProductURLs(self):
      print("Scraping URLs of product ")
      self.Browser.get(self.MainUrl)
      appendurls =[]
      time.sleep(5)

      #Scrape Seller Name
      SellerName = WebDriverWait(self.Browser, self.webdriverwait_timesleep).until(EC.presence_of_all_elements_located((By.XPATH, "//a[@class='mbid']")))
      SellerName = SellerName[0].text
      print("Seller Name : "+SellerName)

      #Excel file name with pattern Folder / SellerName.xlsx
      self.FileName = self.FolderName+"/"+SellerName+'.xlsx'

      #Start scraping data from each page till there is not page left
      pagenumber = 0
      while True:
         pagenumber += 1
         time.sleep(5)
         #There are two classes for product div
         try:
            #Class 1
            lis = WebDriverWait(self.Browser, self.webdriverwait_timesleep).until(EC.presence_of_all_elements_located((By.XPATH, "//ul[@id='ListViewInner']//li[@class='sresult lvresult clearfix li']")))
         except KeyboardInterrupt:
            sys.exit()
         except Exception as e:
            lis = []
            pass
         try:
            #Class 2
            lis2 = WebDriverWait(self.Browser, self.webdriverwait_timesleep).until(EC.presence_of_all_elements_located((By.XPATH, "//ul[@id='ListViewInner']//li[@class='sresult lvresult clearfix li shic']")))
         except KeyboardInterrupt:
            sys.exit()
         except:
            lis2 = []
            pass

         #if both class has nothing it will refresh page and try again then it will run if condition again still nothing then still break the loop
         if lis == [] and lis2 == []:
            self.Browser.refresh()
            time.sleep(5)
            try:
               lis = WebDriverWait(self.Browser, self.webdriverwait_timesleep).until(EC.presence_of_all_elements_located((By.XPATH, "//ul[@id='ListViewInner']//li[@class='sresult lvresult clearfix li']")))
            except KeyboardInterrupt:
               sys.exit()
            except Exception as e:
               lis = []
               pass
            try:
               lis2 = WebDriverWait(self.Browser, self.webdriverwait_timesleep).until(EC.presence_of_all_elements_located((By.XPATH, "//ul[@id='ListViewInner']//li[@class='sresult lvresult clearfix li shic']")))
            except KeyboardInterrupt:
               sys.exit()
            except:
               lis2 = []
               pass
            if lis == [] and lis2 == []:
               break


         #Extract product Url from each class div (Class1)
         for li in lis:
            li = li.get_attribute("innerHTML")
            bs4data = bs4(li,"html.parser")
            h3 = bs4data.find("h3",{"class","lvtitle"})
            a = h3.find("a")
            appendurls.append(a['href'])

         #Extract product Url from each class div (Class2)
         for li in lis2:
            li = li.get_attribute("innerHTML")
            bs4data = bs4(li,"html.parser")
            h3 = bs4data.find("h3",{"class","lvtitle"})
            a = h3.find("a")
            appendurls.append(a['href'])



         #Click on Next page button
         try:
            nextbutton = WebDriverWait(self.Browser, self.webdriverwait_timesleep).until(
            EC.element_to_be_clickable((By.XPATH, "//a[@class='gspr next']")))
            nextbutton.click();
         except KeyboardInterrupt:
            sys.exit()
         #if its not clickable
         except Exception as e:
            #check if its class has changed to disabled
            try:
               aria = WebDriverWait(self.Browser, self.webdriverwait_timesleep).until(EC.presence_of_element_located((By.XPATH, "//a[@class='gspr next-d']")))
               break
            except:
               #check if privacy policy notification is on the pass else pass
               try:
                  time.sleep(4)

                  actions = ActionChains(self.Browser)
                  for i in range(15):
                     actions.send_keys(Keys.PAGE_DOWN)
                     time.sleep(0.4)
                     actions.perform()

                  nextbutton = WebDriverWait(self.Browser, self.webdriverwait_timesleep).until(
                  EC.element_to_be_clickable((By.XPATH, "//a[@class='gspr next']")))


                  nextbutton.click();
               except Exception as e:
                  pass
         break


      print("Total Page(s): ",pagenumber)
      print("Total Product(s) found: ",len(appendurls))
      self.passingdata(appendurls)


   def passingdata(self,urls):
      print("Scraping Products")
      for urlindex in range(0, len(urls)):
         print ("{}/{} Processing {}: ".format(urlindex, len(urls), urls[urlindex]))
         self.ScrapeData(urls[urlindex])

   def BlackText(self):

      try:                                                                                
         ItemSold = WebDriverWait(self.Browser, self.TimeItemSold).until(EC.presence_of_element_located((By.XPATH, "//span[@class='vi-qtyS-hot  vi-qty-vert-algn vi-qty-pur-lnk']//a")))
      except:
         try:
            ItemSold = WebDriverWait(self.Browser, self.TimeItemSold).until(EC.presence_of_element_located((By.XPATH, "//span[@class='vi-qtyS-hot  vi-bboxrev-dsplblk vi-qty-vert-algn vi-qty-pur-lnk']//a")))
         except:
            try:
               ItemSold = WebDriverWait(self.Browser, self.TimeItemSold).until(EC.presence_of_element_located((By.XPATH, "//span[@class='vi-qtyS-hot  vi-bboxrev-dsplblk vi-qty-vert-algn']//a")))
            except:
               ItemSold = ""
      return ItemSold

   def RedText(self):
      try:                                                                                
         ItemSold = WebDriverWait(self.Browser, self.TimeItemSold).until(EC.presence_of_element_located((By.XPATH, "//span[@class='vi-qtyS-hot-red  vi-qty-vert-algn vi-qty-pur-lnk']//a")))
      except:
         try:
            ItemSold = WebDriverWait(self.Browser, self.TimeItemSold).until(EC.presence_of_element_located((By.XPATH, "//span[@class='vi-qtyS-hot-red  vi-bboxrev-dsplblk vi-qty-vert-algn vi-qty-pur-lnk']//a")))
         except:
            try:
               ItemSold = WebDriverWait(self.Browser, self.TimeItemSold).until(EC.presence_of_element_located((By.XPATH, "//span[@class='vi-qtyS-hot-red  vi-bboxrev-dsplblk vi-qty-vert-algn']//a")))
            except:
               ItemSold = ""
      return ItemSold


   def BlueText(self):
      try:                                                                                
         ItemSold = WebDriverWait(self.Browser, self.TimeItemSold).until(EC.presence_of_element_located((By.XPATH, "//span[@class='vi-qtyS  vi-qty-vert-algn vi-qty-pur-lnk']//a")))
      except:
         try:
            ItemSold = WebDriverWait(self.Browser, self.TimeItemSold).until(EC.presence_of_element_located((By.XPATH, "//span[@class='vi-qtyS  vi-bboxrev-dsplblk vi-qty-vert-algn vi-qty-pur-lnk']//a")))
         except:
            try:
               ItemSold = WebDriverWait(self.Browser, self.TimeItemSold).until(EC.presence_of_element_located((By.XPATH, "//span[@class='vi-qtyS  vi-bboxrev-dsplblk vi-qty-vert-algn']//a")))
            except:
               ItemSold = ""

      return ItemSold


   def ScrapeData(self,url):
      time.sleep(5)
      self.Browser.get(url)
      lastdate= ""
      firstdate =  ""
      #COl D
      Brand = ""
      #Col E
      EAN = ""
      #Col B
      excelurl = self.Browser.current_url
      #Col A
      SellerName = WebDriverWait(self.Browser, self.webdriverwait_timesleep).until(EC.presence_of_element_located((By.XPATH, "//span[@class='mbg-nw']"))).text
      
      #Col C

      Title = WebDriverWait(self.Browser, self.webdriverwait_timesleep).until(EC.presence_of_element_located((By.XPATH, "//h1[@id='itemTitle']"))).text

      #will scrape total sold product , Total sold product has three different colors each color has three different type of html structure
      if self.BlackText() == "":
         if self.RedText() == "":
            if self.BlueText() == "":
               ItemSold = ""
            else:
               ItemSold = self.BlueText()
         else:
            ItemSold = self.RedText()
      else:
         ItemSold = self.BlackText()


      if ItemSold != "":
         #Col F
         TotalItemSold = re.findall(r'\b\d[\d,.]*\b',ItemSold.text)[0]
         TotalItemSold = TotalItemSold.replace(',','')
         TotalItemSold = int(TotalItemSold)
      else:
         #Col F
         TotalItemSold = ""


      #Description Table extract Brand and Ean from description table
      try:
         table = WebDriverWait(self.Browser, self.webdriverwait_timesleep).until(EC.presence_of_all_elements_located((By.XPATH, "//div[@id='viTabs_0_is']//table[@role='presentation']//tr//td")))

         for t in range(len(table)):
            if table[t].get_attribute('class') == "attrLabels":
               if "brand" in table[t].text.lower():
                  Brand = table[t+1].text.strip()
               if "ean" in table[t].text.lower():
                  EAN = table[t+1].text.strip()
      except:
         pass

      if Brand == "":
         Brand = "NA"
      if EAN == "":
         EAN = "NA"



      try:
         #Click on Total item sold page
         ItemSold.click()
         time.sleep(4)
         #Get data Table
         datetable = WebDriverWait(self.Browser, self.webdriverwait_timesleep).until(EC.presence_of_all_elements_located((By.XPATH, "//div[@style='padding:1px 10px 0px 10px;_width:100%;']//table[@cellpadding='5']//tr")))


         a = datetable[1].get_attribute('innerHTML')

         #Converting table innerHTMl to bueatiful soup to get data easily
         bs4V = bs4(a, "html.parser")

         tds = bs4V.findAll("td")

         #For First Date#####
         for i in range(len(tds)):
            if i == len(tds)-2:
               firstdate = tds[-2].text


         ###


         #For Last Date, Last date col have 3 different type of html structure
         a1 = datetable[-1].get_attribute('innerHTML')


         bs4V1 = bs4(a1, "html.parser")

         tds1 = bs4V1.findAll("td")


         for i1 in range(len(tds1)):
            if i1 == len(tds1)-2:
               lastdate = tds1[-2].text

         if lastdate == "":

            a1 = datetable[-2].get_attribute('innerHTML')


            bs4V1 = bs4(a1, "html.parser")

            tds1 = bs4V1.findAll("td")


            for i1 in range(len(tds1)):
               if i1 == len(tds1)-2:
                  lastdate = tds1[-2].text

            if lastdate == "":
               a1 = datetable[-3].get_attribute('innerHTML')


               bs4V1 = bs4(a1, "html.parser")

               tds1 = bs4V1.findAll("td")


               for i1 in range(len(tds1)):
                  if i1 == len(tds1)-2:
                     lastdate = tds1[-2].text   

         ####
         
         #Col G            
         totalhours = self.calculateDate(firstdate,lastdate)


         #Col H
         dailyestimate = self.dailyestimate(TotalItemSold,totalhours)

         #Col I
         weeklyestimate = self.weeklyestimate(TotalItemSold,totalhours)

         #Col J
         monthlyestimate = self.monthlyestimate(TotalItemSold,totalhours)

         totalhours = str(totalhours) + " Hours"
      except Exception as e:
         totalhours = ""
         monthlyestimate = ""
         dailyestimate = ""
         weeklyestimate = ""

      DataDict = {"Seller Name": SellerName,"Ebay Link": excelurl,"Title of Listing": Title,"Brand Name": Brand,"EAN": EAN,"Ebay Total sold figure":str(TotalItemSold),"Time period sold in": totalhours,"Daily sale estimate":dailyestimate,"Weekly sale estimate":weeklyestimate,"Monthly Sale estimate":monthlyestimate}
      self.AllRecords.append(DataDict)
      df= pd.DataFrame(self.AllRecords, columns = None)
      df.to_excel(self.FileName, index=False)
                


   def calculateDate(self,firstdate,lastdate):
      #Calculating date difference in seconds then converting it to hours
      firstdate  = dateparser.parse(firstdate)
      lastdate = dateparser.parse(lastdate)


      if firstdate == lastdate:
         firstdate = datetime.now(tz=lastdate.tzinfo)
         
      totalseconds = firstdate - lastdate
      totalseconds = totalseconds.total_seconds()

      totalhours = round(totalseconds/3600,1)


      return totalhours


   #if divisible by zero exception return 0
   def dailyestimate(self,productsold,totalhours):
      try:
         return round(productsold/totalhours*24,1)
      except:
         return 0


   #if divisible by zero exception return 0
   def monthlyestimate(self,productsold,totalhours):
      try:
         return round((productsold/totalhours*24)*30,1)
      except:
         return 0


   #if divisible by zero exception return 0
   def weeklyestimate(self,productsold,totalhours):
      try:
         return round((productsold/totalhours*24)*7,1)
      except:
         return 0



e = EbayScanner()
#e.ScrapeData('https://www.ebay.co.uk/itm/Biofreeze-Spray-118ml-New/223800799271?hash=item341b913027:g:vf8AAOSw4tJd-d9~')
e.ReadingExcelData()
