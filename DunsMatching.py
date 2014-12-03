from selenium import webdriver
from selenium.webdriver.common import keys
#import seleniumLeads5c # can't import a module if the code runs automatically! if name == main...etc needed!
from bs4 import BeautifulSoup
import cPickle as pickle
import random
import time
import re
import xlwt
import gzip
import sys
import xlrd
import unittest
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from fuzzywuzzy import fuzz
#import collections # look up this 


#html = driver.page_source
#print html
#it works! but i can't click around into visitor list
'''
VisitorList = driver.find_element_by_id("ctl00_Menu1n1")
VisitorList.click()
'''
def login():
    # initiate web browser
    driver= webdriver.PhantomJS(executable_path="C:\phantomjs-1.9.7-windows\phantomjs.exe")
    print " phantom!" ## 
    driver.get("https://mintuk.bvdinfo.com/")
    print driver.title
    print "successful at getting website "##
    time.sleep(3)
    try:
        wait = WebDriverWait(driver,30) ## 
        
        driver.find_element_by_xpath('//input').click() #should get me past the acceptable usage page.
        #wait.until(EC.element_to_be_clickable((By.XPATH,'//input'))) ##
        print driver.current_url
    except:
        pass ##
    
    driver.get("https://mintuk.bvdinfo.com/") ## swapped with pass
    print driver.current_url
    """
    wait = WebDriverWait(driver,30)
    loginbutton =wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="headerLoginClient"]')))
    loginbutton.click()
    """
    
    #assert "Lead" in driver.title #### xTESTING###
    print driver.title
    
    # client login
    #elem = driver.find_element_by_id("headerLoginClient") #web object has no method send keys
    #elem.click() # equivalent to pressing enter.
    print driver.title

    # client id entry
    wait = WebDriverWait(driver,30)
    clientName =wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="inputUser"]')))
    clientName.send_keys('')
    time.sleep(1)##
    #clientName.send_keys(keys.RETURN)
    # password entry
    wait = WebDriverWait(driver,30)
    clientPassword =wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="inputPassword"]')))
    time.sleep(1)##
    clientPassword.send_keys('')
    # click login in button
    wait = WebDriverWait(driver,30)
    click_login = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="loginBoxTxt"]/table/tbody/tr[3]/td/table/tbody/tr/td[2]/div')))
    click_login.click()
    
    print driver.title
    time.sleep(random.uniform(2.0,3.0)) ## 14b ##
    print driver.title
    return driver

def getInnerHTML(element,xpath):
    try:
        datapoint = element.find_element_by_xpath(xpath).get_attribute('innerHTML')
    except:
        datapoint = None
    return datapoint

def verifyMatchesFound(driver):
    """Ensure that mint was able to find a match and isn't saying 'no match found, found x instead'. """
    try:
        wait = WebDriverWait(driver,30) ## 14b ##
        output = wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="dResults"]/div'))).get_attribute('innerHTML')
        verification = not(output == "No results found")
    except:
        verification = True
    return verification
    pass # for now
def evaluateMatchScore():
    pass # for
    # previously known as ...

def enterCompanyName(driver,company):
    # performs action on the same page as the functions following it. no need to return driver.
    try:
        wait = WebDriverWait(driver,30) ## 14b ##
        searchBox = wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="ctl00_Global_Main_iQuick"]')))
        print 'found searchBox!'##
        
    except:
        ## 14b ##
        searchBox = wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="ctl00_Global_Main_iQuick"]')))
        print 'found searchBox eventually!'##
   
    time.sleep(random.uniform(0.2,0.5))##
    searchBox.send_keys(company)
    time.sleep(random.uniform(2.0,4.0))
    return driver

class companyObj():
    def __init__(self,compName,regNum,duns,address,turnover):
        self.compName = compName
        self.regNum = regNum
        self.duns = duns
        self.address = address
        self.postcode = " ".join(address.split(' ')[-3:-1])
        self. turnover = turnover
        self.matchScore = None

def collectCompanyDetails(driver):
    wait = WebDriverWait(driver,30)
    resultsTable = wait.until(EC.presence_of_element_located((By.XPATH,' //*[@id="resultsList"]')))  
    resultsList = resultsTable.find_elements_by_xpath('./LI')
    """ collects info from webpage and paass it on to te objects. """
    count = 0
    listOfMatches = []
    while count < len(resultsList):
        li = resultsList[count]
        name = getInnerHTML(li,'.//*[@class="name"]')
        regNum = getInnerHTML(li,'.//*[@class="regnum"]')
        duns =  getInnerHTML(li,'.//*[@class="duns"]')
        address = getInnerHTML(li,'.//*[@class="address"]')
        turnover = getInnerHTML(li,'.//*[@class="turnover"]')
        companyDets = companyObj(name,regNum,duns,address,turnover)
        listOfMatches.append(companyDets)
        print name
        print "%s 's cro is %s" % (name,regNum)
        print "%s 's duns is %s" % (name,duns)
        print "%s 's address is %s" % (name,address)
        print "%s 's turnover is %s" % (name,turnover)
        count +=1
        time.sleep(1) ## slow it down a bit. don't melt bvd's servers.
        
    return listOfMatches # returns list of companyObj. each containing match details. 

def resetSearchBox(driver,company):
    driver.execute_script("window.scrollTo(document.body.scrollHeight,0);")
    print "scrolled to bottom"
    numOfDels = len(company)
    #print numOfDels
    wait = WebDriverWait(driver,30) ## 14b ##
    searchBox = wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="ctl00_Global_Main_iQuick"]')))
    
    searchBox.send_keys(numOfDels*keys.Keys.BACK_SPACE)
    print "backspaces pressed!"
    print "search box cleared!"
    time.sleep(random.uniform(1.5,2.5))
    #return driver ## # driver can update itself! yay!
    
def cycleCompanyName(driver, company):
    print "I'm cycling through company %s " % company
    driver = enterCompanyName(driver,company)
    if verifyMatchesFound(driver) == True:        
        print "the company we are searching for is %s !!!" % company ##
        time.sleep(random.uniform(1.7,2.3))
        listOfMatches = collectCompanyDetails(driver)
        #resetSearchBox(driver,company)
        #return listOfMatches
    else:
        time.sleep(random.uniform(1.7,2.3))
        listOfMatches =[]
        
    resetSearchBox(driver,company)   
    return listOfMatches 
    
def dunsMatchOutput(array):
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('processed',cell_overwrite_ok=True)
    filename = "PortusDunsMatch"
    row = 0
    sheet.write(row,0,'Index')
    sheet.write(row,1,'Company Name')
    sheet.write(row,2,'CHNo')
    sheet.write(row,3,'Duns')
    sheet.write(row,4,'Address')
    sheet.write(row,5,'Full Post Code')
    sheet.write(row,6,'Turnover')
    row = 1
    for compName,compNameMatchList in array.items():
        for compObj in compNameMatchList:
            sheet.write(row,1,compObj.compName)
            sheet.write(row,2,compObj.regNum)
            sheet.write(row,3,compObj.duns)
            sheet.write(row,4,compObj.address)
           
            sheet.write(row,5,compObj.postcode)
            sheet.write(row,6,compObj.turnover)
            row+=1 
    workbook.save('%s.xls'% filename)

def inputArray(filename):
    inputArray =[]
    with open(filename,'r') as f:
        for line in f:
            inputArray.append(line.strip())
    return inputArray
    

def process(filename):    
    companies = inputArray(filename)[:1000]
    driver = login()
    companyDicc = dict()
    for company in companies:
        companyDicc[company]= cycleCompanyName(driver, company)
    dunsMatchOutput(companyDicc)
        
    """   
    for key,item in companyDicc.items():
        if item !=[]:
            print key
            print item[0].compName
            print item[0].duns
            print item[0].address
            print item[0].postcode
            print '\n'
    """
    driver.quit()
    
process('PortusDunsNumbers.txt')
