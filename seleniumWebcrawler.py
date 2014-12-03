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
    driver.get("http://new.leadforensics.com/")
    
    time.sleep(3)
    try:
        wait = WebDriverWait(driver,30) ## 
        
        driver.find_element_by_xpath('//input').click() #should get me past the acceptable usage page.
        #wait.until(EC.element_to_be_clickable((By.XPATH,'//input'))) ##
        print driver.current_url
    except:
        pass ##
    
    driver.get("http://new.leadforensics.com/") ## swapped with pass
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
    clientName =wait.until(EC.element_to_be_clickable((By.ID,"ClientLogin_TextBoxUserName")))
    clientName.send_keys('mguo')
    time.sleep(1)##
    #clientName.send_keys(keys.RETURN)
    # password entry
    wait = WebDriverWait(driver,30)
    clientPassword =wait.until(EC.element_to_be_clickable((By.ID,"ClientLogin_TextBoxPassword")))
    time.sleep(1)##
    clientPassword.send_keys('D       4')
    # click login in button
    wait = WebDriverWait(driver,30)
    click_login = wait.until(EC.element_to_be_clickable((By.ID,"ClientLogin_ButtonLogin")))
    click_login.click()
    
    print driver.title
    time.sleep(random.uniform(2.0,3.0)) ## 14b ##
    print driver.title
    return driver

def navigateToReport(driver):
    # navigates to search page.
    wait = WebDriverWait(driver,30)
    Search = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_Menu1n3"]/table/tbody/tr/td/a/img')))
    time.sleep(random.uniform(0.3,0.5))
    print "got to search page!"
    Search.click()#modified here!
    return driver

def clickSearchButton(driver):
    wait = WebDriverWait(driver,30)
    searchButton= wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_MainContent_ImageButtonSearch"]')))
    #searchButton = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_ImageButtonSearch"]')
    #print "Found SearchButton!"
    searchButton.click()
    return driver

def enterCompanyName(driver,company):
    # performs action on the same page as the functions following it. no need to return driver.
    try:
        wait = WebDriverWait(driver,30) ## 14b ##
        searchBox = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_MainContent_TextBoxSearch"]')))
        
        #print 'found searchBox!'##
        
    except:
        wait = WebDriverWait(driver,30)## 14b ##
        searchBox = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_MainContent_TextBoxSearch"]')))
        #print 'found searchBox eventually!'##
    searchBox.clear()
    time.sleep(random.uniform(0.2,0.5))##
    searchBox.send_keys(company)

def findCompanyByName(driver,company):
    # takes the company name and finds the page with 'history' button.
    
    enterCompanyName(driver,company)##
    #print " entered company Name successfully! "##
    clickSearchButton(driver)
    #print " clicked search button successfully!"##
    """
    except:
        time.sleep(60)
        returnButton = driver.find_element_by_id("ctl00_MainContent_ImageButtonVisList")
        returnButton.click()
        time.sleep(random.uniform(2.0,3.0)) # hideously complex! now it even has recusion! this could ddos the website.
        findCompanyByName(driver,company)
    """

def selectCompanySearchResult(driver):
    # we will only select one. it is simpler to just deal with names that matches to just one company.
    wait = WebDriverWait(driver,20)
    businessFound = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_MainContent_Search_Telerik_ctl00_ctl04_LabelBusiness"]')))
    #print "business found!" //*[@id="ctl00_MainContent_Search_Telerik_ctl00_ctl04_LabelBusiness"]
    businessFound.click()
    time.sleep(random.uniform(1.0,2.0)) ##?REPLACE??
    return driver

def searchCriteria(driver):

    #search by company name, across all time period. not include hidden companies. however, modify 'Type of Search'
    #print driver.current_url
    wait = WebDriverWait(driver,20)
    searchCriteriaMatchExact =  wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_MainContent_RadioButtonExact"]')))
    searchCriteriaMatchExact.click()
  

def viewHistory(driver):

    wait = WebDriverWait(driver,20) ### 60 down to 30 ## 30 down to 20
    viewHistory = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_MainContent_ImageButtonViewHistory"]')))
    #print "History button found!" ##
    viewHistory.click()
    print driver.title ### scaffold ###
    time.sleep(random.uniform(0.1,0.5))##
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    compName = driver.find_element_by_xpath('//*[@id="ctl00_PageHeadingContent_LabelBusinessName"]').text ##
    print compName
    time.sleep(random.uniform(1.0,2.0)) ### Replace with explicit waits please. maybe too short? i need to grab everything! 
    return driver

def clickReturnButton(driver):
    wait = WebDriverWait(driver,20)### 60 down to 30 ## 30 down to 20
    returnButton= wait.until(EC.element_to_be_clickable((By.ID,'ctl00_MainContent_ImageButtonVisList')))
    #print "Return Button Found!"
    returnButton.click()
    return driver

class company(object):
    def __init__(self,companyInfo,sessions):
        self.companyInfo = companyInfo # can objects be passed into the __init__ parameter of another object???
        self.sessions =sessions # where sessions is a dictionary. tbd ###

class companyInfo(object):
    def __init__(self,name,address,telephone,webpage,industry,SICCodes,employees,founded):

        self.name = name
        self.address = address
        self.telephone = telephone
        self.webpage = webpage
        self.industry = industry
        self.SICCodes = SICCodes
        self.employees = employees
        self.founded = founded
class Sessions(object): ################rid of#######
    # you'd think they should be related! Sessions and Session, delete this
    def __init__(self):
        self.SessionsDicc ={}
class Session(object):
    def __init__(self,visitor,duration,referralLink,history_header_id):
        self.visitor = visitor
        self.duration = duration
        self.referralLink = referralLink
        self.history_header_id = history_header_id
        self.pageVisits ={} # tbd 
    def returnPageVisits(self):
        return self.pageVisits
    def pagesPerSession(self):
        return len(self.pageVisits)
        
class pageVisit(object):
    def __init__(self,dateTime,duration,webPage,history_row_id):
        self.dateTime = dateTime
        self.duration = duration
        self.webPage = webPage
        self.history_row_id = history_row_id
        
def companyInfoCollector(driver): 
    print " comp info collector!"
    try:
        name  = driver.find_element_by_xpath('//*[@id="ctl00_PageHeadingContent_LabelBusinessName"]').text
    except:
        name = None
    print " comp name found ", name
    try:
        time.sleep(random.uniform(2.0,3.0)) ##
        address  = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_LabelCombinedAddress"]').text
        
    except:
        address = None
    print " comp add found ",address
    try:
        telephone = driver.find_elemnt_by_xpath('//*[@id="ctl00_MainContent_LabelLITel"]').text
    except:
        telephone =None
    print " comp tel found ", telephone
    try:
        webpage = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_HyperlinkWebsite"]').text
    except:
        webpage = None
    print " comp webpage found ",webpage
    try:
        industry = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_LabelIndustry"]').text
    except:
        industry = None
    print " comp industry found ",industry
    try:
        SICCodes = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_LabelSICCode"]').text
    except:
        SICCodes = None
    print " comp sic codes found ",SICCodes
    try:
        employees = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_LabelEmployees"]').text
    except:
        employees = None
    print ' comp ees found',employees
    try:
        founded = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_LabelFounded"]').text
    except:
        founded = None
    print ' comp year founded found' ,founded
    companyDetails = companyInfo(name,address,telephone,webpage,industry,SICCodes,employees,founded)
    
    #companyDetails = companyInfo("name","address","telephone","webpage","industry","SICCodes","employees","founded")
    return companyDetails ##
    
def dataCollector(table):
    """
    meant to be used in parsePageHTML without being called upon in main loop.
    takes table. collects data of session history and company info effectively. explore option of objects vs dicts
    """
    # can't finish company info here! only the table.
    
    history_headers = table.find_elements_by_xpath('.//*[@class="history_header"]')    
    history_rows = table.find_elements_by_xpath('.//*[@class="history_row" or @class = "history_altrow"]') 
    # they seem to be well ordered when collected into a list, i don't know if i should trust the results however.      
            
    sessions =[]
    for history_header in(history_headers):
        history_header_id = int(history_header.get_attribute(u"id").split("__")[1])
        history_header_info = history_header.find_elements_by_xpath('./TD')
        visitor = history_header_info[0].text # visitor: 1 is printed out! just want 1maybe?
        print visitor
        duration = history_header_info[1].text
        print duration
        referralLink = history_header_info[2].text # referral link, take into account of when link is actually provided as href
        print referralLink
        
        #print "history Header ID is" , history_header_id
        visitorSession = Session(visitor,duration,referralLink,history_header_id)
        print "object ok!"
        sessions.append(visitorSession)

    pagesVisited =[]    
    for history_row in history_rows: 
        history_row_info = history_row.find_elements_by_xpath('./TD')
        dateTime = history_row_info[0].text
        
        duration = history_row_info[1].text
        webPage = history_row_info[2].text
        history_row_id = int(history_row.get_attribute(u"id").split("__")[1])
        #print "history rowID is ", history_row_id
        
        pageVisited = pageVisit(dateTime,duration,webPage,history_row_id)
        pagesVisited.append(pageVisited)
    sessionsDicc ={}
    #CompanySessions = Sessions()
    for session in reversed(sessions):
        print "history header id is", session.history_header_id ##scaffolding
        for page in reversed(pagesVisited):# explain why this is the right order.
            if page.history_row_id > session.history_header_id:
                print "page row id is ", page.history_row_id ##scaffolding
                session.pageVisits[page.history_row_id] = page  # these ids are so uninformative you might as well not have them... or have index instead.
                pagesVisited.remove(page)
        sessionsDicc[session.history_header_id] = session
        #CompanySessions.SessionsDicc[session.history_header_id] = session # unecessarily complex? too many Sessions?
    #print CompanySessions.SessionsDicc
    #return CompanySessions.SessionsDicc
    print sessionsDicc # it was all done in objects previously. now suddenly a dictionary.?!?!
    return sessionsDicc

def parsePageHTML(driver):
    
    wait = WebDriverWait(driver,60)

    compInfo= companyInfoCollector(driver) 

    table = wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="ctl00_MainContent_VisitorList_Telerik_ctl00"]/tbody')))
    
    sessionsDicc = dataCollector(table)
    comp = company(compInfo,sessionsDicc)
    
    print "HTML table found!"#
    
    return comp
    

def searchCompanies(driver,companiesList,start):
    
    try: ############### POSSIBLY REDUNDANT!############
        time.sleep(2)
        searchCriteria(driver)    
        companyHistorySourceCode= dict() # what happens if a mistake is made? will an empty dictionary be returned?
        count = 0
        
        # should the above be outside the try block? THIS WAY WORKS BUT I'M CONFUSED
        errorLog = dict()
        for company in companiesList:
            count+=1 # if mistake is made in nth iteration, count = n would be returned. NOT TRUE NEMORE
            try:
                findCompanyByName(driver,company)
                #print "FoundCompanyByName now executes!" ##
                driver = selectCompanySearchResult(driver) ##
                try: #LAZY TO USE TRY EXCEPT HERE! FIND ANOTHER METHOD!
                    # we have to rely on a deduped company list here.
                    driver = viewHistory(driver)##
                    companyHistorySourceCode[company] = parsePageHTML(driver)

                    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                    print "scrolled down" ## 
                    driver = clickReturnButton(driver) ## there might be an exception here... 
                    print "return button clicked" ##
                except:
                    #print 'no history to be seen!'
                    compName = driver.find_element_by_xpath('//*[@id="ctl00_PageHeadingContent_LabelBusinessName"]').text
                    print compName ##
                    print "alternative route taken" ##
                    companyHistorySourceCode[company] = parsePageHTML(driver)
                    
                    ## if no history, driver = viewHistory in try block won't execute. then driver = selectCompanySearchResult(driver)
                    #testing code to see company name
                
                    driver = clickReturnButton(driver)## repetition?????
                    #count+=1 # if mistake is made in nth iteration, count = n would be returned. NOT TRUE NEMORE
                    #former location. if wrong, you must return to this and think of success count + error count method!
            except:
                #time.sleep(random.uniform(2.0,4.5))
                ## deal with annoying nonsensical no results found button (when there is a company by that name!)
                #driver.find_element_by_xpath('//*[@id="ctl00_MainContent_ImageButtonReturn"]')##
                print " nada" ##
                errorLog[company]= (count-1) + start # count? or count + x?
            #PROPER LOCATION FOR COUNT +=1! TEST IT!
    except:      
        print " cycle is broken! return count and what source code we have"
        
    return companyHistorySourceCode, count, errorLog, driver # if try and except block are not executed normally, count+=1 is not excecuted.
        
def save(filename, dicts):
    """ save objects into a compreseed diskfile."""
    sys.setrecursionlimit(10000)
    fil = gzip.open(filename, 'wb',-1) #changed from opening file in binary...
    pickle.dump(dicts, fil) # dumping one dictionary at a time... right?
    fil.close()
    
def getCompaniesList():

    workbook = xlrd.open_workbook('E:\Marketing Insight And Analysis\LeadForensics\VisitorAtoZ 2014 Jan - Mid May.xlsx')
    worksheet = workbook.sheet_by_name('Results excluding CompName Dups')

    companiesList = list()
    
    num_rows = worksheet.nrows - 1
    curr_row = 0
    while curr_row < num_rows:
	curr_row += 1
	row = worksheet.row(curr_row)
	cell_value = worksheet.cell_value(curr_row, 0) # column 0 is the first column!
	companiesList.append(cell_value)
	
    return companiesList

def outputToExcel(array,filename):
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('Unprocessable Companies',cell_overwrite_ok=True)

    row = 0
    sheet.write(row,0,'Company Name')
    sheet.write(row,1,'Index')

    row = 1
    for compName,ind in array.items():
        sheet.write(row,0,compName)#names.decode('utf-8')
        sheet.write(row,1,ind)
        row+=1
    
    workbook.save('%s.xls'% filename)
    
def processMain(companiesList,start,end):
    companiesList = companiesList[start:end]
    driver = navigateToReport(login())
    beginTime = time.time() #TESTING CODE
    tup = searchCompanies(driver,companiesList,start) ##### this pickling code is messy. sort it out#####
    companyHistorySourceCode= tup[0]
    #print "The keys are ", companyHistorySourceCode.keys(), # tuple access is fine. the dictionary was created.
   
    count = tup[1]
    errorLog = tup[2]
    driver = tup[3]
    driver.quit # this hopefully closes the loop every time, closing the browser instance instead of having 100 instances.
    terminTime = time.time() #Testing CODE
    print 'cycle time', terminTime - beginTime 

    # persistence
    beginTime = time.time()
    save('html dictionary 2013( %s to %s )'%(start,end), companyHistorySourceCode)
    print 'OUTPUT CREATED!'
    terminTime = time.time()
    print 'save time', terminTime - beginTime 
    print "number of companies processed is %d." % count
    print "starting index along list for next iteration is %s." % (count + start)# Where one should start for next turn. skipping error???

    for i in errorLog.items():
        print i
    outputToExcel(errorLog, 'errorLog 2013 ( %s to %s )'%(start,end))
    driver.quit()
    time.sleep(random.uniform(3.0,5.0))


for i in range(25,500,25):
    processMain(getCompaniesList(),i,i+25) 


## 75
#144
#162
#235,239

#457  is bad.
#1600 - 1640
#2000- 2500 encountered major trouble
#2391 to 2400 is suspect!
#2543 is bad. 
#problem with the 0-1000 after 700 as well...
#2825 is bad
#2827 is bad
#3675 is bad
#3838 is bad
#4378 is bad
#html 2013 (457 - 459) does not exist
# process company details here.






