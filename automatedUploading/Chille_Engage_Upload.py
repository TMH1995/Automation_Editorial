import pickle

import Null as Null
import openpyxl
import xlrd
from selenium.webdriver.support.select import Select
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.support.ui import Select
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.webdriver.support.wait import WebDriverWait
from collections import defaultdict


driver = webdriver.Chrome(executable_path="C:\Program Files\Drivers\Chrome_driver\chromedriver.exe")
wk = openpyxl.load_workbook("C:\Editorialsheet\Science Techbook_Grade 8_Unit 1.xlsx")
sheet = wk['Engage']

ActivityAssetSheet = xlrd.open_workbook("C:\Editorialsheet\Science Techbook_Grade 8_Unit 1.xlsx")
AssetDetails = ActivityAssetSheet.sheet_by_name('Engage')

Old_Type=Null

def read_sheet():
    # finding Rows have data
    rows = AssetDetails.nrows
    columns = AssetDetails.ncols
    # print("Total rows are: ", str(rows))
    # print("Total columns are: ", str(columns))

    return sheet,rows, columns


def row_read(r,c, data):

    data = []
    for i in range(c):
        val=AssetDetails.cell_value(r, i)
        data.append(val)
    print(data)

    return data


def data_classification(i,data):

    Type=data[1]
    Page_Num=data[2]
    Page_Title=data[3]
    Content=data[4]
    GUID=data[5]

    information=[Type,Page_Num,Page_Title,Content,GUID]

    if Type=="Concept title":
        # operation
        Create_Concept_Title(information)
        Old_Type=Type

    if Type == "Concept Subtitle":
        # operation
        Old_Type = Type

    if Type == "Teacher Note":
        # operation
        Old_Type = Type

    if Type == "New Line":
        # operation
        New_Line(information)


    if Type == "CIT Text":
        # operation
        Old_Type = Type

    if Type == "Image":
        # operation
        Old_Type = Type

    if Type == "TEI":
        # operation
        Old_Type = Type

    if Type == "Lesson question":
        # operation
        Old_Type = Type

    return

def Create_Concept_Title(Information):

    return




def New_Line(Information):
    if Information[0]=="Concept title":
        Create_Concept_Title(Information)
        Old_Type=Information[0]


    return

# data operations
def data_Operations():
    sheet_parameters=[]
    sheet_parameters=read_sheet()
    sheet=sheet_parameters[0]
    rows=sheet_parameters[1]
    columns=sheet_parameters[2]
    print("rows are: ",rows,"columns are: ",columns,"and the sheet name is: ",sheet)

    data=[]
    for i in range(1,rows):
        # print("i here is: ",i)
        data=row_read(i,columns,data)
        # print(data)
        data_classification(i,data)     # Fill the parameters of each row form the excel sheet at the asset page
    return

def open_Editorial():

    driver.get("https://app.discoveryeducation.com/learn/signin?next=http%3A%2F%2Feditorial.discoveryeducation.com")

    driver.find_element_by_id("username").send_keys("s.anwer@uniparticle.com")
    driver.find_element_by_id("password").send_keys("discovery")
    driver.find_element_by_xpath("//*[@id='comet-page-shell__product-well']/div[1]/div[3]/div/div/div/div[2]/form/div[3]/button").click()
    return


def Engage_Creation():
    TechBookTools = driver.find_element_by_xpath("//*[@id='nav-third_link']")
    AttachAssets_to_TB_Concepts = driver.find_element_by_xpath("//*[@id='cont31']/a")
    actions = ActionChains(driver)
    actions.move_to_element(TechBookTools).move_to_element(AttachAssets_to_TB_Concepts).click().perform()
    element= driver.find_element_by_id("guidTaxId")
    drp=Select(element)
    drp.select_by_visible_text('Science Techbook S&S')




open_Editorial()
Engage_Creation()
data_Operations()