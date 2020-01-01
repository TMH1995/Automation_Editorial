import pickle
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
wk = openpyxl.load_workbook("C:\KnowledgeBank\Assets_Egypt.xlsx")
sheet = wk['images_P2']

ActivityAssetSheet = xlrd.open_workbook("C:\KnowledgeBank\Assets_Egypt.xlsx")
AssetDetails = ActivityAssetSheet.sheet_by_name('images_P2')

# Getting all sheet parameters

def read_sheet():
    # finding Rows have data
    rows = AssetDetails.nrows
    columns = AssetDetails.ncols
    # print("Total rows are: ", str(rows))
    # print("Total columns are: ", str(columns))

    return sheet,rows, columns


# Getting the value of each row

def row_read(r,c, data):

    data = []
    for i in range(c):
        val=AssetDetails.cell_value(r, i)
        data.append(val)
    print(data)

    return data

# Data Classification

def data_classification(i,data):

    Type=data[1]
    image_title=data[2]
    path=data[3]
    Sibiling_Language=data[4]
    Sibiling_Lang_Title=data[5]
    grade_range=data[6]
    Title_Description=data[7]
    Description=data[8]
    Producer=data[9]
    Publisher=data[10]
    GUID=data[11]
    SUR_Mediagroup=data[12]
    volume=data[13]
    Sub_Directory=data[14]
    Format=data[15]

    image_Title_fill(image_title)
    image_Title_Description(Title_Description)
    Producer_selector(Producer)
    Audience_selector()
    SUR_selection()
    #Save_asset()
    #GUID=get_GUID(GUID)
    print("the Guid is: ",GUID)
    write_data(i,11,GUID)

    #print("Operations fn data: ")
    # print(Type," ,",image_title," ,",path," ,",Sibiling_Language," ,",Sibiling_Lang_Title," ,",grade_range," ,",Title_Description)
    # print(Description," ,",Producer," ,",Publisher," ,",GUID," ,",SUR_Mediagroup," ,",volume," ,",Sub_Directory," ,",Format)

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
    for i in range(2,rows):
        # print("i here is: ",i)
        data=row_read(i,columns,data)
        # print(data)
        data_classification(i,data)     # Fill the parameters of each row form the excel sheet at the asset page
    return



# open Editorial Web Page
def open_Editorial():

    driver.get("https://app.discoveryeducation.com/learn/signin?next=http%3A%2F%2Feditorial.discoveryeducation.com")

    driver.find_element_by_id("username").send_keys("s.anwer@uniparticle.com")
    driver.find_element_by_id("password").send_keys("discovery")
    driver.find_element_by_xpath(
        "//*[@id='comet-page-shell__product-well']/div[1]/div[3]/div/div/div/div[2]/form/div[3]/button").click()
    return

# Opening asset data fill page

def Asset_Creation():
    # open Asset create page
    Manage_AssetsandTaxonomy = driver.find_element_by_xpath("//*[@id='nav-first_link']")
    Asset_Create = driver.find_element_by_xpath("//*[@id='cont12']")
    actions = ActionChains(driver)
    actions.move_to_element(Manage_AssetsandTaxonomy).move_to_element(Asset_Create).click().perform()
    # select image type to create an asset
    element = driver.find_element_by_xpath("/html/body/div[2]/div/form/select")
    drp = Select(element)
    drp.select_by_visible_text('Image (Image)')

    driver.find_element_by_xpath("/html/body/div[2]/div/form/input[2]").click()  # click create asset button

    driver.find_element_by_xpath(
        "/html/body/div[6]/div/div/form/div[2]/div[4]/input").click()  # choose active status at status bar

    driver.find_element_by_xpath("//*[@id='tab-1']").click()  # click on Asset details button
    return


# Asset fill methods:

def image_Title_fill(Title):
    driver.find_element_by_id("AssetForm_tblAssetMetaData_title").send_keys(Title)
    return
def image_Title_Description(Title):
    driver.find_element_by_id("AssetForm_tblAssetMetaData_title_description").send_keys(Title)
    return
def Producer_selector(producer_name):
    producer_name="tariq"
    driver.find_element_by_xpath("//*[@id='tab-1-wrap']/ol[4]/li[3]/div[2]/span/div/input[1]").send_keys(producer_name)
    return
def Audience_selector():
    driver.find_element_by_xpath("//*[@id='grades_0']").click()
    driver.find_element_by_xpath("//*[@id='grades_2']").click()
def SUR_selection():
    driver.find_element_by_xpath("//*[@id='tab-4']").click()
    driver.find_element_by_xpath("//*[@id='checkboxServiceUsageRights4']").click()
    driver.find_element_by_xpath("//*[@id='checkboxMediaGroupId4XD0D8B0B6_1934_4222_AA53_BEC03745684E']").click()
    return

def Save_asset():
    driver.find_element_by_xpath("//input[@type='submit' and @value='Save Asset']").click()
    return
def get_GUID(GUID):
    GUID=driver.find_element_by_xpath("/html/body/div[7]/div/div/div[1]/div/div[2]/span[1]").text
def write_data(r,c,data):
    AssetDetails.cell(r,c).value=data
    return



open_Editorial()
Asset_Creation()
data_Operations()