import os.path
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from datetime import date, timedelta
import time


# Set the required variables
print( "Set a date variable of: " + str(date.today() + timedelta( days=-30 ) ) )
start = date.today() + timedelta( days=-30 )
file = 'Daily_Incident_Log.xlsx'
url = 'https://report.boonecountymo.org/mrcjava/servlet/SH01_MP.I00070s'


download_dir = str(os.getcwd())
chrome_options = webdriver.ChromeOptions()
preferences = {"download.default_directory": str(download_dir), "safebrowsing.enabled": "false" }
chrome_options.add_experimental_option("prefs", preferences)

if os.path.exists(file):
    print ("The file,'"+file+"' already exists in "+os.getcwd()+" and is being removed.")
    os.remove(file)
    print( file + " has been removed from the directory." )
else:
    print( "The file '" + file + "' does not exist in the " + os.getcwd() + " directory." )


driver = webdriver.Chrome(options=chrome_options)
driver.get(url)
WebDriverWait( driver, 5 ).until( EC.presence_of_element_located( (By.ID, 'startDate') ) )
print( "Web page is ready!" )
WebDriverWait( driver, 5 ).until( EC.element_to_be_clickable( (By.ID, 'startDate') ) )
print("Clearing the 'From' datepicker field...")
driver.find_element_by_name( 'startDate' ).clear()
print("Adding the date for 30 days back to the 'From' datepicker...")
driver.find_element_by_name( 'startDate' ).send_keys( str( start.strftime( "%m/%d/%Y" ) ) + Keys.RETURN )
print("Pressing the Search button...")
WebDriverWait( driver, 5 ).until( EC.element_to_be_clickable( (By.CSS_SELECTOR, 'button.btn:nth-child(1)' ) ) ).click()
print( "Downloading the file by pressing the 'Export to Excel' button..." )
WebDriverWait( driver, 5 ).until( EC.element_to_be_clickable( (By.CSS_SELECTOR, 'a.btn:nth-child(3)') ) ).click()
print("Waiting for the export...")
time.sleep( 10 )
print( "Closing the browser." )
driver.close()