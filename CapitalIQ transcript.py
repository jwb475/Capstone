from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Create an instance of ChromeOptions
opt = webdriver.ChromeOptions()

# Set the experimental option for debugger address
opt.add_experimental_option("debuggerAddress", "localhost:9222")

# Initialize the ChromeDriver with the specified options
driver = webdriver.Chrome(options=opt)

i = 190
ii = 189
length= 1000



# Function to click through page links
def click_through_pages(page_number):
    page_link_xpath = f"//a[contains(@href, 'Page${page_number}')]"
    WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, page_link_xpath)))
    page_link = driver.find_element(By.XPATH, page_link_xpath)
    page_link.click()
    print(f"Clicked on Page {page_number}")
    WebDriverWait(driver, 25).until(EC.invisibility_of_element_located((By.XPATH, "//*[@id='UpdateProgress1']/div[3]/div")))

def click_through_odd_pages(page_number):
    page_link_xpath = f"//a[contains(@href, 'Page$Next${page_number}')]"
    WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, page_link_xpath)))
    page_link = driver.find_element(By.XPATH, page_link_xpath)
    page_link.click()
    print(f"Clicked on Page {page_number}")
    WebDriverWait(driver, 25).until(EC.invisibility_of_element_located((By.XPATH, "//*[@id='UpdateProgress1']/div[3]/div")))


for i in range (190, length):
    try:
        
        WebDriverWait(driver, 20).until(EC.presence_of_all_elements_located((By.XPATH,"//td[@class=' cColSortedBG']/a[text()='Earnings Call']")))
        rows = driver.find_elements(By.XPATH,"//td[@class=' cColSortedBG']/a[text()='Earnings Call']/ancestor::tr")
       
        for row in rows:
            
            try:
                checkbox = WebDriverWait(row, 20).until(EC.presence_of_element_located((By.XPATH,".//input[@type='checkbox']")))
                # row.find_element(By.XPATH,".//input[@type='checkbox']")
                if checkbox.is_enabled() and checkbox.get_attribute('id') != '_criteria__searchSection__searchToggle__includeScheduled':
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(checkbox)).click()
                    print("Checkbox clicked.")
            except Exception as e:
                print("Error clicking checkbox: {e}")
                quit.driver()

        if len(driver.find_elements(By.XPATH,"//td[@class=' cColSortedBG']/a[text()='Earnings Call']/ancestor::tr"))==0:
            print("No Earnings Call found")
        else:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "_transcriptsGrid__batchPrintButton_ctl00_MenuButton"))).click()
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='_transcriptsGrid__batchPrintButton_ctl00_MenuPanel_0']/a[4]"))).click()
        
        i += 1
        ii += 1
    
        if i%5==0:
            WebDriverWait(driver, 180).until(EC.invisibility_of_element_located((By.ID, "loadingMsg")))
            click_through_pages(i)
        elif ii%5==0:
            WebDriverWait(driver, 180).until(EC.invisibility_of_element_located((By.ID, "loadingMsg")))
            click_through_odd_pages(ii)
        else:
            WebDriverWait(driver, 180).until(EC.invisibility_of_element_located((By.ID, "loadingMsg")))
            click_through_pages(i)
    except: 
        print("error")
        break
