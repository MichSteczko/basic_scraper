from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.remote.webelement import WebElement
import os 
import pandas as pd
from bs4 import BeautifulSoup
import click
from utils.get_webdriver import scrape_versions, get_latest_version_link, extract_zip_file, download_web_driver, parse_json_versions, get_chrome_version, compare_versions
import time 
import logging 
DRIVER_SERVICE = Service(os.path.join(".", "web_drivers", "chromedriver-win32", "chromedriver.exe"))

LOGGER = logging.getLogger(__name__)
LOGGER.setLevel(logging.INFO)
CONSOLE_HANDLER = logging.StreamHandler()
FORMATTER = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
CONSOLE_HANDLER.setFormatter(FORMATTER)
LOGGER.addHandler(CONSOLE_HANDLER)


def get_driver():
    options = Options()
    #options.add_argument("--headless")
    options.add_argument("--window-size=1920,1200")
    driver = webdriver.Chrome(options=options, service=DRIVER_SERVICE)
    driver.get("https://pl.investing.com/economic-calendar/")
    #driver.implicitly_wait(10)
    WebDriverWait(driver, 2000).until(EC.element_to_be_clickable((By.XPATH, "//button[@id='onetrust-accept-btn-handler']"))).click()
    #WebDriverWait(driver, 2000).until(EC.element_to_be_clickable((By.XPATH, "//i[@class='popupCloseIcon largeBannerCloser']"))).click()
    return driver


def filter_dates(driver, start_date: str, end_date: str):
    calendar_icon = driver.find_element(By.XPATH, "//div[@class='float_lang_base_1 js-tabs-economic']//a[@class='newBtn toggleButton LightGray datePickerBtn noText datePickerIconWrap']")#.click()
    driver.execute_script("arguments[0].click()", calendar_icon)
    calendar_container_xpath = "//div[@class='ui-datepicker ui-widget ui-widget-content ui-helper-clearfix ui-corner-all ui-datepicker-multi ui-datepicker-multi-3']"
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, calendar_container_xpath+"//input[@id='startDate']"))).clear()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, calendar_container_xpath+"//input[@id='startDate']"))).send_keys(start_date)
    element = driver.find_element(By.XPATH, calendar_container_xpath+"//input[@id='startDate']")
    print(element.get_attribute('value'))
    assert element.get_attribute('value') == start_date
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, calendar_container_xpath+"//input[@id='endDate']"))).clear()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, calendar_container_xpath+"//input[@id='endDate']"))).send_keys(end_date)
    element2 = driver.find_element(By.XPATH, calendar_container_xpath+"//input[@id='endDate']")
    apply_btn = driver.find_element(By.XPATH, calendar_container_xpath+"//div[@class='ui-datepicker-buttonpane ui-widget-content']//a[@id='applyBtn']")
    print(element2.get_attribute('value'))
    assert element2.get_attribute('value') == end_date
    driver.execute_script("arguments[0].click()", apply_btn)
    return driver

def get_countries():
    path_to_country_file = os.path.join(".", "countries.txt")
    countries = []
    with open(path_to_country_file, 'r', encoding="utf-8") as countries_file:
        for country in countries_file:
            countries.append(country[1:country.find("\n")])
    return countries

def filter_countries(driver: webdriver.Chrome, wanted_countries: list):
    try:
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[@id='button_parent']")))
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//a[@id='filterStateAnchor']"))).click()
    except:
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[@id='TNB_Instrument']")))
        calendar_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//a[@id='filterStateAnchor']")))
        driver.execute_script("arguments[0].click()", calendar_element)

    filters_container_xpath = "//div[@id='filtersWrapper']"
    country_box_xpath = filters_container_xpath+"//div[@id='calendarFilterBox_country']"
    clear_button = driver.find_elements(By.XPATH, country_box_xpath+"//div[@class='left float_lang_base_1']//a")
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable(clear_button[1])).click()
    country_list = driver.find_elements(By.XPATH, country_box_xpath+"//ul[@class='countryOption']//li")
    for country_item in country_list:
        country_name = country_item.find_element(By.XPATH, ".//label").text
        if country_name in wanted_countries:
            country_checkbox = country_item.find_element(By.XPATH, ".//input")
            driver.execute_script("arguments[0].click()", country_checkbox)
    #print(driver.find_element(By.XPATH, filters_container_xpath+"//div[@class='ecoFilterBox submitBox align_center']//a").text)        
    driver.find_element(By.XPATH, filters_container_xpath+"//div[@class='ecoFilterBox submitBox align_center']//a").click()
    #driver.implicitly_wait(30)
    return driver

def get_columns(table_element: WebElement):
    table_headers = table_element.find_elements(By.XPATH, ".//thead//tr")[0]
    correct_table_headers = table_headers.find_elements(By.XPATH, ".//th")
    columns = []
    for table_header in correct_table_headers[:-1]:
        columns.append(table_header.text)
    return columns

def process_first_col(row):
    first_col_values = []
    for col in row:
        first_col_value = col.find_all("td")[0].text
        first_col_values.append(first_col_value)
    return first_col_values

def process_second_col(row):
    second_col_values = []
    for col in row:
        try:
            second_col_value = col.find_all("td")[1].text[1:]
            second_col_values.append(second_col_value)
        except IndexError:
            second_col_values.append("")
    return second_col_values

def process_third_column(row):
    third_col_values = []
    for col in row:
        try:
            third_col_value = col.find_all("td")[2].find("span").text
            third_col_values.append(third_col_value)
        
        except IndexError:
            third_col_values.append("")
        
        except AttributeError:
            third_col_value = col.find_all("td")[2].find_all("i")
            stars_num = len([star["class"] for star in third_col_value if "grayFullBullishIcon" in star["class"]])
            third_col_values.append(stars_num)
    return third_col_values

def process_fourth_column(row):
    fourth_col_values = []
    for col in row:
        try:
            fourth_col_value = col.find_all("td")[3].find("a").text
            fourth_col_values.append(fourth_col_value)
        
        except IndexError:
            fourth_col_values.append("")
        
        except AttributeError:
            fourth_col_value = col.find_all("td")[3].text
            fourth_col_values.append(fourth_col_value)
        
    return fourth_col_values

def process_fifth_column(row):
    fifth_col_values = []
    for col in row:
        try:
            fifth_col_value = col.find_all("td")[4].text
            fifth_col_values.append(fifth_col_value)
        
        except IndexError:
            fifth_col_values.append("")
        
        
    return fifth_col_values

def process_sixth_column(row):
    fifth_col_values = []
    for col in row:
        try:
            fifth_col_value = col.find_all("td")[5].text
            fifth_col_values.append(fifth_col_value)
        
        except IndexError:
            fifth_col_values.append("")
        
        
    return fifth_col_values

def process_seventh_column(row):
    fifth_col_values = []
    for col in row:
        try:
            fifth_col_value = col.find_all("td")[6].text
            fifth_col_values.append(fifth_col_value)
        
        except IndexError:
            fifth_col_values.append("")
        
        
    return fifth_col_values

def is_element_visible(driver, xpath):
    wait1 = WebDriverWait(driver, 2)
    try:
        wait1.until(EC.visibility_of_element_located((By.XPATH, xpath)))
        return True
    except Exception:
        return False

def get_table(date1, date2):
    driver = get_driver()
    date_driver = filter_dates(driver, date1, date2)
    date_driver.implicitly_wait(10)
    wanted = get_countries()
    country_driver = filter_countries(date_driver, wanted)
    country_driver.implicitly_wait(10)
    scroll_range = 10
    current_scroll_num = 0
    while current_scroll_num < scroll_range:
        country_driver.execute_script("window.scrollTo(0, window.scrollY + 1500)")
        time.sleep(5)
        current_scroll_num += 1
    table_element = WebDriverWait(country_driver, timeout=30).until(EC.element_to_be_clickable((By.XPATH, "//table[@id='economicCalendarData']")))
    columns = get_columns(table_element)
    table_html = table_element.get_attribute('innerHTML')
    driver.quit()
    raw = BeautifulSoup(table_html, "html.parser").find("tbody").find_all("tr")
    first = process_first_col(raw)
    second = process_second_col(raw)
    third = process_third_column(raw)
    fourth = process_fourth_column(raw)
    fifth = process_fifth_column(raw)
    sixth = process_sixth_column(raw)
    seventh = process_seventh_column(raw)
    all_cols = zip(first, second, third, fourth, fifth, sixth, seventh)
    data = [[col1, col2, col3, col4, col5, col6, col7] for col1, col2, col3, col4, col5, col6, col7 in all_cols]
    return pd.DataFrame(data=data, columns=columns)
    
def save_as_excel(date1, date2):
    dataframe_to_save = get_table(date1, date2)    
    writer = pd.ExcelWriter("excel_files/investing_kalendarz.xlsx", engine="xlsxwriter")
    dataframe_to_save.to_excel(writer, sheet_name="Kalendarz", index=False)
    writer._save()

@click.command()
@click.option('--date1',)
@click.option('--date2',)
def run_calendar(date1, date2):
    save_as_excel(date1, date2)


if __name__ == "__main__":
    try:
        versions = scrape_versions()
        latest_version_link = get_latest_version_link(versions, method="scrape")
    except ValueError:
        LOGGER.info("Method has been changed to json")
        versions = parse_json_versions("versions.json")
        latest_version_link = get_latest_version_link(versions, method="json")
        print(latest_version_link)
    download_web_driver(latest_version_link)
    extract_zip_file()
    run_calendar()
    


