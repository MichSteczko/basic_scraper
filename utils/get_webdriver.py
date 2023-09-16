import winreg as wrg
from pathlib import Path
from win32com.client import Dispatch 
from bs4 import BeautifulSoup 
from requests.sessions import Session
import re 
import logging
import zipfile
import os
import json

LOGGER = logging.getLogger(__name__)
LOGGER.setLevel(logging.INFO)
CONSOLE_HANDLER = logging.StreamHandler()
FORMATTER = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
CONSOLE_HANDLER.setFormatter(FORMATTER)
LOGGER.addHandler(CONSOLE_HANDLER)

def get_chrome_version():
    parser = Dispatch("Scripting.FileSystemObject")
    path_to_google = Path("SOFTWARE\\Clients\\StartMenuInternet\\Google Chrome\\Capabilities")
    location = wrg.HKEY_LOCAL_MACHINE
    google_key_content = wrg.OpenKeyEx(location, str(path_to_google))
    google_exe_file_location = wrg.QueryValueEx(google_key_content,"ApplicationIcon")[0][:-2]
    version = parser.GetFileVersion(google_exe_file_location)
    LOGGER.info(f"Your current chrome version is {version} 🥵")
    return version

def get_link_content(link):
    return link.find("span").text

def get_link(link):
    return link["href"]

def compare_versions(exact_version, version_to_find, method):
    pattern = r'\d+(\.\d+)*\.'
    if method == "scrape":
        version = version_to_find
    elif method == "json":
        version = version_to_find["version"]
    exact_version_trimmed = re.search(pattern, exact_version).group(0)
    version_to_find_trimmed = re.search(pattern, version).group(0)
    return exact_version_trimmed == version_to_find_trimmed


def scrape_versions():
    session = Session()
    versions = {}
    with session as s:
        distributions_website = s.get("https://chromedriver.chromium.org/downloads")
        html_object = BeautifulSoup(distributions_website.text, "html.parser")
        extracted_links = html_object.find_all("a", {"class": "XqQF9c"}, href=True)
        for link in extracted_links:
            link_content = get_link_content(link)
            potential_distribution_version = re.search(r"\d+\.\d+\.\d+\.\d+", link_content)
            if potential_distribution_version:
                distribution_link = get_link(link)
                distribution_version = potential_distribution_version.group(0)
                versions.update({distribution_version: distribution_link})
    return versions

def get_version_suffix(matched_version: str):
    last_dot_index = matched_version.rfind(".")
    return int(matched_version[last_dot_index+1:])


def get_latest_version(matched_versions: list, method: str):
    latest = 0
    latest_index = 0
    for index, matched_version in enumerate(matched_versions):
        if method == "json":
           current_version = get_version_suffix(matched_version["version"])
        elif method == "scrape":
            current_version = get_version_suffix(matched_version)
        if current_version > latest:
            latest = current_version
            latest_index = index
    try:
        return matched_versions[latest_index]
    except IndexError:
        return None

def get_latest_scraped_link(latest_version):
    return f"https://chromedriver.storage.googleapis.com/{latest_version}/chromedriver_win32.zip"

def get_latest_json_link(latest_version):
    webdriver_link = None
    for option in latest_version["downloads"]["chromedriver"]:
        if option["platform"] == "win32":
            webdriver_link = option["url"]
    return webdriver_link

def get_latest_version_link(versions, method: str):
    my_chrome_version = get_chrome_version()
    latest_versions = [version for version in versions if compare_versions(my_chrome_version, version, method)]
    latest_version = get_latest_version(latest_versions, method)
    if latest_version:
        if method == "scrape":
            return get_latest_scraped_link(latest_version)
        if method == "json":
            return get_latest_json_link(latest_version)
    raise ValueError

def download_web_driver(latest_version_link):
    session = Session()
    with session as s:
        donloaded_file = s.get(latest_version_link)
        with open("chromedriver_win32.zip", "wb") as f: 
            f.write(donloaded_file.content)
    LOGGER.info("Webdriver has been successfully downloaded! 🥰")

def extract_zip_file():
    with zipfile.ZipFile("chromedriver_win32.zip", "r") as zip_ref:
        for file in zip_ref.namelist():
            if file.endswith(".exe"):
                zip_ref.extract(file, "web_drivers")
    os.remove("chromedriver_win32.zip")
    LOGGER.info("Check web_drivers directory to see your latest file 🥴🥴🥴")

def parse_json_versions(path_to_json: str):
    with open(path_to_json, "r") as json_file:
        json_versions = json_file.read()
    return json.loads(json_versions)["versions"]


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