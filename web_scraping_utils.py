from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import os
import shutil

def get_chrome_driver_path():
    driver_path = ChromeDriverManager().install()
    return driver_path

def check_and_get_chrome_driver():
    driver_path = get_chrome_driver_path()
    if os.path.exists(driver_path):
        return driver_path
    else:
        return ChromeDriverManager().install()


def reset_chrome_driver_cache():
    # Clear the ChromeDriver cache directory
    cache_dir = os.path.expanduser("~/.wdm/drivers/")
    if os.path.exists(cache_dir):
        for root, dirs, files in os.walk(cache_dir, topdown=False):
            for name in files:
                os.remove(os.path.join(root, name))
            for name in dirs:
                os.rmdir(os.path.join(root, name))
        os.rmdir(cache_dir)
    print("ChromeDriver cache cleared.")

reset_chrome_driver_cache()

driver_path = check_and_get_chrome_driver()
driver = webdriver.Chrome(service=Service(driver_path))