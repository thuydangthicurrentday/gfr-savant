# Configuration for automation script
import os
from selenium.webdriver.common.by import By


RUN_GET_DOCUMENT_NOTE = True

# Excel file path containing client list
EXCEL_FILE_PATH = "download_gofileroom_data.xlsx"  # Change path as needed

# Web configuration
BASE_URL = "https://www.gofileroom.com/home.html"
LOGIN_URL = "https://www.gofileroom.com/login"  # If login is required

# Download configuration
# DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads")

# Selenium configuration
IMPLICIT_WAIT = 10
EXPLICIT_WAIT = 20
PAGE_LOAD_TIMEOUT = 30

# Retry configuration
MAX_RETRIES = 3
RETRY_DELAY = 2

# Try different selectors for search input
# SEARCH_INPUT_LOCATOR = (By.XPATH, "//input[@name='searchroleName' and @placeholder='Advanced Search']")
SEARCH_INPUT_ALTERNATIVES = [
    (By.XPATH, "//input[@name='searchroleName' and @placeholder='Advanced Search']"),
    (By.XPATH, "//input[@name='searchroleName']"),
    (By.XPATH, "//input[@placeholder='Advanced Search']"),
    (By.XPATH, "//input[@type='text' and @name='searchroleName']"),
    (By.CSS_SELECTOR, "input[name='searchroleName']"),
    (By.CSS_SELECTOR, "input[placeholder='Advanced Search']"),
    (By.CSS_SELECTOR, "input[type='text'][name='searchroleName']"),
]
# Locator for iframe containing search input
SEARCH_CLIENT_IFRAME_LOCATOR = (By.ID, "cabinetpage")
SEARCH_INPUT_LOCATOR = (By.XPATH, "//input[@placeholder='Advanced Search']")
CLIENT_COUNT_LINK_LOCATOR = (By.XPATH, "//gfr-explorer-tree[contains(@class, 'ng-star-inserted')]//a[@class='node-name' and contains(text(), 'CLIENTS(')][1]")

# First gfr-explorer-tree element wrapping Clients()
CLIENT_TREE_ROOT_LOCATOR = (By.XPATH, "(//gfr-explorer-tree)[1]")
DOCUMENT_HEADERS_LOCATOR= (By.XPATH, ".//div[contains(@class, 'wj-cell') and contains(@class, 'wj-header') and contains(@tabindex, '-1')]")
DOCUMENT_ACTION_BTNS_LOCALTOR = (By.XPATH, "//span[text()='Documents Actions']")
EXPORT_DOCUMENT_BTNS_LOCALTOR = (By.XPATH, "//button[.//span[text()='Export Document']]")
OK_BTN_LOCALTOR = (By.XPATH, "//button[text()='Ok' and contains(@class, 'okButton')]")

DOCUMENT_TABLE_LOCATOR = (By.XPATH, "//div[@class='wj-cells' and @wj-part='cells']")
DOCUMENT_TABLE_DIV_LOCATOR = (By.XPATH, ".//div[@class='wj-row' and @role='row']")
DOCUMENT_ROW_FIRST_CELL_LOCATOR= (By.XPATH, ".//div[contains(@class, 'wj-cell') and contains(@class, 'wj-frozen') and contains(@class, 'wj-frozen-col')]")

DOCUMENT_DATA_CELL_LOCATOR= (By.XPATH, ".//div[contains(@class, 'wj-cell') and @role='gridcell']")
NEXT_PAGE_BTN_LOCATOR= (By.XPATH, ".//div[contains(@class, 'paginate-button') and contains(@class, 'next')]")

# Locator for Notes button (located in document viewer iframe)
NOTES_BUTTON_LOCATOR = (By.XPATH, "//div[@class='docview-item' and contains(@onclick, 'opennotes')]")
NOTES_POPUP_WINDOW_TITLE = "NOTES - Google Chrome"
# Locator for note text in popup (text is in <b> tag within 5th <tr> of table)
NOTES_TEXT_LOCATOR = (By.XPATH, "//table/tbody/tr[5]/td[@colspan='5']/font/b")
NOTES_CLOSE_BUTTON_LOCATOR = (By.XPATH, "//input[@type='button' and @value='Close']")
