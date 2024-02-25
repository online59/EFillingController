import logging
import os
import time
import urllib.request
from logging.handlers import RotatingFileHandler
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import (ElementClickInterceptedException, StaleElementReferenceException, NoSuchElementException)
import pandas as pd

# Constants
MAX_ATTEMPTS = 3
WAIT_TIMEOUT = 10

class EFillingController():

    def read_credentials_from_excel(file_path):
        credentials = []
        df = pd.read_excel(file_path, dtype=str)
        for index, row in df.iterrows():
            username = row['username']
            password = row['password']
            credentials.append({'username': username, 'password': password, 'row': index + 1})
        return credentials
    
    def read_filter_options_from_excel(file_path):
        options = []
        df = pd.read_excel(file_path, dtype=str)
        for index, row in df.iterrows():
            option1 = row[0]
            option2 = row[1]
            option3 = row[2]
            option4 = row[3]
            option5 = row[4]
            option6 = row[5]
            option7 = row[6]
            options.append({'1': option1, '2': option2, '3': option3, '4': option4, '5': option5, '6': option6, '7': option7})
        return options
    
    def read_download_directory_from_excel(file_path):
        df = pd.read_excel(file_path, dtype=str)
        return df['directory'][0]

    def setup_debug_logging():
        # Create logger
        logger = logging.getLogger()
        logger.setLevel(logging.INFO)

        # Create console handler and set level to INFO
        # console_handler = logging.StreamHandler()
        # console_handler.setLevel(logging.INFO)

        # Create file handler and set level to DEBUG
        file_handler = logging.FileHandler('activitylog.txt', mode='w')  # Change mode to 'w'
        file_handler.setLevel(logging.INFO)

        # Create formatter
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

        # Add formatter to console handler and file handler
        # console_handler.setFormatter(formatter)
        file_handler.setFormatter(formatter)

        # Add console handler and file handler to logger
        # logger.addHandler(console_handler)
        logger.addHandler(file_handler)

    # Setup debug log file
    setup_debug_logging()

    def retry_function(func, *args, **kwargs):
        """Retry function with a maximum number of attempts."""
        for _ in range(MAX_ATTEMPTS):
            try:
                logging.info(f"Function: {func} performing")
                return func(*args, **kwargs)
            except Exception as e:
                logging.warning(f"Encountered exception: {e}, retrying...")
        raise Exception("Function failed after multiple attempts")
        
    def find_element_with_retry(driver, locator):
        """Find element matching the locator with a retry mechanism."""
        logging.info("Finding element with retry")
        return retry_function(WebDriverWait(driver, WAIT_TIMEOUT).until, EC.visibility_of_element_located(locator))

    def find_all_elements_with_retry(driver, locator):
        """Find all elements matching the locator with a retry mechanism."""
        logging.info("Finding all elements with retry")
        return retry_function(WebDriverWait(driver, WAIT_TIMEOUT).until, EC.visibility_of_all_elements_located(locator))

    def find_clickable_with_retry(driver, locator):
        """Find a clickable element matching the locator with a retry mechanism."""
        logging.info("Finding clickable element with retry")
        return retry_function(WebDriverWait(driver, WAIT_TIMEOUT).until, EC.element_to_be_clickable(locator))

    def click_element_with_retry(driver, element, fallback_locator=None, index=0):
        """
        Click on an element with a retry mechanism.

        Args:
            driver: Selenium WebDriver instance.
            element: Element to click.
            fallback_locator: Locator to use as a fallback if the original element is stale.

        Returns:
            None
        """
        for attempt in range(1, MAX_ATTEMPTS + 1):
            try:
                logging.info(f'Attempt {attempt}: Clicking on {element}')
                element.click()
                logging.info(f'Click successful on attempt {attempt}')
                return
            except ElementClickInterceptedException as intercepted_e:
                logging.warning(f"Attempt {attempt}: Click failed due to intercepted element: {intercepted_e}")
                logging.warning(f"Attempt {attempt}: Trying click on button with JavaScript on element: {element}")
                try:
                    driver.execute_script("arguments[0].click();", element)
                    logging.info(f"Attempt {attempt}: Success clicking button with JavaScript")
                    return
                except Exception as js_e:
                    logging.warning(f"Attempt {attempt}: Failed clicking button with JavaScript: {e}")
                    continue  # Retry finding elements
            except StaleElementReferenceException as stale_e:
                logging.warning(f"Click failed due to stale element: {stale_e}")
                if fallback_locator is None:
                    logging.warning(f"No fallback locator set, skip trying")
                    continue
                try:
                    logging.warning(f"Attempt {attempt}: Trying to locate an element")
                    element = WebDriverWait(driver, WAIT_TIMEOUT).until(EC.visibility_of_all_elements_located(fallback_locator))[index]
                    logging.info("Trying to click on a newly located element")
                    click_element_with_retry(driver, element)
                    logging.info(f"Attempt {attempt}: Success waiting for element to be clickable and clicking")
                    return
                except Exception as click_e:
                    logging.warning(f"Attempt {attempt}: Cannot click on an element: {click_e}")
                    continue  # Retry finding elements
            except NoSuchElementException as no_ele_e:
                logging.warning(f"Click failed due to no element: {no_ele_e}")
                logging.info("Trying to locate element after element presented")
                element = WebDriverWait(driver, WAIT_TIMEOUT).until(EC.visibility_of_all_elements_located(fallback_locator))[index]
                try:
                    logging.info("Trying to click on a newly located element")
                    click_element_with_retry(driver, element)
                    logging.info(f"Attempt {attempt}: Success waiting for element to be clickable and clicking")
                except Exception as click_e:
                    logging.warning(f"Attempt {attempt}: Cannot click on an element: {click_e}")
                    continue  # Retry finding elements
            except Exception as e:
                logging.error(f"Attempt {attempt}: Waiting for element to be clickable and clicking failed: {e}")
                move_element_to_viewport(driver, element)
                continue  # Retry finding elements
        raise RuntimeError(f"Failed to click element after {MAX_ATTEMPTS} attempts")

    def is_element_in_viewport(driver, element):
        """
        Check if an element is within the viewport.

        Args:
            driver: Selenium WebDriver instance.
            element: WebElement.

        Returns:
            bool: True if element is in the viewport, False otherwise.
        """
        for _ in range(MAX_ATTEMPTS):  # Retry up to 3 times
            try:
                logging.info("Try moving to element with Selenium")
                ActionChains(driver).move_to_element(element).perform()
                logging.info("Successful moving to element")
                return element.is_displayed()
            except Exception as e:
                logging.warning("Cannot move to element by Selenium, try using JavaScript...")
                try:
                    logging.info("Try moving to element with JavaScript")
                    return driver.execute_script("""
                        var elem = arguments[0];
                        var bounding = elem.getBoundingClientRect();
                        return (
                            bounding.top >= 0 &&
                            bounding.left >= 0 &&
                            bounding.bottom <= (window.innerHeight || document.documentElement.clientHeight) &&
                            bounding.right <= (window.innerWidth || document.documentElement.clientWidth)
                        );
                    """, element)
                except Exception as js_e:
                    logging.warning(f"Encountered exception: {js_e}, retrying...")
                    continue  # Retry checking element visibility
        logging.error("Failed to check element visibility after multiple attempts")
        return False  # Return False if unable to determine visibility

    def move_element_to_viewport(driver, element):
        """
        Move the element to the viewport if it's not already there.

        Args:
            driver: Selenium WebDriver instance.
            element: WebElement.

        Returns:
            None
        """
        if not is_element_in_viewport(driver, element):
            try:
                logging.info("Try moving to element with Selenium")
                ActionChains(driver).move_to_element(element).perform()
                logging.info("Successful moving to element to viewport")
                return True
            except Exception as sel_e:
                logging.error("Cannot scroll to the element using Selenium", sel_e)
                try:
                    logging.info("Try moving to element with JavaScript")
                    driver.execute_script("arguments[0].scrollIntoView(true);", element)
                    logging.info("Successful moving to element to viewport")
                except Exception as js_e:
                    logging.error("Cannot scroll to the element using JavaScript", js_e)

    def login(username, password, login_url):
        """
        Login to the website.

        Args:
            username: Username for login.
            password: Password for login.
            login_url: URL for login page.

        Returns:
            WebDriver instance after successful login.
        """
        logging.info("Logging in...")
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
        driver.get(login_url)
        try:
            username_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'username')))
            username_field.send_keys(username)
            password_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'passwordField')))
            password_field.send_keys(password)
            password_field.send_keys(Keys.RETURN)
            WebDriverWait(driver, 10).until(EC.title_is('ยื่นแบบ'))
            logging.info("Login successful")
            return driver
        except Exception as e:
            logging.error("An error occurred during login: %s", e)
            return None

    def navigate_to_pdf_page(driver):
        logging.info("Navigating to all tax form page...")
        retry_function(driver.get, 'https://efiling.rd.go.th/rd-efiling-web/form-status')

    def open_filter_panel(driver):
        logging.info("Opening filter panel...")
        try:
            filter_button = find_element_with_retry(driver, (By.XPATH, "//div[@class='collapsed' and @aria-expanded='true']"))
            click_element_with_retry(driver, filter_button)
        except Exception as e:
            logging.error(f"Failed to open filter panel: {e}")

    def select_dropdown_item(driver, form, select_item):
        logging.info(f"Selecting '{select_item}' from dropdown menu...")
        try:
            dropdown_button = find_element_with_retry(driver, (By.CSS_SELECTOR, f"ng-select[formcontrolname='{form}']"))
            click_element_with_retry(driver, dropdown_button)
        except Exception as e:
            logging.error(f"Failed to open dropdown menu: {e}")

        try:
            select_item_button = find_element_with_retry(driver, (By.XPATH, f"//span[@class='ng-option-label ng-star-inserted' and contains(text(), '{select_item}')]"))
            click_element_with_retry(driver, select_item_button)
        except Exception as e:
            logging.error(f"Failed to select item from dropdown: {e}")

    def input_item(driver, form, input_item):
        logging.info(f"Inputting '{input_item}' into form...")
        try:
            input_element = find_element_with_retry(driver, (By.XPATH, f"//input[@formcontrolname='{form}']"))
            input_element.send_keys(input_item)
        except Exception as e:
            logging.error(f"Failed to input item: {e}")

    def fill_form(driver, filter_form):
        logging.info("Filling filter form...")
        for filter in filter_form:
            form_type = filter['type']
            form = filter['form']
            input_item = filter['item']

            if input_item is None:
                continue

            if form_type == 'dropdown':
                select_dropdown_item(driver, form, input_item)
            elif form_type == 'input':
                input_item(driver, form, input_item)
            else:
                continue
        
        # Click search button
        try:
            search_button = find_element_with_retry(driver, (By.XPATH, "//button[@type='submit']"))
            click_element_with_retry(driver, search_button)
        except Exception as e:
            logging.error(f"Failed to click search button: {e}")

    def convert_thai_month_to_eng(tax_month):
        """Convert Thai month abbreviation to English."""
        logging.info("Converting Thai's month into Eng")
        thai_to_eng_month = {
            "ม.ค.": "JAN",
            "ก.พ.": "FEB",
            "มี.ค.": "MAR",
            "เม.ย.": "APR",
            "พ.ค.": "MAY",
            "มิ.ย.": "JUN",
            "ก.ค.": "JUL",
            "ส.ค.": "AUG",
            "ก.ย.": "SEP",
            "ต.ค.": "OCT",
            "พ.ย.": "NOV",
            "ธ.ค.": "DEC"
        }

        # Check if the provided tax_month is None
        if tax_month is None:
            return "MONTH"

        # Check if the provided tax_month is in the list of prohibited names
        if tax_month in thai_to_eng_month:
            logging.info("Converting Thai's month into Eng")
            return thai_to_eng_month[tax_month]
        else:
            logging.info("Tax month not found in the list")
            return tax_month

    def convert_thai_tax_form_to_eng(tax_form):
        """Convert Thai tax form abbreviation to English."""
        logging.info("Converting Thai's tax form into Eng")
        thai_to_eng_tax_form = {
            "ภ.ง.ด.1": "PND1",
            "ภ.ง.ด.2": "PND2",
            "ภ.ง.ด.3": "PND3",
            "ภ.ง.ด.53": "PND53",
            "ภ.ง.ด.90": "PND90",
            "ภ.ง.ด.91": "PND91",
            "ภ.ง.ด.93": "PND93",
            "ภ.ง.ด.94": "PND94",
            "ภ.ง.ด.96": "PND96",
            "ภ.ง.ด.97": "PND97",
            "ภ.ง.ด.98": "PND98",
            "ภ.ง.ด.99": "PND99",
            "ภ.ง.ด.1ก": "PND1KOR",
            "ภ.ง.ด.2ก": "PND2KOR",
            "ภ.ง.ด.3ก": "PND3KOR",
            "ภ.ง.ด.53ก": "PND53KOR",
            "ภ.ง.ด.90ก": "PND90KOR",
            "ภ.ง.ด.91ก": "PND91KOR",
            "ภ.ง.ด.91 (นายจ้างยื่นแทน)": "PND91KOR EMPLOYER",
            "ภ.ง.ด.93ก": "PND93KOR",
            "ภ.ง.ด.94ก": "PND94KOR",
            "ภ.ง.ด.96ก": "PND96KOR",
            "ภ.ง.ด.97ก": "PND97KOR",
            "ภ.ง.ด.98ก": "PND98KOR",
            "ภ.ง.ด.99ก": "PND99KOR",
            "ภ.พ.30": "PP30",
            "ภ.พ.36": "PP36",
            "ภ.ธ.40": "PP40",
            "งบการเงิน": "FinancialStatement",
            "แบบแจ้งเงินได้ต่างด้าว": "FORIEGNINCOME",
        }

        # Check if the provided tax_form is in the list of prohibited names
        if tax_form in thai_to_eng_tax_form:
            logging.info("Converting Thai's tax form into Eng")
            return thai_to_eng_tax_form[tax_form]
        else:
            logging.info("Tax form not found in the list")
            return tax_form

    def get_file_name(driver, filter_form, username, download_directory):
        """
        Constructing a file name with the URL.

        Args:
            driver: Selenium WebDriver instance.
            filter_form: Dictionary containing filter information.
            username: Username for the current user.
            download_directory: Directory where files will be downloaded.

        Returns:
            str: File name.
        """
        try:
            logging.info("Getting download filename...")
            tax_name = filter_form[0]['item']
            tax_year = filter_form[1]['item']
            tax_month = filter_form[2]['item']

            # Convert variables to strings before joining
            tax_name = str(convert_thai_tax_form_to_eng(tax_name))
            tax_year = str(tax_year)
            tax_month = str(convert_thai_month_to_eng(tax_month)).upper()

            try:
                current_url = driver.current_url
                url_extr = current_url.split('/')[-1]  # Extract filename from URL
            except Exception as e:
                logging.warning("Cannot extract name from URL or URL not found")
                url_extr = ""

            # Check if the URL is a TAX_FORM or RECEIPT
            if "RECEIPT" in url_extr:
                base_filename = f"RECEIPT_{tax_month}-{tax_year} {username}.pdf"
                logging.info(f"Base filename: {base_filename}")
            else:
                base_filename = f"{tax_name} {tax_month}-{tax_year} {username}.pdf"
                logging.info(f"Base filename: {base_filename}")

            # Check if the base filename already exists
            if os.path.exists(os.path.join(download_directory, base_filename)):
                logging.info("Base filename already exists, finding next available filename...")
                # Find the next available filename by adding a numbering suffix
                index = 1
                while True:
                    numbered_filename = f"{base_filename[:-4]} {index}.pdf"  # Append _{index} before the extension
                    if not os.path.exists(os.path.join(download_directory, numbered_filename)):
                        filename = numbered_filename
                        logging.info(f"Numbered filename: {filename}")
                        break
                    index += 1
            else:
                filename = base_filename
                logging.info(f"Final filename: {filename}")
            
            logging.info("Filename creation successful")
            return filename
        
        except Exception as e:
            logging.error("Error in get_file_name function: %s", e)
            return ""  # Return empty string or handle the error appropriately

    def get_default_download_folder():
        """Retrieve default download folder."""
        logging.info("Retrieving default download folder...")
        try:
            user_home = os.path.expanduser("~")
            default_download_folder = os.path.join(user_home, "Downloads")
            if os.path.exists(default_download_folder):
                logging.info("Default download folder retrieved successfully.")
                return default_download_folder
            else:
                logging.warning("Default download folder does not exist.")
                return None
        except Exception as e:
            logging.error("Error occurred while retrieving default download folder: %s", e)
            return None

    def download_pdf(driver, download_directory, filename=None):
        """
        Download PDF file into the designated folder.

        Args:
            driver: Selenium WebDriver instance.
            download_directory: Directory to save the downloaded PDF file.
            filename: Name of the downloaded file.

        Returns:
            None
        """
        logging.info("Attemp downloading PDF...")
        attempt = 0
        while attempt < MAX_ATTEMPTS:
            logging.info(f"Download PDF on {attempt + 1} attempt")
            current_url = driver.current_url
            
            if (download_directory == ""):
                logging.info("Download directory not found, retrieving default download folder...")
                download_directory = get_default_download_folder()

            if filename is None:
                logging.info("Filename not found, using default name")
                filename = "download_file.pdf"

            try:
                logging.info(f"Creating destination directory at: {download_directory}")
                os.makedirs(download_directory, exist_ok=True)
                logging.info(f"Destination directory created successfully")
            except Exception as e:

                logging.warning(f"Failed to create destination directory (Attempt {attempt + 1}): {e}")
                attempt += 1
                continue  # Try another attempt

            try:
                logging.info("Join file name with OS path")
                saved_directory = os.path.join(download_directory, filename)
                logging.info(f"Filename joined successfully: {saved_directory}")
            except Exception as e:
                logging.warning("Error, filename incorrect")
                attempt += 1
                continue  # Try another attempt

            try:
                logging.info("Retrieve PDF URL for downloading")
                urllib.request.urlretrieve(current_url, saved_directory)
                logging.info(f"PDF downloaded successfully to: {saved_directory}")
                return  # Exit the function after successful download
            except Exception as e:
                logging.warning(f"Failed to download PDF: {e}")
                attempt += 1
                continue  # Try another attempt

            attempt += 1

        logging.error("Failed to download PDF after multiple attempts")

    def find_and_download_pdf(driver, filter_form, username, download_directory):
        """Find and download PDF."""
        logging.info("Finding and downloading PDF...")
        tax_name = filter_form[0]['item']
        tax_year = filter_form[1]['item']
        tax_month = filter_form[2]['item']

        tax_month = str(convert_thai_month_to_eng(tax_month))

        try:
            logging.info("Formatting destination path")
            destination = [username, tax_year, tax_month]
            destination = "/".join(destination)
            logging.info(f"Destination path formatted successfully: {destination}")
        except Exception as e:
            logging.error("Error formatting directory path: %s", e)  # Print specific error message

        try:
            logging.info("Constructing path to folder for download")
            final_directory = os.path.join(download_directory, destination)
            logging.info(f"Path constructed successfully: {final_directory}")
        except Exception as e:
            logging.error("Error joining destination folder: %s", e)  # Print specific error message

        last_clicked_index = 0
                
        while True:
            
            try:
                button_elements = find_all_elements_with_retry(driver, (By.XPATH, '//button[@aria-controls="dropdown-basic" and @id="button-basic"]'))
            except Exception as e:
                logging.error("Failed to find dropdown button: %s", e)
                break

            if last_clicked_index + 1 > len(button_elements):
                break

            try:
                click_element_with_retry(driver, button_elements[last_clicked_index], fallback_locator=(By.XPATH, '//button[@aria-controls="dropdown-basic" and @id="button-basic"]'), index=last_clicked_index)
            except Exception as e:
                logging.error("Failed to click on dropdown menu", e)
                break
            
            try:
                dropdown_menu = find_clickable_with_retry(driver, (By.XPATH, '//a[@class="dropdown-item" and contains(text(), "พิมพ์ภาพแบบ/ภาพใบเสร็จ")]'))
            except Exception as e:
                logging.error("Failed to find dropdown item: %s", e)
                break

            try:
                click_element_with_retry(driver, dropdown_menu)
            except Exception as e:
                logging.error("Failed to click on dropdown item", e)
                break

            try:
                download_buttons = find_all_elements_with_retry(driver, (By.XPATH, '//button[contains(text(), "ดาวน์โหลด")]'))
            except Exception as e:
                logging.error("Failed to find download buttons: %s", e)
                break

            for download_button in download_buttons:

                try:
                    click_element_with_retry(driver, download_button)
                except Exception as e:
                    logging.error("Failed to click on download button", e)
                    continue

                try:
                    logging.info("Switching to new tab")
                    tabs = driver.window_handles
                    driver.switch_to.window(tabs[-1])
                except Exception as e:
                    logging.error("Error switching to new tab: %s", e)
                    continue
                    
                try:
                    logging.info("Joining destination to filename")
                    filename = os.path.join(final_directory, get_file_name(driver, filter_form, username, final_directory))
                    logging.info(f"Filename joined successfully: {filename}")
                except Exception as e:
                    logging.error("Error joining destination to filename: %s", e)
                
                try:
                    download_pdf(driver, final_directory, filename=filename)
                except Exception as e:
                    logging.error("Error downloading PDF: %s", e)
                    driver.close() # Close the tab
                    return

                try:
                    driver.close()
                except Exception as e:
                    logging.error("Error closing tab: %s", e)
                    return

                try:
                    driver.switch_to.window(tabs[0])
                except Exception as e:
                    logging.error("Error switching to original tab: %s", e)
                    return

            try:
                close_button = find_clickable_with_retry(driver, (By.XPATH, '//button[contains(@class, "btn button-box button-box-close-modal") and contains(text(), "ปิด")]'))
            except Exception as e:
                logging.error("Failed to find close button: %s", e)
                return

            try:
                click_element_with_retry(driver, close_button)
            except Exception as e:
                logging.error("Failed to click on close button")
                return

            logging.info(f"Current button click counting: {last_clicked_index}")

            last_clicked_index += 1   

    def switch_to_next_page(driver):
        """Switch to the next page in the same URL."""
        logging.info("Switching to next page...")
        try:
            next_page_button = find_clickable_with_retry(driver, (By.XPATH, '//li[@title="หน้าถัดไป"]'))   
        except Exception as e:
            logging.error("Failed to switch to next page: %s", e)
            return False

        try:
            if "disabled" not in next_page_button.get_attribute("class"):
                logging.info(f'Next page class containing: {next_page_button.get_attribute("class")}')
                click_element_with_retry(driver, next_page_button)
                return True
            else:
                logging.info("No more pages to switch to")
                return False
        except Exception as e:
            logging.error("Failed to click next page: %s", e)
            return False

    def logout(driver):
        """Logout from the site."""
        logging.info("Logging out...")
        driver.quit()

    # Main controller
    def login_and_download_all_pdfs(username, password, login_url, filter_form, download_directory):
        driver = login(username, password, login_url)
        try:
            navigate_to_pdf_page(driver)
            
            # Open filter panel
            open_filter_panel(driver)

            # Fill filter form
            fill_form(driver, filter_form)

            # Wait for page to load
            time.sleep(2)
                
            # Download pdfs from every items shown in the page
            while True:
                find_and_download_pdf(driver, filter_form, username, download_directory)
                if (not switch_to_next_page(driver)):
                    break

            time.sleep(5)

        finally:
            logout(driver)

    # Main function
    def main():

        # Read credentials from Excel file
        accounts = read_credentials_from_excel("credentials.xlsx")

        # Read filter options from Excel file
        options = read_filter_options_from_excel("filter_options.xlsx")

        download_directory = read_download_directory_from_excel("download_directory.xlsx")

        # Set default download directory
        DEFAULT_DOWNLOAD_DIRECTORY = "C:/Users/Se_sA/Downloads/EFillingController"

        login_url = "https://efiling.rd.go.th/rd-efiling-web/login"

        filter_form = [
            {'form': 'taxForm', 'item': 'ภ.ง.ด.1', 'type': 'dropdown'},
            {'form': 'taxYear', 'item': '2566', 'type': 'dropdown'},
            {'form': 'taxMonth', 'item': 'ม.ค.', 'type': 'dropdown'},
            {'form': 'nid', 'item': None, 'type': 'input'},
            {'form': 'fullName', 'item': None, 'type': 'input'},
            {'form': 'refNo', 'item': None, 'type': 'input'},
            {'form': 'taxformStatus', 'item': None, 'type': 'dropdown'},
        ]

        for account in accounts:

            login_and_download_all_pdfs(account['username'], account['password'], login_url, filter_form, DEFAULT_DOWNLOAD_DIRECTORY)


if __name__ == "__main__":
    controller = EFillingController()
    controller.main()