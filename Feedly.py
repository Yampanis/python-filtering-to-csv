import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import undetected_chromedriver as uc
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import requests
import os
import pickle
from dotenv import load_dotenv
import json
import datetime
import pyperclip
import re
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from selectolax.parser import HTMLParser
from urllib.request import unquote, quote
from urllib.parse import urlparse
import pandas as pd

load_dotenv()
OPEN_API_KEY = os.getenv("OPEN_API_KEY")

def infinite_scroll(driver, max_scrolls=5):
    """Scrolls the page down multiple times to load all content."""
    for _ in range(max_scrolls):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)  

def cleanup_cookies():
    try:
        if os.path.exists("cookies"):
            for file in os.listdir("cookies"):
                if file.endswith(".pkl"):
                    os.remove(os.path.join("cookies", file))
            print("Cleaned up existing cookies")
    except Exception as e:
        print(f"Error cleaning cookies: {e}")

def save_cookies(driver, path):
    """Save cookies to a file"""
    if not os.path.exists('cookies'):
        os.makedirs('cookies')
    with open(path, 'wb') as file:
        pickle.dump(driver.get_cookies(), file)

def load_cookies(driver, path):
    try:
        if os.path.exists(path) and os.path.getsize(path) > 0:
            with open(path, 'rb') as file:
                cookies = pickle.load(file)
                for cookie in cookies:
                    try:
                        driver.add_cookie(cookie)
                    except Exception as e:
                        print(f"Error adding cookie: {e}")
                        continue
            return True
        return False
    except (EOFError, pickle.UnpicklingError) as e:
        print(f"Error loading cookies: {e}")
        try:
            os.remove(path)
            print("Removed corrupted cookie file")
        except OSError:
            pass
        return False

def get_base64_str(source_url):
    try:
        url = urlparse(source_url)
        path = url.path.split("/")
        if url.hostname == "news.google.com" and len(path) > 1 and path[-2] in ["articles", "read"]:
            return {"status": True, "base64_str": path[-1]}
        return {"status": False, "message": "Invalid Google News URL format."}
    except Exception as e:
        return {"status": False, "message": f"Error in get_base64_str: {str(e)}"}

def get_decoding_params(base64_str):
    try:
        url = f"https://news.google.com/articles/{base64_str}"
        response = requests.get(url)
        response.raise_for_status()

        parser = HTMLParser(response.text)
        data_element = parser.css_first("c-wiz > div[jscontroller]")
        if data_element is None:
            return {"status": False, "message": "Failed to fetch data attributes from articles URL."}

        return {
            "status": True,
            "signature": data_element.attributes.get("data-n-a-sg"),
            "timestamp": data_element.attributes.get("data-n-a-ts"),
            "base64_str": base64_str,
        }

    except requests.exceptions.RequestException as req_err:
        try:
            url = f"https://news.google.com/rss/articles/{base64_str}"
            response = requests.get(url)
            response.raise_for_status()

            parser = HTMLParser(response.text)
            data_element = parser.css_first("c-wiz > div[jscontroller]")
            if data_element is None:
                return {"status": False, "message": "Failed to fetch data attributes from RSS URL."}

            return {
                "status": True,
                "signature": data_element.attributes.get("data-n-a-sg"),
                "timestamp": data_element.attributes.get("data-n-a-ts"),
                "base64_str": base64_str,
            }

        except requests.exceptions.RequestException as rss_req_err:
            return {"status": False, "message": f"Request error with RSS URL: {str(rss_req_err)}"}

    except Exception as e:
        return {"status": False, "message": f"Unexpected error in get_decoding_params: {str(e)}"}

def decode_url(signature, timestamp, base64_str):
    try:
        url = "https://news.google.com/_/DotsSplashUi/data/batchexecute"
        payload = [
            "Fbv4je",
            f'["garturlreq",[["X","X",["X","X"],null,null,1,1,"US:en",null,1,null,null,null,null,null,0,1],"X","X",1,[1,1,1],1,1,null,0,0,null,0],"{base64_str}",{timestamp},"{signature}"]',
        ]
        headers = {
            "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36",
        }

        response = requests.post(
            url,
            headers=headers,
            data=f"f.req={quote(json.dumps([[payload]]))}")
        response.raise_for_status()

        parsed_data = json.loads(response.text.split("\n\n")[1])[:-2]
        decoded_url = json.loads(parsed_data[0][2])[1]

        return {"status": True, "decoded_url": decoded_url}
    except requests.exceptions.RequestException as req_err:
        return {"status": False, "message": f"Request error in decode_url: {str(req_err)}"}
    except (json.JSONDecodeError, IndexError, TypeError) as parse_err:
        return {"status": False, "message": f"Parsing error in decode_url: {str(parse_err)}"}
    except Exception as e:
        return {"status": False, "message": f"Error in decode_url: {str(e)}"}

def decode_google_news_url(source_url, interval=None):
    try:
        base64_response = get_base64_str(source_url)
        if not base64_response["status"]:
            return base64_response

        decoding_params_response = get_decoding_params(base64_response["base64_str"])
        if not decoding_params_response["status"]:
            return decoding_params_response

        if interval:
            time.sleep(interval)

        decoded_url_response = decode_url(
            decoding_params_response["signature"],
            decoding_params_response["timestamp"],
            decoding_params_response["base64_str"],
        )

        return decoded_url_response
    except Exception as e:
        return {"status": False, "message": f"Error in decode_google_news_url: {str(e)}"}

##scopes = [
##    'https://www.googleapis.com/auth/spreadsheets',
##    'https://www.googleapis.com/auth/drive'
##]
##credentials_for_upload_sheet = ServiceAccountCredentials.from_json_keyfile_name(
##    r"C:\Users\the_b\Desktop\Feedly\credentials.json", scopes)
##file = gspread.authorize(credentials_for_upload_sheet)
##sheet3 = file.open("Rory Testing Sheet 2024 ")
##wks3 = sheet3.worksheet("NEW DATA")


try:
    titles_read_df = pd.read_excel(r'titles_to_check.xlsx', sheet_name='Sheet1')
    titles_read = titles_read_df['Titles'].tolist()
except Exception as ex:
    print('Error: ' + str(ex))
    titles_read = []
try:
    negative_titles_df = pd.read_excel(r'negative_titles.xlsx', sheet_name='Sheet1')
    negative_titles_read = negative_titles_df['Titles'].tolist()
except Exception as ex:
    print('Error negative titles: ' + str(ex))
    negative_titles_read = []
try:
    negative_keywords_df = pd.read_excel(r'negatives.xlsx', sheet_name='Sheet1')
    negative_keywords = negative_keywords_df['Negative'].tolist()
except Exception as ex:
    print('Error negatives: '+ str(ex))
    negative_keywords = []
# Convert negative keywords to a set for efficient lookups
negative_keywords_set = set(negative_keywords)

# Define today's date
today_str = datetime.datetime.now().strftime(
    "%a, %d %b %Y %H:%M:%S")  # Format matches article date
if today_str[12:14] == '24':
    today_str = today_str[:12] + '00' + today_str[14:]
new_today_str = datetime.datetime.strptime(today_str, "%a, %d %b %Y %H:%M:%S")
start_range = new_today_str - datetime.timedelta(hours=3)
# Setup Selenium WebDriver
chrome_options = Options()
# chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option('useAutomationExtension', False)
chrome_options.add_argument("--disable-bluetooth") 
chrome_options.add_argument("--log-level=3")
# Specify the path to chromedriver using Service
chromedriver_path = os.getenv("CHROMEDRIVER_PATH")  # Update with the correct path if necessary
service = Service(executable_path=chromedriver_path)

# Initialize WebDriver with the service and options
driver = webdriver.Chrome(service=service, options=chrome_options)

##options = Options()
##options.add_argument("--start-maximized")
##options.add_argument("--disable-blink-features=AutomationControlled")
##options.add_argument("--ignore-certificate-errors")
##
##driver = uc.Chrome(options=options, driver_executable_path=r'C:\Users\the_b\Desktop\Feedly\browser\chromedriver.exe', version_main=130)

def append_to_excel(existing_file, new_data, sheet_name):
    """Append data to Excel with formatting"""
    writer = pd.ExcelWriter(existing_file, engine='xlsxwriter')
    
    try:
        # Read existing data
        existing_data = pd.read_excel(existing_file, sheet_name=None)
        if sheet_name in existing_data:
            # Combine existing and new data
            combined_data = pd.concat([existing_data[sheet_name], new_data], ignore_index=True)
        else:
            combined_data = new_data
        # Write to Excel with formatting
        combined_data.to_excel(writer, sheet_name=sheet_name, index=False)
        # Get workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        # Add formats
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
        })
        # Write headers with format
        for col_num, value in enumerate(combined_data.columns.values):
            worksheet.write(0, col_num, value, header_format)
        # Auto-fit columns
        for idx, col in enumerate(combined_data.columns):
            series = combined_data[col]
            max_len = max(
                series.astype(str).apply(len).max(),
                len(str(series.name))
            ) + 1
            worksheet.set_column(idx, idx, max_len)
            
    except Exception as e:
        print(f"Error in append_to_excel: {e}")
    finally:
        writer.close()

# def adjust_column_width(workbook_path):
#     """Adjust column widths based on content"""
#     try:
#         # Load the workbook
#         workbook = load_workbook(workbook_path)
#         worksheet = workbook.active
        
#         # Iterate through columns
#         for column in worksheet.columns:
#             max_length = 0
#             column_letter = column[0].column_letter
            
#             # Find longest content in column
#             for cell in column:
#                 try:
#                     if len(str(cell.value)) > max_length:
#                         max_length = len(str(cell.value))
#                 except:
#                     pass
            
#             # Set width with some padding
#             adjusted_width = (max_length + 2)
#             worksheet.column_dimensions[column_letter].width = adjusted_width
        
#         # Save the workbook
#         workbook.save(workbook_path)
#         print(f"Adjusted column widths in {workbook_path}")
#     except Exception as e:
#         print(f"Error adjusting column widths: {e}")

# Function to check if a title contains any negative keywords
def contains_negative_keywords(title, keywords_set):
    # Split title into words and check intersection with negative keywords
    words_in_title = set(title.lower().split())
    return not words_in_title.isdisjoint(keywords_set)

def check_title_against_keywords(title, negative_keywords):
    """
    Check if a title contains any keyword (with word boundaries).
    """
    title_lower = title.lower()
    for keyword in negative_keywords:
        # Use word boundary (\b) regex for word matching
        keyword_pattern = r'\b' + re.escape(keyword.lower()) + r'\b'
        if re.search(keyword_pattern, title_lower):
            print(f"Title: {title_lower} contains keyword: {keyword_pattern}")
            return False
    return True

def url_contains_keyword(url, negative_keywords):
    """
    Check if a URL contains any keyword as a substring.
    Strips spaces and ensures proper matches.
    """
    url_lower = url.strip().lower()  # Normalize URL
    for keyword in negative_keywords:
        # Normalize keyword (strip spaces, lowercase)
        keyword_clean = keyword.lower()
        # Substring matching
        if keyword_clean in url_lower:
            print(f"Url: {url_lower} contains keyword: {keyword_clean}")
            return False
    return True


def feedly_login(driver, email, password):
    cookies_path = "cookies/feedly_cookies.pkl"
    driver.get("https://feedly.com/i/discover")
    driver.set_page_load_timeout(30)  
    
    time.sleep(3.5)

    cookie_login_successful = False
    try: 
        if load_cookies(driver, cookies_path):
            driver.refresh()
            try:
                WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.ID, "feedlyFrame"))
                )
                cookie_login_successful = True
                return True
            except Exception as e:
                print(f"Cookie login failed: {e}")
    except Exception as e:
        print(f"Error loading cookies: {e}")

    if not cookie_login_successful:    
        try:
            print ("Logging in with Gmail and password")
            login_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "(//button[contains(.,'Log In')])[1]"))
            )
            login_button.click()
            time.sleep(2)

            google_login = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(.,'with Google')]")
            ))
            google_login.click()

            email_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//input[@type='email']"))
            )
            email_input.send_keys(email)
            email_input.send_keys(Keys.ENTER)
            time.sleep(2)

            password_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//input[@type='password']"))
            )
            password_input.send_keys(password)
            password_input.send_keys(Keys.ENTER)
            time.sleep(10)
            # save_cookies(driver, cookies_path)
        except Exception as e:
            print(f"Error in feedly_login: {str(e)}")
            driver.quit()

def login_to_chatgpt_com(driver):
    try:
        login_button = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, '//button[@data-testid="login-button"]'))
        )
        login_button.click()
        time.sleep(5)
    except:
        print("Login button not found on chatgpt.com.")
        return
    # Attemp to login
    try:
        email = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, '//input[@id="email-input"]'))
        )#driver.find_element(By.XPATH, "//input[@id='email-input']")
        email.click()
        email.send_keys("rory@thebestreputation.com")
        email.send_keys(Keys.RETURN)
        time.sleep(5)
    except Exception as e:
        print("Unable to process username! Error: " + str(e))
        return
    # Attemp password
    try:
        passwd = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, '//input[@id="password"]'))
        )
        passwd.click()
        passwd.send_keys("Barbeque2045!")
        passwd.send_keys(Keys.RETURN)
        time.sleep(5)
    except Exception as e:
        print("Unable to process password to login! Error: " + str(e))
        return

def ask_chatgpt_com(brw, prompt):
    brw.get("https://chatgpt.com/")
    time.sleep(5)
    try:
        modal_cls = brw.find_element(By.XPATH, "//*[contains(@id,'radix-:r')]/div/div/a")
        modal_cls.close()
        time.sleep(2)
    except Exception as ex:
        pass
##    WebDriverWait(brw, 10).until(EC.presence_of_element_located((By.ID, 'prompt-textarea')))

    # Locate the input box and submit the prompt
    # Use JavaScript to reveal the hidden input box if necessary
    try:
        input_box = WebDriverWait(brw, 10).until(
            EC.element_to_be_clickable((By.ID, 'prompt-textarea'))
        )
        brw.execute_script("arguments[0].style.display = 'block';", input_box)  # Force display
        time.sleep(3)  # Allow JavaScript change to take effect
        input_box.clear()
        pyperclip.copy(prompt)

        # Click to focus the input box if necessary
        input_box.click()
        
        input_box.send_keys(Keys.CONTROL, 'v')
##        time.sleep(45)  # Wait for ChatGPT to generate a response
        input_box.send_keys(Keys.RETURN)
    except Exception as e:
        return f"Error: Unable to locate or interact with the input box. Details: {e}"

    # Wait until response is generated
    try:
        time.sleep(85)
    except Exception as e:
        return f"Error: Response generation timed out. Details: {e}"

    # Capture and return the response text
    try:
        response_elements = brw.find_elements(By.CSS_SELECTOR, ".markdown div.border-token-border-medium div.overflow-y-auto")
        if len(response_elements) > 1:
            response_text = response_elements[-1].get_attribute('outerText') if response_elements else "No response found"
        else:
            response_text = response_elements[0].get_attribute('outerText') if response_elements else "No response found"
    except Exception as ex:
        try:
            response_elements = brw.find_elements(By.CSS_SELECTOR, ".markdown p")
            resp_elems = ''
            if len(response_elements) > 1:
                for resel in resp_elems:
                    resp_elems += resel.get_attribute('outerText') + '\n'
                resp_elems = resp_elems.strip()
                response_text = resp_elems
            else:
                response_text = response_elements[0].get_attribute('outerText')
        except Exception as ex:
            print('Response text is empty!')
            response_text = ''
    if '#' in response_text:
        lines = response_text.strip().split('\n')
        # Initialize a list to hold the articles
        articles = []
        for line in lines:
            parts = line.split('#')  # Split by tab character
            if len(parts) >= 4:
                try:
                    try:
                        url = re.sub(r'\d. ','',parts[0].strip())
                    except:
                        url = parts[0].strip()
                except Exception as e:
                    url = 'N/A'
                try:
                    title = parts[1].strip()
                except Exception as e:
                    title = 'N/A'
                try:
                    description = parts[2].strip()
                except Exception as e:
                    description = 'N/A'
                try:
                    reach_out = parts[3].strip()
                except Exception as e:
                    reach_out = 'N/A'
                try:
                    reasons = parts[4].strip()
                except Exception as e:
                    reasons = 'N/A'
                try:
                    keywords = parts[5].strip()  # All keywords
                except Exception as e:
                    keywords = 'N/A'
                try:
                    if '|' in parts[6]:
                        location = re.sub(r'\|',', ', parts[6])
                    else:
                        location = parts[6].strip()
                except Exception as e:
                    location = 'N/A'

                # Create a dictionary for the article
                article = {
                    "url": url,
                    "title": title,
                    "description": description,
                    "reach_out": reach_out,
                    "reasons": reasons,
                    "keywords": keywords,
                    "location": location,
                }

                # Add the article dictionary to the list
                articles.append(article)
            # Convert the list of articles to JSON format
        json_output = json.dumps(articles, indent=4)
        return json_output
    elif '\t' in response_text:
        lines = response_text.strip().split('\n')
        # Initialize a list to hold the articles
        articles = []
        for line in lines:
            parts = line.split('\t')  # Split by tab character
            if len(parts) >= 4:
                try:
                    try:
                        url = re.sub(r'\d. ','',parts[0].strip())
                    except:
                        url = parts[0].strip()
                except Exception as e:
                    url = 'N/A'
                try:
                    title = parts[1].strip()
                except Exception as e:
                    title = 'N/A'
                try:
                    description = parts[2].strip()
                except Exception as e:
                    description = 'N/A'
                try:
                    reach_out = parts[3].strip()
                except Exception as e:
                    reach_out = 'N/A'
                try:
                    reasons = parts[4].strip()
                except Exception as e:
                    reasons = 'N/A'
                try:
                    keywords = parts[5].strip()  # All keywords
                except Exception as e:
                    keywords = 'N/A'
                try:
                    if '|' in parts[6]:
                        location = re.sub(r'\|',', ', parts[6])
                    else:
                        location = parts[6].strip()
                except Exception as e:
                    location = 'N/A'

                # Create a dictionary for the article
                article = {
                    "url": url,
                    "title": title,
                    "description": description,
                    "reach_out": reach_out,
                    "reasons": reasons,
                    "keywords": keywords,
                    "location": location,
                }

                # Add the article dictionary to the list
                articles.append(article)
        # Convert the list of articles to JSON format
        json_output = json.dumps(articles, indent=4)
        return json_output
    else:
        print("Unexpected response structure:", response_text)
        return None
##    except Exception as e:
##        return f"Error: Unable to retrieve response. Details: {e}"

def call_chatgpt_api(prompt, api_key):
    url = "https://api.openai.com/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    # Dynamically set max_tokens to stay within model limits
    prompt_tokens = len(prompt.split())
    max_tokens_response = min(4096 - prompt_tokens, 3700)
    data = {
        "model": "gpt-4",  # or use the model of your choice
        "messages": [{
            "role": "user",
            "content": prompt
        }],
        "max_tokens": 100  # adjust the number of tokens as needed
    }
    new_data = ''

    max_retries = 5
    for attempt in range(max_retries):
        if new_data == '':
            response = requests.post(url, headers=headers, json=data)
        else:
            response = requests.post(url, headers=headers, json=new_data)
        time.sleep(7)

        print(f"Attempt {attempt + 1}: Status Code: {response.status_code}"
              )  # Log the status code

        if response.status_code == 429:  # Rate limit exceeded
            wait_time = 2**attempt  # Exponential backoff
            print(f"Rate limit exceeded. Waiting for {wait_time} seconds...")
            time.sleep(wait_time)
            continue  # Retry the request
        elif response.status_code == 200:  # Successful response
            try:
                response_json = response.json()
                # Ensure the expected structure is present
                if 'choices' in response_json and len(
                        response_json['choices']) > 0:
                    content = response_json['choices'][0]['message']['content']
                    if not content.strip():  # Check if content is empty
                        print("Received empty response, retrying...")
                        continue  # Retry if the response is empty
                    # Split the raw data into individual lines
                    lines = content.strip().split('\n')
                    # Initialize a list to hold the articles
                    articles = []

                    # Process each line to extract data
                    for line in lines:
                        parts = line.split('#')  # Split by tab character
                        if len(parts) >= 4:
                            try:
                                try:
                                    url = re.sub(r'\d. ','',parts[0].strip())
                                except:
                                    url = parts[0].strip()
                            except Exception as e:
                                url = '-'
                            try:
                                title = parts[1].strip()
                            except Exception as e:
                                title = '-'
                            try:
                                description = parts[2].strip()
                            except Exception as e:
                                description = '-'
                            try:
                                reach_out = parts[3].strip()
                            except Exception as e:
                                reach_out = '-'
                            try:
                                reasons = parts[4].strip()
                            except Exception as e:
                                reasons = '-'
                            try:
                                keywords = parts[5].strip()  # All keywords
                            except Exception as e:
                                keywords = '-'
                            try:
                                location = parts[6].strip()
                            except Exception as e:
                                location = '-'

                            # Create a dictionary for the article
                            article = {
                                "url": url,
                                "title": title,
                                "description": description,
                                "reach_out": reach_out,
                                "reasons": reasons,
                                "keywords": keywords,
                                "location": location,
                            }

                            # Add the article dictionary to the list
                            articles.append(article)
                    # Convert the list of articles to JSON format
                    json_output = json.dumps(articles, indent=4)
                    return json_output
                else:
                    print("Unexpected response structure:", response_json)
                    return None
            except json.JSONDecodeError:
                print("Failed to decode JSON response:", response.text)
                return None
        else:
            try:
                new_data = {
                    "model": "gpt-3.5-turbo",  # or use the model of your choice
                    "messages": [{
                        "role": "user",
                        "content": prompt
                    }],
                    "max_tokens": max_tokens_response  # adjust the number of tokens as needed
                }
                continue
            except Exception as e:
                print(
                    f"Unexpected status code: {response.status_code}, response: {response.text}"
                )
                response.raise_for_status(
                )  # Raise an error for other bad responses


def scroll_down(driver, element_selector):
    try:
        # Find the element (e.g., <main> tag or another container) and scroll within it
        main_element = WebDriverWait(driver, 5).until(
            EC.visibility_of_element_located((By.XPATH, element_selector))
        )
        # Scroll to the bottom of the element
        driver.execute_script(
            "arguments[0].scrollTo(0, arguments[0].scrollHeight);", main_element)
        time.sleep(1.5)
    except Exception as e:
        pass


def scrape_today_articles(driver):
    new_articles = []
    batch_size = 500

    last_height = driver.execute_script(
        "return document.querySelector('#feedlyFrame').scrollHeight")
    while True:
        # Scroll down and wait for page load
        scroll_down(driver, "//*[@id='feedlyFrame']")

        # Check if we reached the bottom
        new_height = driver.execute_script(
            "return document.querySelector('#feedlyFrame').scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height
        time.sleep(1.5)

    try:
        articles = driver.find_elements(By.CLASS_NAME, "entry.magazine")
        time.sleep(3)
        for article in articles:
            try:
                date_span = article.find_element(By.CSS_SELECTOR, "span[title*='Published']")
                article_date = date_span.get_attribute("title")
                new_article_date = re.sub(r'\n', '', str(article_date)).strip()
                new_article_date = re.sub(r'.*Received: | GMT.*', '', new_article_date).strip()
                if new_article_date[17:19] == '24':
                    new_article_date = new_article_date[:17] + '00' + new_article_date[19:]
                new_article_date_conv = datetime.datetime.strptime(new_article_date, "%a, %d %b %Y %H:%M:%S")
                if start_range <= new_article_date_conv <= new_today_str:
                    title = article.find_element(By.CSS_SELECTOR, "a.EntryTitleLink").text
                    link = article.find_element(By.CSS_SELECTOR, "a.EntryTitleLink").get_attribute("href")
                    
                    if len(new_articles) % batch_size == 0:
                        driver.execute_script("window.localStorage.clear();")
                        driver.execute_script("window.sessionStorage.clear();")

                    article = (title, link)
                    if article not in new_articles:
                        new_articles.append(article)
                        
            except Exception as e:
                print('Error for article: ' + str(e))
        time.sleep(1.5)
        driver.get(
            'https://feedly.com/i/collection/content/user/9e62dc2d-90e6-453b-88f4-47b630b9a4aa/category/global.all'
        )
        time.sleep(10)
        counter = 0
        new_last_height = driver.execute_script(
            "return document.querySelector('#feedlyFrame').scrollHeight")
        while counter < 10:
            # Scroll down and wait for page load
            scroll_down(driver, "//*[@id='feedlyFrame']")
            # Check if we reached the bottom
            new_height1 = driver.execute_script( 
                "return document.querySelector('#feedlyFrame').scrollHeight")
            if new_height1 == new_last_height:
                break
            new_last_height = new_height1
            counter += 1
            time.sleep(1.5)

        new_articles1 = driver.find_elements(By.CLASS_NAME, "entry.magazine")
        for article in new_articles1:
            try:
                date_span1 = article.find_element(By.CSS_SELECTOR, "span[title*='Published']")
                article_date1 = date_span1.get_attribute("title")
                new_article_date1 = re.sub(r'\n', '', str(article_date1)).strip()
                new_article_date1 = re.sub(r'.*Received: | GMT.*', '', new_article_date1).strip()
                if new_article_date1[17:19] == '24':
                    new_article_date1 = new_article_date1[:17] + '00' + new_article_date1[19:]
                new_article_date1_conv = datetime.datetime.strptime(new_article_date1, "%a, %d %b %Y %H:%M:%S")
                if start_range <= new_article_date1_conv <= new_today_str:
                    title = article.find_element(By.CSS_SELECTOR, "a.EntryTitleLink").text
                    link = article.find_element(By.CSS_SELECTOR, "a.EntryTitleLink").get_attribute("href")
                    
                    article = (title, link)
                    if article not in new_articles:
                        new_articles.append(article)
            except Exception as e:
                print('Error for article: ' + str(e))
    except Exception as e:
        print('Error execution: ' + str(e))
        pass
    finally:
        driver.quit()

    return new_articles


def main(email, password):
    df_no_duplicates = pd.DataFrame()

    # Define headers and empty data
    if not os.path.exists(r"Rory Testing Sheet 2024.xlsx"):
        headers = ["Url","Title","Description","Reach Out","Reasons", "Keywords","Location"]
        data = []
        df = pd.DataFrame(data, columns=headers)
        # Save to an Excel file
        df.to_excel(r"Rory Testing Sheet 2024.xlsx", index=False, engine="openpyxl")
        
        # adjust_column_width(r"Rory Testing Sheet 2024.xlsx")
    
    titles_to_check = []
    existing_titles = []
    titles = []
    pg_links = []
    pg_links_check = []
    titles_check = []
    titles_neg = []
    try:
        feedly_login(driver, email, password)
        time.sleep(1.4) 
        articles = scrape_today_articles(driver)

        # Process articles
        if articles:
            print('Total collected articles: ' + str(len(articles)))
            unique_new_articles_neg = [article for article in articles if article[0] not in titles_read]
            print('Total unique articles before negatives check: ' + str(len(unique_new_articles_neg)))
            unique_new_articles = [article for article in unique_new_articles_neg if article[0] not in negative_titles_read]
            print('Total unique articles after negatives check: ' + str(len(unique_new_articles)))
            total_titles = len(titles_read)
            print('Total titles read: ' + str(total_titles))
            for article in unique_new_articles:
                try:
                    if article[0] in existing_titles:
                        pass
                    else:
                        existing_titles.append(str(article[0]))
                        if 'news.google' in str(article[1]):
                            try:
                                decoded_url = decode_google_news_url(str(article[1]), interval=8)
                                if decoded_url.get("status"):
                                    if url_contains_keyword(decoded_url["decoded_url"], negative_keywords) or check_title_against_keywords(article[0], negative_keywords):
                                        titles_neg.append(str(article[0]))
                                        pass
                                    else:
                                        pg_links.append(decoded_url["decoded_url"])
                                        titles.append(str(article[0]))
                                else:
                                    pg_links.append(str(article[1]))
                                    titles.append(str(article[0]))
                            except Exception as e:
                                print('Error trying to convert google news link: ' + str(e))
                                pg_links.append(str(article[1]))
                        else:
                            pg_links.append(str(article[1]))
                            titles.append(str(article[0]))
                except Exception as ex:
                    print('Error converting urls: ' + str(ex))

            # # Save results with column width adjustment
            # if not df_no_duplicates.empty:
            #     append_to_excel(r"Rory Testing Sheet 2024.xlsx", df_no_duplicates, 'Sheet1')
            #     adjust_column_width(r"Rory Testing Sheet 2024.xlsx")
                
            #     if not df_titles.empty:
            #         append_to_excel(r'titles_to_check.xlsx', df_titles, 'Sheet1')
            #         adjust_column_width(r'titles_to_check.xlsx')
                
            #     if not df_titles_neg.empty:
            #         append_to_excel(r'negative_titles.xlsx', df_titles_neg, 'Sheet1')
            #         adjust_column_width(r'negative_titles.xlsx')

       
        # Use your OpenAI API key
        openai_api_key = OPEN_API_KEY

        if not openai_api_key:
            raise ValueError("OPENAI_API_KEY not found in .env file")
        
        print('Total pages to process: ' + str(len(pg_links)))
        print('Total titles to process: ' + str(len(titles)))
        if len(pg_links) > 0:
            print('Processing with ChatGPT Response...')
            # Process elements in chunks of 10
            chunk_size = 20
            options = Options()
            options.add_argument("--start-maximized")
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_argument("--ignore-certificate-errors")

            brw = uc.Chrome(options=options, driver_executable_path=os.getenv("CHROMEDRIVER_PATH"))
            brw.set_window_size(1370, 780)
            brw.get("https://chatgpt.com/")
            time.sleep(10) 
            login_to_chatgpt_com(brw)
            time.sleep(5)  # Allow page to load
        
            gpt_links = []
            gpt_titles = []
            gpt_descs = []
            gpt_reach = []
            gpt_reasons = []
            gpt_keywords = []
            gpt_loc = []
            while pg_links:
                chunk_lnks = pg_links[:chunk_size]
                chunk_ttls = titles[:chunk_size]
                prompt_text = ''
                for i in range(len(chunk_lnks)):
                    prompt_text += str(chunk_lnks[i]) + '\t' + str(chunk_ttls[i]) + '\n'
                new_prompt_text = prompt_text.strip()

                try:
                    modal_cls = brw.find_element(By.XPATH, "//*[contains(@id,'radix-:r')]/div/div/a")
                    modal_cls.close()
                    time.sleep(2)
                except Exception as ex:
                    pass

                text = f"""
                    Do that for the following:
        
                    {new_prompt_text}
                """

                # Step 2: Send the text to ChatGPT
                prompt = f"""PROMPT. BEFORE ANYTHING ELSE, Please Read entirely before performing any response! You must follow this prompt exactly or it will not work. Strictly follow the format with no deviations.
    
            Task: 
            Extract the following information from each article and format it with # separated values so when I get response with my Python script, in each line, I can separate each value by arr.split('#') and process each column value into corresponding lists. Please follow the formatting prompt outlined later in this prompt.
            
            Columns Required:
            Column A: The URL of the article. Examples of the URL which I dont want are:
            4.        https://www.thailand-business-news.com/thailand-prnews/jolicares-journey-from-steroid-scandal-to-leading-herbal-skincare-brand
             `https://www.kltv.com/2024/10/30/smith-county-elected-officials-trade-barbs-over-allegations-falsifying-timesheets/
             1`https://www.tillamookcountypioneer.net/tillamook-school-district-update-investigation-into-allegations-from-friday-football-game/
             The URL needs to be clean, without any leading number, dots, or quotes.
             Example of the URL I want:
             https://cityhub.com.au/gagged-with-a-sex-toy-shocking-details-of-st-pauls-hazing-scandal-emerge/
            Column B: The title of the article. You may use the title I have provided if available.
            Column C: A meta description of the article, summarizing it in about 100 - 200 words, with a minimum of 50 words for straightforward articles.
            Column D: Identify the individual or company (if no specific person has been named) from the article that made a mistake, wrong decision, poor decision, or is the defendant in the situation, even if no formal charges or wrongdoing have been confirmed. Only put the specific name. Only if the article is entirely positive with no controversy, write 'N/A Does Not Apply.'
            Column E: Start with the specific name of the wrongdoer or entity at fault, followed by an explanation of their role and why they are relevant, even if no formal charges or wrongdoing have been confirmed. Only if the article is entirely positive with no controversy, write 'N/A Does Not Apply.'
            Column F: Extract only the 20 most important and relevant keywords from the article. Focus on key entities and significant actions, avoiding common words and publication-related information. List exactly 20 keywords, separated by a vertical bar (|), with no spaces, tabs or periods between bars. Example: apple|banana|dragon fruit|kiwi|daves fruit.
            Column G: Identify where the article takes place. Include the town/city, county/state, and country if mentioned. If any of the information is missing, exclude those specific fields but combine all available information into a single cell. No vertical bar (|) or tabs between word/words.
            
            Automated Checklist Instructions:
            Before submitting any result, check for the presence of data in each column (A to G), especially Column F (keywords). If any column is missing, put 'N/A' as value so it keeps the correct columns order.
            Column A (URL): Ensure the URL is valid and filled.
            Column B (Title): Ensure the article title is filled from the article or provided.
            Column C (Meta Description): Ensure the meta description is between 50-200 words.
            Column D (Wrongdoer): Ensure the individual or company name is provided (or 'N/A Does Not Apply' if no wrongdoer is involved).
            Column E (Explanation): Ensure the explanation of the wrongdoer is concise and relevant.
            Column F (Keywords): Ensure exactly 20 keywords are extracted and properly formatted using the following rules: You cannot skip this step. 
            No spaces or periods between bars.
            No apostrophes or commas.
            No tabs.
            Use vertical bars to separate keywords (|).
            Column G (Location): Ensure the location is combined into a single cell with available details without vertical bar(|) or tabs.
            Important: Make sure that all columns have values. If some column is missing a value or value is an empty string, replace it with 'N/A' 
            Handling Duplicates, Technical Issues, and Errors:
            If a URL cannot be accessed due to technical issues, note "Unable to access URL" and proceed to the next URL.
            If the URL has a formatting issue (such as special characters), note "Formatting issue" and proceed to the next URL.
            If a duplicate URL is encountered, process it fully as if it were new. You may not skip any URL.
            Important: If any column is missing or incorrectly filled, it must be flagged and reviewed before moving forward.
            Final Output:
            Once all fields are validated against the automated checklist, provide the batch output formatted with # separated values, so when I get response with my Python script, in each line, I can separate each value by arr.split('#') and process each column value into corresponding lists.
            
            Batching and Confirmation:
            Break the URLs into manageable batches of 10 and process them in order.
            You will ensure you have completed each column before submitting results.
            Continue processing until all URLs are completed in the same manner.
            Double-check your work for 100% accuracy before giving the response for each batch.  You will give the data in the correct # separated format with no lines after each row.
            Strictly follow the format with no deviations. When all URLs are completed, notify me that you are done.
            IMPORTANT: ALWAYS provide me the batch output in the response each value separated with # sign and make sure to provide batch output inside your code or batch response.
            Make sure that all columns for each line have correct and corresponding value. If any of values is missing, or you can't find it, replace it with N/A.
            Make sure that there's 6 values total (one for each column) and 5 hashtags (#) which is separation sign for each column, which is format for getting response for my Python script so I can separate each value by arr.split('#') and process each column value into corresponding lists.
            Please make sure that you process all URLs from the prompt in your response and not just giving me one and asking to confirm. Just do as I asked in the prompt.
    
            {text}
                """
                time.sleep(8)
                try:
                    gpt_response = ask_chatgpt_com(brw, prompt)#call_chatgpt_api(prompt, openai_api_key)

                    if gpt_response is not None:

                        gpt_json_loaded = json.loads(gpt_response)
                        
                        for i in range(len(gpt_json_loaded)):
                            gpt_links.append(chunk_lnks[i])
                            gpt_titles.append(chunk_ttls[i])
                            gpt_descs.append(gpt_json_loaded[i]['description'])
                            gpt_reach.append(gpt_json_loaded[i]['reach_out'])
                            gpt_reasons.append(gpt_json_loaded[i]['reasons'])
                            gpt_keywords.append(gpt_json_loaded[i]['keywords'])
                            gpt_loc.append(gpt_json_loaded[i]['location'])
                    else:
                        pass
                except Exception as e:
                    print('Not collected data for this batch! Error: ' + str(e))
                    pass
                try:
                    modal_cls = brw.find_element(By.XPATH, "//*[contains(@id,'radix-:r')]/div/div/a")
                    modal_cls.close()
                    time.sleep(2)
                except Exception as ex:
                    pass
##                    for j in range(len(chunk_lnks)):
##                        gpt_links.append(chunk_lnks[j])
##                        gpt_titles.append(chunk_ttls[j])
##                        gpt_descs.append('-')
##                        gpt_reach.append('-')
##                        gpt_reasons.append('-')
##                        gpt_keywords.append('-')
##                        gpt_loc.append('-')

                # Remove the processed chunk from the list
                pg_links = pg_links[chunk_size:]
                titles = titles[chunk_size:]
            if len(gpt_descs) > 0:
                df = pd.DataFrame({
                    "Url": gpt_links,
                    "Title": gpt_titles,
                    "Description": gpt_descs,
                    "Reach Out": gpt_reach,
                    "Reasons": gpt_reasons,
                    "Keywords": gpt_keywords,
                    "Location": gpt_loc
                })

                df_no_duplicates = df.drop_duplicates(subset=['Title'])
            else:
                print("No new links in articles!")
                df_no_duplicates = pd.DataFrame()

            try:
                df_titles = pd.DataFrame({
                    "Titles": gpt_titles
                })
            except Exception as ex:
                df_titles = pd.DataFrame()
            try:
                append_to_excel(r'titles_to_check.xlsx', df_titles, 'Sheet1')
            except Exception as ex:
                print('Error appending data: ' + str(ex))
            try:
                df_titles_neg = pd.DataFrame({
                    "Titles": titles_neg
                })
            except Exception as ex:
                df_titles_neg = pd.DataFrame()
            try:
                append_to_excel(r'negative_titles.xlsx', df_titles_neg, 'Sheet1')
            except Exception as ex:
                print('Error appending data: ' + str(ex))

            if not df_no_duplicates.empty:
                try:
                    append_to_excel(r"Rory Testing Sheet 2024.xlsx",df_no_duplicates,'Sheet1')

                except Exception as e:
                    print('Error adding to Google Spreadsheet: ' + str(e))
        else:
            print("No new links in articles!")
            df_no_duplicates = pd.DataFrame()

    finally:
        try:
            brw.quit()
        except Exception as e:
            pass
        print('Completed!')


if __name__ == "__main__":
    start = time.time()
    email = "m08067064@gmail.com"
    password = "thebestrep2025"
    main(email, password)
    end = time.time()
    print("Time Taken: {:.6f}s".format(end - start))
