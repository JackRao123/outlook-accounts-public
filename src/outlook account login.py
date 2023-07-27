from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import selenium.webdriver.support.expected_conditions as EC
from selenium.webdriver.common.by import By
import selenium.webdriver.support.ui as ui
from selenium.common.exceptions import TimeoutException
import random
from time import sleep as wait
import sys
from pyvirtualdisplay import Display
import json
import zipfile
import threading
import time
import warnings
import os
warnings.simplefilter("ignore", DeprecationWarning)
CREATE_ACC_CONFIG_FILE = r'..\CONFIG\create_acc_config.json'


class Login:
    def __init__(self):
        try:
            with open(CREATE_ACC_CONFIG_FILE) as json_data_file:
                config = json.load(json_data_file)
                self.proxy_file = config['create_proxy_file']
 
               
                #self.api_key = config['captcha_api_key']
                self.headless = config['headless']
        except FileNotFoundError:
            print("No config.json found")
        self.outlook_url = "https://outlook.live.com/owa/?nlp=1"
        self.chrome_options = Options()
        self.exe_path = r"../DRIVERS/chromedriver114.exe"
            

        self.chrome_options.add_argument('--disable-web-security')
        self.chrome_options.add_argument('--disable-site-isolation-trials')
        self.chrome_options.add_argument('--disable-application-cache')

        self.chrome_options.add_argument("--window-size=1920x1080")
        if self.headless == "True":
            self.chrome_options.add_argument('--headless')

        proxy_string = self.get_proxy_from_file()
        if(proxy_string!="NULL"):
            self.chrome_options = self.attach_proxy_to_options(self.chrome_options, proxy_string )
        self.password = ""
        self.waiting_for_user_to_input_email = False


    def is_visible(self, locator, timeout=30):
        try:
            ui.WebDriverWait(self.driver, timeout).until(EC.visibility_of_element_located((By.ID, locator)))
            return True
        except TimeoutException:
            return False

    def attach_proxy_to_options(self, chrome_options, proxy_string):  # proxy strign is user:pass@host:port
        # takes a chrome_options (webdriver.ChromeOptions()) and adds a proxy to it.

        user_password, host_port = proxy_string.split("@")
        user, password = user_password.split(":")
        host, port = host_port.split(":")

         

        PROXY_HOST = host  # rotating proxy or host
        PROXY_PORT = port  # port
        PROXY_USER = user  # username
        PROXY_PASS = password  # password

        manifest_json = """
        {
            "version": "1.0.0",
            "manifest_version": 2,
            "name": "Chrome Proxy",
            "permissions": [
                "proxy",
                "tabs",
                "unlimitedStorage",
                "storage",
                "<all_urls>",
                "webRequest",
                "webRequestBlocking"
            ],
            "background": {
                "scripts": ["background.js"]
            },
            "minimum_chrome_version":"22.0.0"
        }
        """

        background_js = """
        var config = {
                mode: "fixed_servers",
                rules: {
                singleProxy: {
                    scheme: "http",
                    host: "%s",
                    port: parseInt(%s)
                },
                bypassList: ["localhost"]
                }
            };

        chrome.proxy.settings.set({value: config, scope: "regular"}, function() {});

        function callbackFn(details) {
            return {
                authCredentials: {
                    username: "%s",
                    password: "%s"
                }
            };
        }

        chrome.webRequest.onAuthRequired.addListener(
                    callbackFn,
                    {urls: ["<all_urls>"]},
                    ['blocking']
        );
        """ % (PROXY_HOST, PROXY_PORT, PROXY_USER, PROXY_PASS)

        pluginfile = rf'../ProxyPlugins/proxy_auth_plugin{random.randint(0,10000)}.crx'

        with zipfile.ZipFile(pluginfile, 'w') as zp:
            zp.writestr("manifest.json", manifest_json)
            zp.writestr("background.js", background_js)

        chrome_options.add_extension(pluginfile)



        return chrome_options

    def get_proxy_from_file(self):
        with open(r'' + self.proxy_file, "r") as json_file:
            json_data = json.load(json_file)

        proxy_list = []
        proxies = json_data['proxies']
        for key, value in proxies.items():
            if value['active'] == True:
                proxy_list.append(value['proxy'])

        if(len(proxy_list)==0):
            return "NULL"
        return random.choice(proxy_list)
    def page_has_loaded(self):
        page_state = self.driver.execute_script('return document.readyState;')
        return True

    def log_in(self, email, password):
        try:

            proxy_string = self.get_proxy_from_file()

            chrome_opts = Options()
            if(proxy_string!="NULL"):
                chrome_opts  = self.attach_proxy_to_options( Options(), proxy_string    )
            self.driver = webdriver.Chrome(options = chrome_opts, executable_path=self.exe_path)

            self.driver.get(self.outlook_url)



            if self.page_has_loaded() is True:
                if self.is_visible("i0116") is True: #the 'email = ' box
                    pass



            self.driver.find_element(By.ID, "i0116").send_keys(email)
            self.driver.find_element(By.ID,"idSIButton9").click()
            time.sleep(1) #gotta wait a little bit so that when we click the 'sign in' button we dont get a error saying that the element isnt there (we gotta wait for it to load)
            self.driver.find_element(By.ID,"i0118").send_keys(password)
            self.driver.find_element(By.XPATH, '//*[@id="idSIButton9"]').click()


            if self.is_visible("passwordError",2):
                self.driver.find_element(By.ID, "i0118").send_keys("Testing123$") # old password used for testing with outlook account creator bot.
                self.driver.find_element(By.XPATH, '//*[@id="idSIButton9"]').click()

            start_timer = time.time()
            locator_id_array = ["iShowSkip", "idSIButton9", "iCancel", "id__0"]
            while(time.time() - start_timer < 30):
                for id in locator_id_array:
                    try:
                        self.driver.find_element(By.ID, id).click() 
                    except:
                        pass
                    time.sleep(0.5)

            
            # if self.is_visible("iShowSkip", 3.5):
            #     self.driver.find_element(By.ID, "iShowSkip").click() #'stay signed in?'

            # if self.is_visible("idSIButton9", 3.5):
            #     self.driver.find_element(By.ID, "idSIButton9").click() #'stay signed in?'
            
            # if self.is_visible("iCancel", 3.5):
            #     self.driver.find_element(By.ID, "iCancel").click() #'stay signed in?'

        

            time.sleep(100000)#so that the driver doesn't close when its done loggin in



        except Exception as e:
            print('Error on line {}'.format(sys.exc_info()[-1].tb_lineno), type(e).__name__, e)

 

class DriverManager:
    def __init__(self):
        self.drivers = []
        self.waiting_for_user_to_input_email = False



    def add_driver(self, driver):
        self.drivers.append(driver)

    def input_thread(self):
        if self.waiting_for_user_to_input_email == False:
            logInBot = Login()
            self.waiting_for_user_to_input_email = True
            email = input("Enter emails to log in to (comma or space separated): ")
            self.waiting_for_user_to_input_email= False
            emails = email.split(",")
            emails = email.split(" ")

            print(f"emails = {emails}")

            for email in emails:

                thr = threading.Thread(target = driverManager.start_thread, args = (email,))  
                thr.start()
                time.sleep(0.1)



    def start_thread(self, email):
        logInBot = Login()
        current_thread = threading.Thread(target=logInBot.log_in, args=(email, self.password))
        current_thread.start()

if __name__ == "__main__":
     
 
        

    
    driverManager = DriverManager()

    threads = []

    print("Outlook Account Login Tool")
    password = input("Enter master password: ")
    os.system('cls')
    driverManager.password = password

    waiting_for_user_to_input_email = False
    while True:

        time.sleep(0.1)  # we have to put this here because otherwise the computer is too fast?? or something like that idk and it starts new threads
        if driverManager.waiting_for_user_to_input_email == False:
            input_thread = threading.Thread(target=driverManager.input_thread)
            input_thread.start()






