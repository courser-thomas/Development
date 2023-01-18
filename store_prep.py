"""
//--------------------------------------------------------------
//
// HOBL
// Copyright(c) Microsoft Corporation
// All rights reserved.
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files(the ""Software""),
// to deal in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and / or sell copies
// of the Software, and to permit persons to whom the Software is furnished to do so,
// subject to the following conditions :
//
// The above copyright notice and this permission notice shall be included
// in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
// FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.IN NO EVENT SHALL THE AUTHORS
// OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
// WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF
// OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
//--------------------------------------------------------------
"""

##
# Update store apps and then disable store updates on RS5 and 19h1 only
# 
##

import os
import logging
import time
from parameters import Params
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchWindowException
from selenium.common.exceptions import ElementNotVisibleException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
import scenarios.app_scenario


class StorePrep(scenarios.app_scenario.Scenario):
    module = __module__.split('.')[-1]

    # Override collection of config data, traces, and execution of callbacks 
    is_prep = True
    new_store = False


    def get_driver(self):
        desired_caps = {}
        desired_caps["app"] = "Microsoft.WindowsStore_8wekyb3d8bbwe!App"
        driver = self._launchApp(desired_caps)
        driver.maximize_window()
        time.sleep(5)
        return driver

    def navigate_to_store_downloads(self):
        logging.info("Launching WinAppDriver.exe on DUT.")
        self._call([(self.dut_exec_path + "\\WindowsApplicationDriver\\WinAppDriver.exe"), (self.dut_ip + " " + self.app_port)], blocking=False)
        driver = self.get_driver()

        # Navigate to store updates
        try:
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.NAME, 'See More'))).click()
            WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, '//*[contains(@Name, "Downloads and updates")]'))).click()
        except:
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.NAME, 'Library'))).click()
            self.new_store = True

        return driver


    def runTest(self):
        logging.info("Uninstalling Microsoft Whiteboard app if installed")
        self._call(["powershell.exe", "Get-AppxPackage *Microsoft.Whiteboard* | Remove-AppxPackage"], expected_exit_code="")

        start_time = time.time()
        logging.info("Launching WinAppDriver.exe on DUT.")
        self.driver = self.navigate_to_store_downloads()

        last_round = False
        timeout = 600
        
        # Waiting untill all the updates are installed
        # We click "Get updates" at each loop and wait for timeout time
        # We click "Resume all" for any app in paused state
        while True:
            try:
                WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.NAME, 'Get updates'))).click()

                if self.new_store:
                    time.sleep(2)
                    WebDriverWait(self.driver, timeout, poll_frequency=3.0, ignored_exceptions=(ElementNotVisibleException)).until(EC.element_to_be_clickable((By.NAME, 'Get updates')))
                    try:
                        WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.NAME, "Update all")))
                        self.driver.find_element_by_name('Update all').click()
                    except:
                        break
                    WebDriverWait(self.driver, timeout, poll_frequency=3.0, ignored_exceptions=(NoSuchElementException)).until_not(EC.presence_of_element_located((By.NAME, 'Update all')))
                else:
                    try:
                        WebDriverWait(self.driver, 25).until(EC.element_to_be_clickable((By.XPATH, '//*[@Name="Update all" or @Name="Pause all" or @Name="Resume all"]')))
                        self.driver.find_element_by_name('Update all').click()
                    except TimeoutException:
                        if last_round:
                            WebDriverWait(self.driver, 2).until(EC.presence_of_element_located((By.NAME, 'All your trusted apps and games from Microsoft Store have the latest updates.')))
                            logging.info("All apps are installed")
                            break
                        pass
                    except NoSuchElementException:
                        logging.info('Update all not found')
                        pass

                    WebDriverWait(self.driver, timeout, poll_frequency=3.0, ignored_exceptions=(ElementNotVisibleException)).until_not(EC.element_to_be_clickable((By.NAME, 'Pause all')))

                    try:
                        self.driver.find_element_by_xpath('Resume all').click()
                    except :
                        try:
                            self.driver.find_element_by_name('Update').click()
                        except :
                            last_round = True
                            pass
                        pass
                timeout = timeout // 2

            except NoSuchWindowException:
                logging.info("New window opened")
                time.sleep(5)
                self.driver = self.navigate_to_store_downloads()
                pass
            except TimeoutException:
                logging.info("Time out exception")
                if last_round:
                    break
                # last_round = False
                pass
            
            
        # Need to add code to log if errors are observed during the store update
        
        # Navigate to store settings
        if self.new_store:
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.NAME, 'User profile'))).click()
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.NAME, 'App settings'))).click()
        else:
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.NAME, 'See More'))).click()
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.NAME, 'Settings'))).click()
        time.sleep(2)

        # Disable updates
        if self.new_store:
            ele = self.driver.find_element_by_class_name("ToggleSwitch").find_element_by_xpath('//*[contains(@Name, "Update apps automatically")]')
            if ele.is_selected():
                ele.click()
                time.sleep(1)
        else:
            try:
                self.driver.find_element_by_name("Update apps automatically On").click()
                time.sleep(1)
            except:
                logging.info("Update apps automatically is off")
                pass

            try:
                self.driver.find_element_by_name("Update apps automatically when I'm on Wi-Fi On").click()
                time.sleep(1)
            except:
                logging.info("Update apps automatically when I'm on Wi-Fi is off")
                pass

        # Close store driver
        self.driver.close()
        logging.info("Installing apps took: {0:.1f}s".format(time.time() - start_time))

    def tearDown(self):    
        logging.info("Performing teardown.")
        scenarios.app_scenario.Scenario.tearDown(self)
        time.sleep(2)
        self._kill("WinAppDriver.exe")
