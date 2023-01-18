
'''
//--------------------------------------------------------------
//
// HOBL
// Copyright(c) Microsoft Corporation
// All rights reserved.
//
// MIT License
//
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files(the 'Software'),
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
'''
from __future__ import unicode_literals
from __future__ import print_function
from __future__ import division
from __future__ import absolute_import

##
#
# Copy OneDrive resources to DUT and configure settings 
#
##

from future import standard_library
# standard_library.install_aliases()
import builtins
import logging
import time
import os
import scenarios.app_scenario
from selenium.webdriver.common.keys import Keys
from parameters import Params
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

import sys, errno
import shutil
import remote_file_ops
import win32wnet


class oneDrivePrep(scenarios.app_scenario.Scenario):
    module = __module__.split('.')[-1]

    #Get parameters
    onedrive_username = Params.get('global', 'msa_account')
    onedrive_password = Params.get('global', 'dut_password')

    is_prep = True

    def setUp(self):
        logging.info("Launching WinAppDriver.exe on DUT.")
        self._call([(self.dut_exec_path + "\\WindowsApplicationDriver\\WinAppDriver.cmd"), (self.dut_ip + " " + self.app_port)], blocking=False)
        time.sleep(1)
        desired_caps = {}
        desired_caps["app"] = "Root"
        self.driver = self._launchApp(desired_caps)
        scenarios.app_scenario.Scenario.setUp(self)
        
    def unCheck(self, checkbox_name):
        time.sleep(1)
        try:
            elem = self.driver.find_element_by_xpath('//*[contains(@Name,"' + checkbox_name + '")]')
            if elem.is_selected():
                elem.click()
                logging.info("'" + checkbox_name + "' has been unchecked")
            else:
                logging.info("'" + checkbox_name + "' is already unchecked")
        except:
            logging.info("'" + checkbox_name + "' not found, skipping")

    def runTest(self):
        if Params.get("global", "local_execution") == "0":
            # For remote execution, use the cmd.exe that matches the OS
            cmd = "cmd.exe"
        else:
            # For local execution, since this is 32b Python, we need to explicitly specify the native cmd.exe
            cmd = "c:\\windows\\sysnative\\cmd.exe"

        # What does this do?
        self._call([cmd, r'/C reg add HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Modules\GlobalSettings\Sizer /v PageSpaceControlSizer /t REG_BINARY /d a0000000010000000000000056050000 /f'])
        # Prevent "Deleted files are removed everywhere" reminder
        self._call([cmd, r'/C reg add HKLM\SOFTWARE\Policies\Microsoft\OneDrive /v DisableFirstDeleteDialog /t REG_DWORD /d 1 /f'])
        # Prevent mass delete popup
        self._call([cmd, r'/C reg add HKCU\Software\Microsoft\OneDrive\Accounts\Personal /v MassDeleteNotificationDisabled /t REG_DWORD /d 1 /f'])
        # Prevent first delete popup
        self._call([cmd, r'/C reg add HKCU\Software\Microsoft\OneDrive /v FirstDeleteDialogsShown /t REG_DWORD /d 1 /f'])
        time.sleep(2)

        # Inject ESCAPE just in case Start menu was left open
        ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
        time.sleep(1)

        # Click Start and escape to work around bug where we can't enter text into Start menu after daily_prep
        start_button = self._get_search_button(self.driver)
        start_button.click()
        time.sleep(3)
        ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
        time.sleep(2)

        # Click Start and escape to work around bug where we can't enter text into Start menu after daily_prep
        start_button = self._get_search_button(self.driver)
        start_button.click()
        time.sleep(3)
        ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
        time.sleep(2)

        logging.info("Starting OneDrive thru start menu")
        start_button = self._get_search_button(self.driver)
        start_button.click()
        time.sleep(5)    
        
        self.slow_send_keys("app:OneDrive")
        time.sleep(5)
        
        app_item = self.driver.find_element_by_name("Results").find_element_by_xpath('//*[contains(@Name,"OneDrive")]')
        ActionChains(self.driver).click(app_item).perform()
        time.sleep(2) 

        # self.driver.find_element_by_name("Maximize").click()
        
        logging.info("Checking if we are already signed in to Onedrive account")
        try:
            self.driver.find_element_by_xpath('//*[contains(@Name, "Set up OneDrive")]')
        except:
            logging.info("OneDrive account is already signed in")
            pass
        else:
            logging.info("We should sign in to Onedrive account")
            logging.info("Using Onedrive account of " + self.onedrive_username)
            email_field = self.driver.find_element_by_name("Enter your email address")
            email_field.click()
            time.sleep(1)
            email_field.send_keys(self.onedrive_username)
            # ActionChains(self.driver).send_keys(self.onedrive_username).perform()
            time.sleep(1)
            self.driver.find_element_by_name("Sign in").click()
            time.sleep(5)
            try:
                logging.info("Entering email password for OneDrive")
                self.driver.find_element_by_xpath('//*[contains(@Name, "Enter the password for")]').click()
                ActionChains(self.driver).send_keys(self.onedrive_password).perform()
                time.sleep(1)
                self.driver.find_element_by_name("Sign in").click()
                time.sleep(10)
            except:
                pass
            logging.info("Click Next")
            self.driver.find_element_by_name("Next").click()
            time.sleep(2)
            try:
                self.driver.find_element_by_name("Use this folder").click()
                time.sleep(3)
            except:
                pass
            # ActionChains(self.driver).send_keys(Keys.TAB).perform()
            # ActionChains(self.driver).send_keys(Keys.ENTER).perform()
            time.sleep(15)
            self.driver.find_element_by_name("Not now").click()
            time.sleep(3)
            self.driver.find_element_by_name("Next").click()
            time.sleep(3)
            self.driver.find_element_by_name("Next").click()
            time.sleep(3)
            self.driver.find_element_by_name("Next").click()
            time.sleep(3)
            self.driver.find_element_by_name("Later").click()
            time.sleep(3)
            # self.driver.find_element_by_name("Next").click()
            # time.sleep(3)
            logging.info("Opening my OneDrive folder")
            self.driver.find_element_by_name("Open my OneDrive folder").click()
            time.sleep(5)

        # Finish setting up OneDrive, if needed
        try:
            ok_Click = WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.NAME, "OK")))
            ok_Click.click()
            
            start_button = self._get_search_button(self.driver)
            start_button.click()
            time.sleep(5)

            self.slow_send_keys("app:OneDrive")
            time.sleep(2)

            app_item = self.driver.find_element_by_name("Results").find_element_by_xpath('//*[contains(@Name,"OneDrive")]')
            ActionChains(self.driver).click(app_item).perform()
            time.sleep(5)
        except:
            pass
        
        # Right click on OneDrive to open settings 
        try:
            time.sleep(2)
            try:
                self.driver.find_element_by_xpath('//Button[starts-with(@Name,"OneDrive")]')
            except:
                self.driver.find_element_by_name("Notification Chevron").click()
                time.sleep(8)
                pass
            ele = self.driver.find_element_by_xpath('//Button[starts-with(@Name,"OneDrive")]')
            ActionChains(self.driver).move_to_element(ele).perform()
            time.sleep(0.5)
            ActionChains(self.driver).context_click().perform()
            time.sleep(1)
            self.driver.find_element_by_xpath('//*[contains(@Name,"Settings menu")]').click()
            try:
                ActionChains(self.driver).context_click(self.driver.find_element_by_name("Control Host").find_element_by_name("OneDrive - Personal")).perform()
                time.sleep(1)
            except:
                pass
        except:
            logging.info("Starting OneDrive thru start menu")
            start_button = self._get_search_button(self.driver)
            start_button.click()
            time.sleep(5)    
            
            self.slow_send_keys("app:OneDrive")
            time.sleep(2)

            self.driver.find_element_by_name("Maximize").click()
            time.sleep(2)

            app_item = self.driver.find_element_by_name("Results").find_element_by_xpath('//*[contains(@Name,"OneDrive")]')
            ActionChains(self.driver).click(app_item).perform()
            time.sleep(2)

            ActionChains(self.driver).context_click(self.driver.find_element_by_name("Control Host").find_element_by_name("OneDrive")).perform()
            time.sleep(1)

        try:
            self.driver.find_element_by_name("Settings").click()
            time.sleep(1)
        except:
            logging.info("Sending ESCAPE sequence")
            ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
            time.sleep(1)
            logging.info("Trying to move cursor")
            ActionChains(self.driver).move_by_offset(0, 10).perform()
            time.sleep(1)
            logging.info("Trying to right click on OneDrive")
            ActionChains(self.driver).context_click().perform()
            time.sleep(1)
            logging.info("Trying to select Settings")
            self.driver.find_element_by_name("Settings").click()
            time.sleep(1)

        # Settings tab settings
        self.driver.find_element_by_name("Settings").click()
        time.sleep(1)
        self.unCheck("battery saver mode")
        time.sleep(1)
        self.unCheck("metered")
        time.sleep(1)
        self.unCheck("share with me or edit my shared")
        time.sleep(1)
        self.unCheck("files are deleted in the cloud")
        time.sleep(1)
        self.unCheck("from the cloud")
        time.sleep(1)
        self.unCheck("Notify me when sync is auto-paused")
        time.sleep(1)
        self.unCheck("sync pauses automatically")
        time.sleep(1)
        self.unCheck("memories")
        time.sleep(1)
        # self.unCheck("Save space and download files as you use them")
        # time.sleep(1)
        # try:
        #     self.driver.find_element_by_xpath('//*[contains(@AutomationId,"CommandButton_")]').click()
        #     time.sleep(1)
        # except:
        #     logging.info("No need to confirm 'Disable Files On-Demand'")

        # Office tab settings
        self.driver.find_element_by_name("Microsoft OneDrive").find_element_by_name("Office").click()
        time.sleep(1)
        radio_1 = "Let me choose to merge changes or keep both copies"
        radio_2 = "Always keep both copies (rename the copy on this computer)"
        if self.driver.find_element_by_name(radio_1).is_selected():
            self.driver.find_element_by_name(radio_2).click()
            logging.info("'" + radio_2 + "' has been unchecked")
            time.sleep(1)
        else:
            logging.info("'" + radio_2 + "' is already unchecked")
        # self.driver.find_element_by_xpath('//*[contains(@Name,"OK")]').click()
        self.driver.find_element_by_name("OK").click()
        time.sleep(1)

        logging.info("Closing OneDrive window")
        self.driver.find_element_by_xpath('//Window[contains(@Name,"OneDrive")]').find_element_by_xpath('//*[contains(@Name,"Close")]').click()

        self.createPrepStatusControlFile()  

    def tearDown(self):
        logging.info("Performing teardown.")
        scenarios.app_scenario.Scenario.tearDown(self)
        self._kill("WinAppDriver.exe")

    def slow_send_keys(self, keys):
        for key in keys:
            ActionChains(self.driver).send_keys(key).perform()
            time.sleep(0.2)

