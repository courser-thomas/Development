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
# Locally play a DRM movie
#
# Setup instructions:
#   Purchase "Halo 2" from the store, and download.
##

import logging
import os
import scenarios.app_scenario
from parameters import Params
from appium import webdriver
import time
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
import selenium.common.exceptions as exceptions
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By


class LvpDrm(scenarios.app_scenario.Scenario):

    module = __module__.split('.')[-1]
    # Set default parameters
    Params.setDefault(module, 'title', 'Halo 2')  # Just has to be unique substring of title
    Params.setDefault(module, 'duration', '1200')  # 1200 Seconds
    Params.setDefault(module, 'trace_provider', 'multimedia.wprp')
    Params.setDefault(module, 'trace_file', module + ".etl")
    
    # Get parameters
    title = Params.get(module, 'title')
    duration = Params.get(module, 'duration')
    platform = Params.get('global', 'platform')
    training_mode = Params.get('global', 'training_mode')

    # Local parameters
    prep_list = ["halo2_prep"]

    # If training mode, just 1 fast loop
    if training_mode == "1":
        duration = '10'

    def prepCheck(self):
        self.assert_list = ""

        if self.assert_list != "":
            self._assert(self.assert_list)

    def setUp(self):
        prep_status = self.checkPrepStatus(self.prep_list)
        if prep_status !="":
             self._assert(prep_status)

        # Check if video is purchased
        title = 'Remaking the Legend'

        logging.info("Performing setup: Launching WinAppDriver.exe on DUT.")
        self._call([(self.dut_exec_path + "\\WindowsApplicationDriver\\WinAppDriver.cmd"), (self.dut_ip + " " + self.app_port)], blocking=False)
        
        logging.info("Launching Movies and TV")
        desired_caps = {}
        desired_caps["app"] = "Microsoft.ZuneVideo_8wekyb3d8bbwe!microsoft.ZuneVideo"
        driver = self._launchApp(desired_caps)
        driver.maximize_window()
        logging.info("Looking for  " + self.title)
              
        # Search for movie title
        time.sleep(1)
        driver.find_element_by_name("Purchased").click()
        time.sleep(1)
        driver.find_element_by_name("Search").click()
        time.sleep(1)
        driver.find_element_by_name("Search").send_keys(title + Keys.ENTER)
        time.sleep(10)
        
        try:
            logging.info("Checking if the movie is already purchased")
            driver.find_element_by_xpath("//*[contains(@Name, 'In your collection (1)')]")
        except:
            self.assert_list += "Video is not available in your collection"

        # Click on Movie Title
        driver.find_element_by_class_name("GridViewItem").find_element_by_xpath("//*[contains(@Name, '" + self.title + "')]").click()
        time.sleep(10)

        #driver.find_element_by_name("Download Remaking the Legend - Halo2: Anniversary").click()
        driver.find_element_by_xpath("//*[contains(@Name, 'Download')]").click()
        time.sleep(3)

        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'Download'))).click()
        
            # Need to wait for download here
            #WebDriverWait(driver, 400).until(EC.presence_of_element_located((By.NAME, 'Remove Remaking the Legend - Halo 2: Anniversary download')))
            WebDriverWait(driver, 400).until(EC.presence_of_element_located((By.xpath, "//*[contains(@Name, 'Remove')]")))
            logging.info("Halo 2 downloaded")
        except:
            logging.info("Halo 2 is already downloaded")
            pass

        driver.close()
        # Killing WinApp driver
        self._kill("WinAppDriver.exe")
        scenarios.app_scenario.Scenario.setUp(self)
        time.sleep(2)


    def runTest(self):

        logging.info("Performing setup: Launching WinAppDriver.exe on DUT.")
        self._call([(self.dut_exec_path + "\\WindowsApplicationDriver\\WinAppDriver.cmd"), (self.dut_ip + " " + self.app_port)], blocking=False)
        
        logging.info("Launching Movies and TV")
        desired_caps = {}
        desired_caps["app"] = "Microsoft.ZuneVideo_8wekyb3d8bbwe!microsoft.ZuneVideo"
        self.driver = self._launchApp(desired_caps)
        self.driver.maximize_window()
        self.playMovie()
        logging.info("Duration: " + self.duration + " seconds")
        # Let play for specified duration
        time.sleep(float(self.duration))

    def playMovie(self):
        logging.info("Playing " + self.title)

        # Navigate to Purchased menu
        self.driver.find_element_by_name("Purchased").click()

        # Find specified movie title
        self.driver.find_element_by_xpath('//*[contains(@Name,"' + self.title + '")]').click()
        time.sleep(5)

        # Common case is to Restart the movie, so search for that button first to minimize operations.
        # If not found, then try to find the Play button, as is the case for the first time playing.
        try:
            self.driver.find_element_by_xpath('//*[contains(@Name,"' + "Restart" + '")]').click()
        except exceptions.NoSuchElementException:
            try:
                self.driver.find_element_by_xpath('//*[contains(@Name,"' + "Play" + '")]').click()
            except:
                self.fail("Could not find button to start movie")

        time.sleep(10)
        try:
            logging.info("Selecting Repeat button on DUT.")
            self.driver.find_element_by_name("More").click()
            time.sleep(2)
            self.driver.find_element_by_name("Turn repeat on").click()
            time.sleep(1)
            # Move cursor into window so menu disappears.
            ActionChains(self.driver).move_by_offset(-500, -500).perform()
        except NoSuchElementException as EX:
            logging.info("Repeat button not found.")
            pass
        time.sleep(2)
        try:
            self.driver.find_element_by_name("Full Screen").click()
        except:
            logging.info("Did not find the button for full screen.  Must already be full screen.")
        time.sleep(2)
        return


    def tearDown(self):
        logging.info("Performing teardown.")
        try:
            ActionChains(self.driver).move_by_offset(-100, 100).perform()
            time.sleep(2)
            self.driver.find_element_by_name("Exit Full Screen").click()
            time.sleep(2)
        except NoSuchElementException as EX:
            logging.info("Exit Full Screen button not found.")
            time.sleep(2)
            pass
        try:
            self.driver.find_element_by_name("Close Movies & TV").click()
            time.sleep(2)
        except NoSuchElementException as EX:
            logging.info("Close button not found.")
            time.sleep(2)
            pass
        scenarios.app_scenario.Scenario.tearDown(self)
        self._kill("WinAppDriver.exe")


    def kill(self):
        try:
            logging.debug("Killing Video.UI.exe")
            self._kill("Video.UI.exe")
        except:
            pass

        
