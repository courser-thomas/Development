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
# Play a local video
#
# Setup instructions:
#   Copy Tears of Steel to default Video folder
##

from builtins import str
import builtins
import logging
import os
import scenarios.app_scenario
from parameters import Params
from appium import webdriver
import time
from selenium.webdriver.common.action_chains import ActionChains
import selenium.common.exceptions as exceptions
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException

class LVP(scenarios.app_scenario.Scenario):

    module = __module__.split('.')[-1]
    # Set default parameters
    # Params.setDefault(module, 'title', 'tears_of_steel')  # Just has to be unique substring of title
    Params.setDefault(module, 'title', 'ToS-4k-1920')  # Just has to be unique substring of title
    Params.setDefault(module, 'duration', '300')  # Seconds
    Params.setDefault(module, 'trace_provider', 'multimedia.wprp')
    Params.setDefault(module, 'trace_file', module + ".etl")
    Params.setDefault(module, 'airplane_mode', '1')
    Params.setDefault(module, 'radio_enable', '1')

    # Get parameters
    title = Params.get(module, 'title')
    duration = Params.get(module, 'duration')
    airplane_mode = Params.get(module, 'airplane_mode')
    platform = Params.get('global', 'platform')
    training_mode = Params.get('global', 'training_mode')
    radio_enable = Params.get(module, 'radio_enable')
 

    # Local parameters
    prep_list = ["lvp_prep"]
    enable_screenshot = '1'

    # If training mode, just 1 fast loop
    if training_mode == "1":
        duration = '10'
        
    airplane_enabled_duration = int(duration) + 15

    def prepCheck(self):
        assert_list = ""
        
        # Check if preps ran
        assert_list += self.checkPrepStatus(self.prep_list)
        
        if assert_list != "":
            self._assert(assert_list)


    def setUp(self):
        prep_status = self.checkPrepStatus(self.prep_list)
        if prep_status !="":
             self._assert(prep_status)

        scenarios.app_scenario.Scenario.setUp(self, callback_test_begin="")
        logging.info("Launching WinAppDriver.exe on DUT.")

        self._call([(self.dut_exec_path + "\\WindowsApplicationDriver\\WinAppDriver.cmd"), (self.dut_ip + " " + self.app_port)], blocking=False)
        time.sleep(1)

        logging.info("Starting App")
        self.driver = self.launchApp() 
        self.driver.maximize_window()
        self.playMovie()

        if self.airplane_mode == '1':
            try:
                # Sete up 2nd config -Prerun command string for lvp_wrapper.cmd
                override_str = '[{\'Scenario\': \'' + self.module + '\'}]'
                config_str = os.path.join(self.dut_exec_path, "config_check.ps1 -Prerun -LogFile " + self.dut_data_path, self.testname + "_ConfigPre") + " -OverrideString " + '\\\"' + override_str + '\\\""'
                logging.info("CONFIG_STR: " + config_str)
                logging.info("Enabling airplane mode for " + str(self.airplane_enabled_duration) + " seconds.")
                if self.radio_enable == '0':
                    logging.info("cmd.exe /C " + os.path.join(self.dut_exec_path + "lvp_resources" + "lvp_wrapper.cmd") + ' ' + str(self.airplane_enabled_duration) + ' ' + self.dut_exec_path + ' ' + "APM " + config_str)
                    self._call(["cmd.exe", "/C " + os.path.join(self.dut_exec_path, "lvp_resources", "lvp_wrapper.cmd") + ' ' + str(self.airplane_enabled_duration) + ' ' + self.dut_exec_path + ' ' + "APM " + config_str], blocking = False) 
                else:
                    logging.info("cmd.exe /C " + os.path.join(self.dut_exec_path + "lvp_resources" + "lvp_wrapper.cmd") + ' ' + str(self.airplane_enabled_duration) + ' ' + self.dut_exec_path + ' ' + "RadioEnable " + config_str)
                    self._call(["cmd.exe", "/C " + os.path.join(self.dut_exec_path, "lvp_resources", "lvp_wrapper.cmd") + ' ' + str(self.airplane_enabled_duration) + ' ' + self.dut_exec_path + ' ' + "RadioEnable " + config_str], blocking = False) 
            except:
                pass

        # Delay to let airplane mode enable
        time.sleep(10)


        # Start recording power
        self._callback(Params.get('global', 'callback_test_begin'))


    def runTest(self):
        logging.info("Duration: " + self.duration + " sec.")
        # Let play for specified duration
        time.sleep(float(self.duration))
 
    def launchApp(self):
        logging.info("Launching LVP")
        desired_caps = {}
        desired_caps["app"] = "Microsoft.ZuneVideo_8wekyb3d8bbwe!microsoft.ZuneVideo"

        driver = self._launchApp(desired_caps)
        
        if self.training_mode == "1":
            time.sleep(5)
            # Look to see if video library access pop-up exists
            try:
                driver.find_element_by_name("Yes").click()
                time.sleep(2)
            except NoSuchElementException as EX:
                pass
            # Look to see if Got it pop-up exists
            try:
                logging.info("Click Got it pop-up")
                driver.find_element_by_name("What's new popup got it").click()
                time.sleep(2)
            except NoSuchElementException as EX:
                pass

        return driver

    def playMovie(self):
        logging.info("Playing " + self.title)

        # Navigate to Personal menu
        #self.driver.find_element_by_xpath('//*[contains(@Name,"Personal")]').click()
        self.driver.find_element_by_name("Personal").click()

        # Find specified movie title	
        self.driver.find_element_by_xpath('//*[contains(@Name,"' + self.title + '")]').click()
        time.sleep(4)

        try:
            logging.info("Selecting Repeat button on DUT.")
            self.driver.find_element_by_name("More").click()
            time.sleep(1.3)

            self.driver.find_element_by_name("Turn repeat on").click()
            time.sleep(2)
        except NoSuchElementException as EX:
            logging.info("Repeat button not found.")
            time.sleep(2)
            pass

        # set to full screen
        try:
            self.driver.find_element_by_name("Full Screen").click()
        except:
            logging.info("Did not find the button for full screen.")

        # Move mouse away from the menu
        time.sleep(5)
        ActionChains(self.driver).move_by_offset(150, -150).click().perform()

    def tearDown(self):
        logging.info("Performing teardown.")        
        # Take screenshot at end of loop to make sure everything was closed properly
        if self.enable_screenshot == '1' and Params.get("global", "local_execution") == "1":
            self._screenshot(name="end_screen.png")

       # Stop recording power
        self._callback(Params.get('global', 'callback_test_end'))
        if self.airplane_mode =='1':
            try:
                if self.radio_enable == '0':
                    logging.info("cmd.exe /C " + os.path.join(self.dut_exec_path, "lvp_resources\AirplaneMode.exe -Disable"))
                    self._call(["cmd.exe", "/C " + os.path.join(self.dut_exec_path, "lvp_resources", "AirplaneMode.exe") + " -Disable"], blocking = False) 
                else:
                    logging.info("cmd.exe /C " + os.path.join(self.dut_exec_path, "lvp_resources\RadioEnable.exe -Enable"))
                    self._call(["cmd.exe", "/C " + os.path.join(self.dut_exec_path, "lvp_resources", "radioenable.exe") + " -Enable"], blocking = False) 
            except:
                pass

        # if self.airplane_mode == '1':
        #     try:
        #         # Disable airplanemode and enable radio if disabled
        #         if self.radio_enable == '0':
        #             logging.info("cmd.exe /C " + os.path.join(self.dut_exec_path + "lvp_resources" + "lvp_wrapper.cmd") + ' ' + duration + ' ' + self.dut_exec_path + ' ' + "APM ")
        #             self._call(["cmd.exe", "/C " + os.path.join(self.dut_exec_path, "lvp_resources", "lvp_wrapper.cmd") + ' ' + duration + ' ' + self.dut_exec_path + ' ' + "APM"], blocking = False) 
        #         else:
        #             logging.info("cmd.exe /C " + os.path.join(self.dut_exec_path + "lvp_resources" + "lvp_wrapper.cmd") + ' ' + duration + ' ' + self.dut_exec_path + ' ' + "RadioEnable")
        #             self._call(["cmd.exe", "/C " + os.path.join(self.dut_exec_path, "lvp_resources", "lvp_wrapper.cmd") + ' ' + duration + ' ' + self.dut_exec_path + ' ' + "RadioEnable"], blocking = False) 
        #     except:
        #         pass


        # Allow plenty of time for wifi to come back up
        time.sleep(30)
        scenarios.app_scenario.Scenario.tearDown(self, callback_test_end="")

        self._kill("WinAppDriver.exe")

    def kill(self):
        try:
            logging.debug("Killing Video.UI.exe")
            self._kill("Video.UI.exe")
        except:
            pass 
        