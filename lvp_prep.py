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
# Prep for Local Video Playback
# 
# Setup instructions:
##

import builtins
import os
import logging
import scenarios.app_scenario
from parameters import Params
import time
from appium import webdriver
from selenium.common.exceptions import NoSuchElementException


class LvpPrep(scenarios.app_scenario.Scenario):
    module = __module__.split('.')[-1]
    # Set default parameters
    # Params.setDefault(module, 'video_file', 'tears_of_steel_1080p_24fps_10mbps_loop_30min.mp4')
    Params.setDefault(module, 'video_file', 'ToS-4k-1920.mov')
    # default time to wait for application to populate its video library
    Params.setDefault(module, 'duration', '15')  # Seconds

    # Get parameters
    video_file = Params.get(module, 'video_file')
    dut_architecture = Params.get('global', 'dut_architecture')
    # wait time for application to populate its video library
    duration = Params.get(module, 'duration')

    is_prep = True


    def runTest(self):
        if Params.get("global", "local_execution") == "0":
            self.userprofile = self._call(["cmd.exe", "/C echo %USERPROFILE%"])
        else:
            self.userprofile = os.environ['USERPROFILE']    


        source = os.path.join("scenarios", "lvp_resources", self.video_file)
        dest = os.path.join(self.userprofile, "Videos")
        dest_file = os.path.join(dest, self.video_file)
        # Check if video file already exists on DUT
        if self._check_remote_file_exists(dest_file, False):
            logging.info("Movie file " + self.video_file + " already found on DUT.  Skipping upload")
        else:
            logging.info("Uploading movie file " + self.video_file + " to " + dest)
            self._upload(source, dest)
        # Copy over resources
        self._remote_make_dir(self.dut_exec_path + "\\lvp_resources")
        logging.info("Uploading additional test files to " + self.dut_exec_path + "\\lvp_resources")
        self._upload("scenarios\\lvp_resources\\lvp_wrapper.cmd", (self.dut_exec_path + "\\lvp_resources"))
        self._upload("scenarios\\lvp_resources\\" + self.dut_architecture + "\\AirplaneMode.*", (self.dut_exec_path + "\\lvp_resources"))
        self._upload("scenarios\\lvp_resources\\" + self.dut_architecture + "\\RadioEnable.*", (self.dut_exec_path + "\\lvp_resources"))
        self._upload("scenarios\\lvp_resources\\sleep.exe", (self.dut_exec_path + "\\lvp_resources"))

        # Code added to launch video player software in order to give it time to populate its video library.  Avoids an issue experienced by some devices where LVP will fail to find the video file before failing the test
        logging.info("Launching WinAppDriver.exe on DUT.")

        self._call([(self.dut_exec_path + "\\WindowsApplicationDriver\\WinAppDriver.cmd"), (self.dut_ip + " " + self.app_port)], blocking=False)
        time.sleep(1)

        logging.info("Launching LVP")
        desired_caps = {}
        desired_caps["app"] = "Microsoft.ZuneVideo_8wekyb3d8bbwe!microsoft.ZuneVideo"

        driver = self._launchApp(desired_caps)

        # waiting for a specified period of time to allow the software to update their video library
        logging.info("Waiting for " + self.duration + " seconds to allow the video library to populate.")
        time.sleep(float(self.duration))
        # Look to see if 'What's New' pop-up exists
        try:
            logging.info("Checking for What's New pop-up.")
            driver.find_element_by_name("What's new popup got it").click()
            logging.info("Pop-up dismissed.")
            time.sleep(2)
        except NoSuchElementException:
            logging.info("Pop-up not found.")



    def tearDown(self):
        self.createPrepStatusControlFile()
        scenarios.app_scenario.Scenario.tearDown(self)

        self._kill("WinAppDriver.exe")



    def kill(self):
        try:
            logging.debug("Killing Video.UI.exe")
            self._kill("Video.UI.exe")
        except:
            pass
