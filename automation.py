import difflib
import os
import re
import shutil
import time
import traceback
from getpass import getuser
from urllib.parse import urlparse

import cv2
import numpy as np
import pyautogui
import requests
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait


class Automation:
    def __init__(self, config):
        self.config = config
        self.pause_record = False
        self.scrollable_actions_list = []
        self.jquery = requests.get("https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js").text
        self.record_screen_loop = False
        options = webdriver.ChromeOptions()
        options.add_argument(f"--user-data-dir=C:\\Users\\{getuser()}\\AppData\\Local\\Google\\Chrome\\User Data")
        self.driver = webdriver.Chrome(executable_path="chromedriver.exe")
        self.actions = ActionChains(self.driver)
        if not os.path.exists("screenshots"):
            os.mkdir("screenshots")
        if not os.path.exists("screen_records"):
            os.mkdir("screen_records")
        if not os.path.exists("All_Screen_Records"):
            os.mkdir("All_Screen_Records")
        if not os.path.exists("All_ScreenShots"):
            os.mkdir("All_ScreenShots")
        if not os.path.exists("TSV_Files"):
            os.mkdir("TSV_Files")

    def open_website(self, url, window):
        window.setDisabled(True)
        try:
            self.driver.get(url)
            self.driver.maximize_window()
            self.browser_width, self.browser_height = pyautogui.size()
            self.driver.execute_script("window.focus();")
        except Exception as e:
            print(f"Error: {e}")
        window.setEnabled(True)

    def close_browser(self):
        self.driver.quit()

    def take_screenshot(self, window, prefix, suffix, second_instruction, delay):
        window.setDisabled(True)
        try:
            time.sleep(delay)
            url = self.driver.current_url
            parsed_url = urlparse(url)
            path = parsed_url.path
            domain = parsed_url.hostname
            domain = re.sub("[^0-9a-zA-Z]+", "_", domain)
            if not os.path.exists(f"screenshots\\{domain}"):
                os.mkdir(f"screenshots\\{domain}")
            filename = f"{prefix}1_{suffix}_{domain}.png"
            file_path = os.path.join(os.curdir, "screenshots", domain)
            if os.path.exists(os.path.join(file_path, filename)):
                files = list(filter(lambda file: re.search(f"{prefix}(\d+)?_{suffix}_{domain}.png", file),
                                    os.listdir(file_path)))
                n_list = [1]
                for file in files:
                    p = re.findall(f"{prefix}(\d+)?", file)
                    if p:
                        n = int(p[0]) if p[0] else 1
                        n_list.append(n)
                filename = f"{prefix}{max(n_list) + 1}_{suffix}_{domain}.png"
            file_full_path = os.path.join(file_path, filename)
            with open(os.path.join(os.curdir, "TSV_Files", f"{domain}.tsv"), "a") as f:
                f.write(f"{filename}\t{second_instruction}\n")
            self.driver.save_screenshot(file_full_path)
            shutil.copyfile(file_full_path, os.path.join(os.curdir, "All_ScreenShots", filename))
        except Exception as e:
            print(f"Error: {e}")
        window.setEnabled(True)

    def take_scrolled_screenshot(self, window, prefix, suffix, second_instruction, delay):
        window.setDisabled(True)
        try:
            options = webdriver.ChromeOptions()
            options.headless = True
            driver = webdriver.Chrome(executable_path="chromedriver.exe", options=options)
            driver.get(self.driver.current_url)
            s = lambda x: driver.execute_script('return document.body.parentNode.scroll' + x)
            driver.set_window_size(s('Width'), s('Height'))

            time.sleep(delay)
            url = self.driver.current_url
            parsed_url = urlparse(url)
            path = parsed_url.path
            domain = parsed_url.hostname
            domain = re.sub("[^0-9a-zA-Z]+", "_", domain)
            if not os.path.exists(f"screenshots\\{domain}"):
                os.mkdir(f"screenshots\\{domain}")
            filename = f"{prefix}1_{suffix}_{domain}.png"
            file_path = os.path.join(os.curdir, "screenshots", domain)
            if os.path.exists(os.path.join(file_path, filename)):
                files = list(filter(lambda file: re.search(f"{prefix}(\d+)?_{suffix}_{domain}.png", file),
                                    os.listdir(file_path)))
                n_list = [1]
                for file in files:
                    p = re.findall(f"{prefix}(\d+)?", file)
                    if p:
                        n = int(p[0]) if p[0] else 1
                        n_list.append(n)
                filename = f"{prefix}{max(n_list) + 1}_{suffix}_{domain}.png"
            file_full_path = os.path.join(file_path, filename)
            with open(os.path.join(os.curdir, "TSV_Files", f"{domain}.tsv"), "a") as f:
                f.write(f"{filename}\t{second_instruction}\n")
            driver.find_element(By.TAG_NAME, 'body').screenshot(file_full_path)
            shutil.copyfile(file_full_path, os.path.join(os.curdir, "All_ScreenShots", filename))
            driver.quit()
        except Exception as e:
            traceback.print_exc()
        window.setEnabled(True)

    def refresh_page(self):
        self.driver.refresh()

    def record_screen(self, window, prefix, suffix, second_instruction, delay):
        try:
            url = self.driver.current_url
            parsed_url = urlparse(url)
            path = parsed_url.path
            domain = parsed_url.hostname
            domain = re.sub("[^0-9a-zA-Z]+", "_", domain)
            if not os.path.exists(f"screen_records\\{domain}"):
                os.mkdir(f"screen_records\\{domain}")
            filename = f"{prefix}1_{suffix}_{domain}.mp4"
            file_path = os.path.join(os.curdir, "screen_records", domain)
            if os.path.exists(file_path):
                files = list(filter(lambda file: re.search(f"{prefix}(\d+)?_{suffix}_{domain}.mp4", file),
                                    os.listdir(file_path)))
                n_list = [1]
                for file in files:
                    p = re.findall(f"{prefix}(\d+)?", file)
                    if p:
                        n = int(p[0]) if p[0] else 1
                        n_list.append(n)
                filename = f"{prefix}{max(n_list) + 1}_{suffix}_{domain}.mp4"
                # Specify resolution
                resolution = pyautogui.size()

                # Specify video codec
                codec = cv2.VideoWriter_fourcc(*"mp4v")

                # Specify frames rate. We can choose any
                # value and experiment with it
                fps = 20

                # Creating a VideoWriter object
                file_full_path = os.path.join(file_path, filename)
                out = cv2.VideoWriter(file_full_path, codec, fps, resolution)
                with open(os.path.join(os.curdir, "TSV_Files", f"{domain}.tsv"), "a") as f:
                    f.write(f"{filename}\t{second_instruction}\n")
                self.record_screen_loop = True

                Xs = [0, 15, 9, 12, 8, 5, 0, 0]
                Ys = [0, 13, 14, 20, 21, 15, 18, 0]

                while self.record_screen_loop:
                    if not self.pause_record:
                        mouse_x, mouse_y = pyautogui.position()
                        # Take screenshot using PyAutoGUI
                        img = pyautogui.screenshot()
                        # Convert the screenshot to a numpy array
                        frame = np.array(img)

                        # Convert it from BGR(Blue, Green, Red) to
                        # RGB(Red, Green, Blue)
                        frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                        Xthis = [mouse_x + 1 * x for x in Xs]
                        Ythis = [mouse_y + 1 * y for y in Ys]
                        points = list(zip(Xthis, Ythis))
                        points = np.array(points, 'int32')
                        cv2.fillPoly(frame, [points], color=[255, 255, 255])
                        cv2.polylines(frame, [points], isClosed=True, color=[0, 0, 0])
                        # Write it to the output file
                        out.write(frame)
                        # Optional: Display the recording screen
                        # cv2.imshow('Live', frame)
                    else:
                        time.sleep(0.1)
                self.pause_record = False
                # Release the Video writer
                out.release()
                shutil.copyfile(file_full_path, os.path.join(os.curdir, "All_Screen_Records", filename))
                # Destroy all windows
                cv2.destroyAllWindows()
        except Exception as e:
            print(f"Error: {e}")

    def get_element_at_mouse(self, window):
        # window.setDisabled(True)
        try:
            self.driver.execute_script(self.jquery)
            self.driver.execute_script("""
                    function getXPath( element ){
                        var xpath = '';
                        for ( ; element && element.nodeType == 1; element = element.parentNode )
                        {
                            var id = $(element.parentNode).children(element.tagName).index(element) + 1;
                            id > 1 ? (id = '[' + id + ']') : (id = '');
                            xpath = '/' + element.tagName.toLowerCase() + id + xpath;
                        }
                        return xpath;
                    };
                    $("body").one("mousemove", function(e) {
                    var mouseX = e.pageX - window.pageXOffset;
                    var mouseY = e.pageY - window.pageYOffset;
                    var element = document.elementFromPoint(mouseX, mouseY);
                    console.log(getXPath(element));
                    localStorage.selectedEvent = getXPath(element);
                    })""")
            time.sleep(3)
            # window.setEnabled(True)
            return self.driver.execute_script("""var event= localStorage.selectedEvent;
                                                localStorage.selectedEvent = null;
                                              return event;""")
        except Exception as e:
            print(f"Error: {e}")
        # window.setEnabled(True)
        return None

    def play_recorded_actions(self, window, delay):
        # window.setDisabled(True)
        print(self.scrollable_actions_list)
        wait = WebDriverWait(self.driver, 5)
        # new_page = self.driver.window_handles[1]
        # self.driver.switch_to.window(new_page)
        try:
            window_size = pyautogui.size()
            for x_path, click_or_not in self.scrollable_actions_list:
                try:
                    time.sleep(delay)
                    wait.until(EC.presence_of_element_located((By.XPATH, x_path)))
                    element = self.driver.find_element(By.XPATH, x_path)
                    if element:
                        element_size = element.size
                        pos = element.location
                        action = self.actions.move_to_element(element)
                        if click_or_not:
                            action.click()
                        action.perform()
                        time.sleep(1)
                        # location = element.location_once_scrolled_into_view
                        # page_offset = self.driver.execute_script("return window.pageYOffset;")
                        # ih = self.driver.execute_script("return window.innerHeight;")
                        # oh = self.driver.execute_script("return window.outerHeight;")
                        # print(ih,oh)
                        # chrome_offset = window_size[1]-ih
                        # ex = pos["x"]+8+element_size["width"]/2
                        # ey = (pos["y"]-page_offset)+chrome_offset-(element_size["height"]/2)
                        # pyautogui.moveTo(ex,ey,2)
                        # print(pyautogui.size(),self.driver.get_window_size(),self.driver.get_window_position(),element.location,page_offset)
                    else:
                        print("unknown element")
                except Exception as e:
                    print(f"Error: {e}")
        except Exception as e:
            print(f"Error: {e}")
        # window.setEnabled(True)

    def enter_user_details(self, value):
        try:
            self.driver.execute_script(self.jquery)
            self.driver.execute_script("""
                           $("body").one("mousemove", function(e) {
                           var mouseX = e.pageX - window.pageXOffset;
                           var mouseY = e.pageY - window.pageYOffset;
                           var element = document.elementFromPoint(mouseX, mouseY);
                           $(element).val("%s");
                           })""" % value)
        except Exception as e:
            print(f"Error: {e}")

    def click_on_element(self):
        try:
            self.driver.execute_script(self.jquery)
            self.driver.execute_script("""
                           $("body").one("mousemove", function(e) {
                           var mouseX = e.pageX - window.pageXOffset;
                           var mouseY = e.pageY - window.pageYOffset;
                           var element = document.elementFromPoint(mouseX, mouseY);
                           element.click();
                           })""")
        except Exception as e:
            print(f"Error: {e}")

    def fill_form_automatically(self, window, user_form, stop_record_button):
        user_form_keys = user_form.keys()
        forms = self.driver.find_elements(By.TAG_NAME, "form")

        for form in forms:
            inputs = form.find_elements(By.TAG_NAME, "input")
            drop_downs = form.find_elements(By.TAG_NAME, "select")
            for drop_down in drop_downs:
                id = drop_down.get_attribute("id")
                # input_class = drop_down.get_attribute("class")
                placeholder = drop_down.get_attribute("placeholder")
                informations = [placeholder, id]
                for information in informations:
                    information = str(information).lower()
                    if information:
                        cutoff = 0.7
                        if len(information) > 40:
                            cutoff = 0.1
                        elif len(information) > 20:
                            cutoff = 0.3
                        elif len(information) > 10:
                            cutoff = 0.4
                        elif len(information) > 5:
                            cutoff = 0.6
                        assumed = difflib.get_close_matches(information, user_form_keys, cutoff=cutoff, n=1)
                        if assumed:
                            try:
                                data_to_fill = user_form[assumed[0]]
                                try:
                                    Select(drop_down).select_by_visible_text(data_to_fill)
                                    time.sleep(1)
                                except Exception as e:
                                    print(f"Error: {e}")

                            except Exception as e:
                                print(f"Error: {e}")
                            break
            for inp in inputs:
                id = inp.get_attribute("id")
                # input_class = inp.get_attribute("class")
                placeholder = inp.get_attribute("placeholder")
                informations = [placeholder, id]
                for information in informations:
                    information = str(information).lower()
                    if information:
                        cutoff = 0.6
                        if len(information) > 40:
                            information = information[:20]
                            cutoff = 0.2
                        elif len(information) > 20:
                            cutoff = 0.3
                        elif len(information) > 10:
                            cutoff = 0.4
                        elif len(information) > 5:
                            cutoff = 0.5
                        assumed = difflib.get_close_matches(information, user_form_keys, cutoff=cutoff, n=1)
                        if assumed:
                            try:
                                data_to_fill = user_form[assumed[0]]
                                inp.clear()
                                for ch in data_to_fill:
                                    inp.send_keys(ch)
                                    time.sleep(0.03)
                                # print(information, assumed)
                                time.sleep(0.5)
                            except Exception as e:
                                print(f"Error: {e}")
                            break
        window.stop_record()
