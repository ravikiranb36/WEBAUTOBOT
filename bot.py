import json
import os
import subprocess
import sys
import time
from threading import Thread

import pyautogui
from PySide6 import QtCore
from PySide6.QtCore import QSortFilterProxyModel
from PySide6.QtCore import Qt
from PySide6.QtGui import QShortcut
from PySide6.QtWidgets import (QApplication, QLabel, QLineEdit, QComboBox, QHBoxLayout, QVBoxLayout,
                               QWidget, QFormLayout, QCompleter, QPushButton, QPlainTextEdit, QSpinBox, QMessageBox,
                               QStatusBar, QCheckBox, QMenuBar, QMenu, QMainWindow)

from automation import Automation


class ExtendedComboBox(QComboBox):
    def __init__(self, parent=None):
        super(ExtendedComboBox, self).__init__(parent)

        self.setFocusPolicy(Qt.StrongFocus)
        self.setEditable(True)

        # add a filter model to filter matching items
        self.pFilterModel = QSortFilterProxyModel(self)
        self.pFilterModel.setFilterCaseSensitivity(Qt.CaseInsensitive)
        self.pFilterModel.setSourceModel(self.model())

        # add a completer, which uses the filter model
        self.completer = QCompleter(self.pFilterModel, self)
        # always show all (filtered) completions
        self.completer.setCompletionMode(QCompleter.UnfilteredPopupCompletion)
        self.setCompleter(self.completer)

        # connect signals
        self.lineEdit().textEdited.connect(self.pFilterModel.setFilterFixedString)

    # on model change, update the models of the filter and completer as well
    def setModel(self, model):
        super(ExtendedComboBox, self).setModel(model)
        self.pFilterModel.setSourceModel(model)
        self.completer.setModel(self.pFilterModel)

    # on model column change, update the model column of the filter and completer as well
    def setModelColumn(self, column):
        self.completer.setCompletionColumn(column)
        self.pFilterModel.setFilterKeyColumn(column)
        super(ExtendedComboBox, self).setModelColumn(column)


class BotWindow(QWidget):
    def __init__(self):
        """
        Initialises Bot Window
        """
        self.instructuns_index = 0
        self.websites_index = 0
        self.non_scrollable_actions_list = []
        self.user_details = []
        try:
            with open("config.json") as f:
                self.config = json.load(f)
                self.user_details = self.config["user_details"]
        except Exception as e:
            print(f"Error: {e}")
        super().__init__()

        # Selenium automation object
        self.automation = Automation(self.config)

        with open("config.json", "r") as f:
            self.config = json.load(f)
        # self.configure_working_dir()
        self.urls_list = self.get_websites_list()
        self.instructions = list(self.config["instructions"].keys())
        # self.prefix_list = self.get_prefixes_list()
        self.main_layout = QVBoxLayout()
        self.create_widgets()
        self.setWindowTitle("Bot")
        self.setLayout(self.main_layout)
        self.show()

    # def configure_working_dir(self):
    #     work_dir = self.config["work_dir"]
    #     if work_dir:
    #         return
    #     else:
    #         d = QFileDialog.getExistingDirectory(self,"Select Bot Working folder")
    #         self.config["work_dir"] = d
    #         with open("config.json", "w") as f:
    #             json.dump(self.config, f)

    def create_widgets(self):
        self.setFixedSize(500, 650)

        # Menu Bar
        open_folder_hbox_layout = QHBoxLayout()
        open_scr_shot_folder_button = QPushButton("Open Screen Shots")
        open_scr_shot_folder_button.clicked.connect(
            lambda: subprocess.call(f'explorer {os.path.join(os.curdir, "All_ScreenShots")}', shell=True)
        )
        open_folder_hbox_layout.addWidget(open_scr_shot_folder_button)

        open_scr_rec_folder_button = QPushButton("Open Screen Records")
        open_scr_rec_folder_button.clicked.connect(
            lambda: subprocess.call(f'explorer {os.path.join(os.curdir, "All_Screen_Records")}', shell=True)
        )
        open_folder_hbox_layout.addWidget(open_scr_rec_folder_button)
        self.main_layout.addLayout(open_folder_hbox_layout)

        self.form_layout = QFormLayout()

        self.run_name = QLineEdit()
        self.form_layout.addRow(QLabel("Run name"), self.run_name)

        self.websites_combobox = ExtendedComboBox()
        self.websites_combobox.addItems(self.urls_list)
        self.form_layout.addRow(QLabel("Select URL's"), self.websites_combobox)

        self.home_page_button = QPushButton("Goto Homepage")
        self.home_page_button.clicked.connect(self.goto_home_page)
        self.form_layout.addWidget(self.home_page_button)

        self.refresh_hbox_layout = QHBoxLayout()
        self.refresh_button = QPushButton("Refresh Page")
        self.refresh_button.setDisabled(True)
        self.refresh_button.clicked.connect(self.automation.refresh_page)
        self.refresh_hbox_layout.addWidget(self.refresh_button)
        self.next_site_button = QPushButton("Next Site")
        self.next_site_button.clicked.connect(self.goto_next_site)
        self.next_site_button.setDisabled(True)
        self.refresh_hbox_layout.addWidget(self.next_site_button)

        # Radio button to set keep on top
        self.keep_on_top_check_box = QCheckBox("Keep on top")
        self.keep_on_top_check_box.clicked.connect(self.toggle_window_on_top)
        self.keep_on_top_check_box.setChecked(True)
        self.toggle_window_on_top()
        self.refresh_hbox_layout.addWidget(self.keep_on_top_check_box)
        self.main_layout.addLayout(self.refresh_hbox_layout)

        self.instruction_box = QPlainTextEdit()
        self.instruction_box.setReadOnly(True)
        self.form_layout.addRow(QLabel("Instruction"), self.instruction_box)

        self.screenshot_delay = QSpinBox()
        self.screenshot_delay.setRange(0, 10)
        self.form_layout.addRow(QLabel("ScreenShot delay"), self.screenshot_delay)

        self.prefixes_combobox = ExtendedComboBox()
        # self.prefixes_combobox.addItems(self.prefix_list)
        self.form_layout.addRow(QLabel("Prefixes"), self.prefixes_combobox)

        self.main_layout.addLayout(self.form_layout)

        self.screenshot_hbox_layout1 = QHBoxLayout()
        self.take_screenshot_button = QPushButton("Get Screenshot")
        self.take_screenshot_button.clicked.connect(self.get_screenshot)
        self.take_screenshot_button.setDisabled(True)
        self.take_scrolling_screenshot_button = QPushButton("Get Scrolling Screenshot")
        self.take_scrolling_screenshot_button.setDisabled(True)
        self.take_scrolling_screenshot_button.clicked.connect(self.get_scrolling_screenshot)
        self.screenshot_hbox_layout1.addWidget(self.take_screenshot_button)
        self.screenshot_hbox_layout1.addWidget(self.take_scrolling_screenshot_button)
        self.main_layout.addLayout(self.screenshot_hbox_layout1)

        # Action record row 1
        self.action_record_hbox_layout1 = QHBoxLayout()
        self.add_click_check_box = QCheckBox("Add click")
        self.action_record_hbox_layout1.addWidget(self.add_click_check_box)
        self.record_actions_check_box = QCheckBox("Record Actions")
        self.action_record_hbox_layout1.addWidget(self.record_actions_check_box)
        self.play_action_delay = QSpinBox()
        self.play_action_delay.setRange(1, 10)
        self.action_record_hbox_layout1.addWidget(self.play_action_delay)
        self.add_action_button = QPushButton("Add Action")
        self.add_action_button.setDisabled(True)
        self.add_action_button.clicked.connect(self.add_action)
        self.action_record_hbox_layout1.addWidget(self.add_action_button)
        self.main_layout.addLayout(self.action_record_hbox_layout1)

        # Action record row 2
        self.action_record_hbox_layout2 = QHBoxLayout()
        self.clear_action_button = QPushButton("Clear Actions")
        self.clear_action_button.setDisabled(True)
        self.clear_action_button.clicked.connect(self.clear_actions)
        self.action_record_hbox_layout2.addWidget(self.clear_action_button)
        self.play_action_button = QPushButton("Play Action")
        self.play_action_button.setDisabled(True)
        self.play_action_button.clicked.connect(self.play_actions)
        self.action_record_hbox_layout2.addWidget(self.play_action_button)
        self.main_layout.addLayout(self.action_record_hbox_layout2)

        self.action_hbox_layout = QHBoxLayout()
        self.previous_button = QPushButton("Previous")
        self.previous_button.clicked.connect(self.goto_prev_instruction)
        self.previous_button.setDisabled(True)
        self.skip_button = QPushButton("Skip")
        self.skip_button.clicked.connect(self.skip_instruction)
        self.skip_button.setDisabled(True)
        self.next_button = QPushButton("Next")
        self.next_button.clicked.connect(self.goto_next_instruction)
        self.next_button.setDisabled(True)
        self.action_hbox_layout.addWidget(self.previous_button)
        self.action_hbox_layout.addWidget(self.skip_button)
        self.action_hbox_layout.addWidget(self.next_button)

        self.main_layout.addLayout(self.action_hbox_layout)

        self.screen_recoreder_layout = QHBoxLayout()
        self.start_record_button = QPushButton("Record")
        self.start_record_button.setDisabled(True)
        self.start_record_button.clicked.connect(self.record_screen)
        self.screen_recoreder_layout.addWidget(self.start_record_button)
        self.pause_record_button = QPushButton("Pause")
        self.pause_record_button.clicked.connect(self.pause_record)
        self.pause_record_button.setDisabled(True)
        self.screen_recoreder_layout.addWidget(self.pause_record_button)
        self.stop_record_button = QPushButton("Stop Record")
        self.stop_record_button.clicked.connect(self.stop_record)
        self.stop_record_button.setDisabled(True)
        self.screen_recoreder_layout.addWidget(self.stop_record_button)
        self.main_layout.addLayout(self.screen_recoreder_layout)

        # user details HBOX  layout
        self.user_details_hbox_layout = QHBoxLayout()
        self.user_details_hbox_layout.addWidget(QLabel("Select Form Details"))
        # Details name combo box
        self.details_name_combobox = ExtendedComboBox()
        self.user_details_hbox_layout.addWidget(self.details_name_combobox)
        self.details_name_combobox.addItems(self.user_details.keys())
        self.details_name_combobox.currentIndexChanged.connect(self.user_detail_name_changed)
        # Details list combo box
        self.details_list_combobox = ExtendedComboBox()
        self.user_details_hbox_layout.addWidget(self.details_list_combobox)
        self.user_detail_name_selected = self.details_name_combobox.currentText()
        if self.user_detail_name_selected in self.user_details:
            self.details_list_combobox.addItems(self.user_details[self.user_detail_name_selected])

        self.main_layout.addLayout(self.user_details_hbox_layout)

        # Auto fill form HBOX
        self.auto_fill_form_hbox_layout = QHBoxLayout()
        self.auto_fill_form_hbox_layout.addWidget(QLabel("Select Form"))
        # Auto fill form Combobox
        self.auto_fill_form_combobox = ExtendedComboBox()
        forms = self.config["user_form"].keys()
        self.auto_fill_form_combobox.addItems(forms)
        self.auto_fill_form_hbox_layout.addWidget(self.auto_fill_form_combobox)
        # Auto fill form button
        self.auto_fill_form_button = QPushButton("Auto Fill Form")
        self.auto_fill_form_hbox_layout.addWidget(self.auto_fill_form_button)
        self.auto_fill_form_button.clicked.connect(self.fill_form_automatically)

        self.main_layout.addLayout(self.auto_fill_form_hbox_layout)

        # Status Bar
        self.status_bar = QStatusBar()
        self.main_layout.addWidget(self.status_bar)

        # Short Cut Keys
        QShortcut('Ctrl+H', self).activated.connect(
            lambda: self.home_page_button.click() if self.home_page_button.isEnabled() else
            self.display_message_box("Home page button not enabled"))
        QShortcut('Ctrl+N', self).activated.connect(
            lambda: self.next_site_button.click() if self.next_site_button.isEnabled() else
            self.display_message_box("Next site button not enabled"))
        QShortcut('Ctrl+S', self).activated.connect(
            lambda: self.take_screenshot_button.click() if self.take_screenshot_button.isEnabled() else
            self.display_message_box("Screenshot button not enabled"))
        QShortcut('Ctrl+Shift+S', self).activated.connect(
            lambda: self.take_scrolling_screenshot_button.click() if self.take_scrolling_screenshot_button.isEnabled() else
            self.display_message_box("Screenshot button not enabled"))
        QShortcut('Ctrl+R', self).activated.connect(
            lambda: self.pause_record_button.click() if self.pause_record_button.isEnabled() else
            self.display_message_box("Pause button not enabled"))
        QShortcut('Ctrl+A', self).activated.connect(
            lambda: self.add_action_button.click() if self.add_action_button.isEnabled() else
            self.display_message_box("Add Action not enabled"))
        QShortcut('Ctrl+Shift+V', self).activated.connect(
            lambda: self.start_record_button.click() if self.start_record_button.isEnabled() else
            self.stop_record_button.click() if self.stop_record_button.isEnabled() else
            self.display_message_box("Start or Stop record button not enabled"))
        QShortcut('Ctrl+P', self).activated.connect(
            lambda: self.play_action_button.click() if self.play_action_button.isEnabled() else
            self.display_message_box("Play action button not enabled"))
        QShortcut('Ctrl+C', self).activated.connect(
            lambda: self.add_click_check_box.toggle())
        QShortcut('Ctrl+L', self).activated.connect(
            lambda: self.click_on_element())
        QShortcut('Ctrl+M', self).activated.connect(
            lambda: self.showNormal() if self.isMinimized() else self.showMinimized())
        QShortcut('Ctrl+D', self).activated.connect(
            lambda: self.move(pyautogui.size()[0] - 30, pyautogui.size()[1] - 30))
        QShortcut('Ctrl+U', self).activated.connect(
            lambda: self.move(pyautogui.size()[0] / 2, pyautogui.size()[1] / 3))
        QShortcut('Ctrl+B', self).activated.connect(
            lambda: self.automation.driver.back())
        QShortcut('Ctrl+F', self).activated.connect(
            lambda: self.automation.driver.forward())
        QShortcut('Ctrl+Q', self).activated.connect(
            lambda: QApplication.instance().quit())
        QShortcut('Ctrl+E', self).activated.connect(
            lambda: self.enter_user_details())
        QShortcut('Ctrl+Shift+U', self).activated.connect(
            lambda: self.details_name_combobox.setFocus())
        QShortcut('Ctrl+Shift+L', self).activated.connect(
            lambda: self.details_list_combobox.setFocus())
        QShortcut('Ctrl+Shift+F', self).activated.connect(
            lambda: self.auto_fill_form_button.click() if self.auto_fill_form_button.isEnabled() else
            self.display_message_box("Auto Fill button not enabled"))

    @staticmethod
    def get_websites_list():
        urls = []
        with open("urls.txt") as f:
            for line in f.readlines():
                line = line.strip()
                if line:
                    urls.append(line)
        return urls

    # @staticmethod
    # def get_instructions_list():
    #     instructions = []
    #     with open("instructions.txt") as f:
    #         for line in f.readlines():
    #             line = line.strip()
    #             if line:
    #                 instructions.append(line)
    #     return instructions

    # @staticmethod
    # def get_prefixes_list():
    #     prefixes = []
    #     with open("prefixes.txt") as f:
    #         for line in f.readlines():
    #             line = line.strip()
    #             if line:
    #                 prefixes.append(line)
    #     return prefixes

    def goto_home_page(self):
        self.can_enable_next_button = True
        self.status_bar.showMessage("Wait Home Page is loading...", 3000)
        self.skip_button.setEnabled(True)
        self.instructuns_index = 0
        self.websites_index = self.websites_combobox.currentIndex()
        url = self.websites_combobox.currentText()
        url = "https://" + url
        new_thread = Thread(target=self.automation.open_website, args=(url, self))
        new_thread = Thread(target=self.automation.open_website, args=(url, self))
        new_thread.start()
        # new_thread.join()
        # self.setEnabled(True)
        self.refresh_button.setEnabled(True)
        if not self.set_instruction():
            return
        self.start_record_button.setEnabled(True)
        self.take_screenshot_button.setEnabled(True)
        self.take_scrolling_screenshot_button.setEnabled(True)
        self.next_site_button.setEnabled(True)
        self.add_action_button.setEnabled(True)
        self.clear_action_button.setEnabled(True)
        self.play_action_button.setEnabled(True)

    def goto_next_site(self):
        if self.websites_index >= len(self.urls_list):
            self.websites_index = 0
        else:
            self.websites_index = self.websites_combobox.currentIndex() + 1
        self.websites_combobox.setCurrentIndex(self.websites_index)
        self.status_bar.showMessage(f"Wait opening {self.websites_combobox.currentText()}", 3000)
        self.clear_actions()
        self.goto_home_page()

    def goto_next_instruction(self):
        self.previous_button.setEnabled(True)
        self.clear_actions()
        self.status_bar.showMessage("Going to next instruction", 2000)
        self.next_button.setDisabled(True)
        if self.instructuns_index >= len(self.instructions) - 1:
            self.display_message_box("This is last instruction")
            return
        self.instructuns_index += 1
        self.set_instruction()

    def goto_prev_instruction(self):
        self.next_button.setEnabled(True)
        self.clear_actions()
        self.status_bar.showMessage("Going to Previous instruction", 2000)
        if self.instructuns_index <= 0:
            self.display_message_box("This is first instruction you can't go behind this")
            self.previous_button.setDisabled(True)
            return
        self.instructuns_index -= 1
        self.set_instruction()

    def skip_instruction(self):
        if self.instructuns_index >= len(self.instructions) - 1:
            self.status_bar.showMessage("Instructions are out of index", 2000)
            self.display_message_box("You reached to last instruction")
            return
        self.status_bar.showMessage("Skipping this instruction", 2000)
        self.next_button.setEnabled(True)

    def get_screenshot(self):
        prefix = self.prefixes_combobox.currentText()
        suffix = self.run_name.text()
        if prefix and suffix:
            try:
                second_instruction = self.config["instructions"][self.instructions[self.instructuns_index]][1]
            except:
                second_instruction = ""
            if self.can_enable_next_button:
                self.next_button.setEnabled(True)
            delay = self.screenshot_delay.value()
            self.status_bar.showMessage(f"Wait for {delay}s capturing screenshot", 2000)
            new_thread = Thread(target=self.automation.take_screenshot,
                                args=(self, prefix, suffix, second_instruction, delay))
            new_thread.start()
        else:
            self.display_message_box("Pls check prefix or Run name entered properly")

    def get_scrolling_screenshot(self):
        try:
            prefix = self.prefixes_combobox.currentText()
            suffix = self.run_name.text()
            if prefix and suffix:
                try:
                    second_instruction = self.config["instructions"][self.instructions[self.instructuns_index]][1]
                except:
                    second_instruction = ""
                if self.can_enable_next_button:
                    self.next_button.setEnabled(True)
                delay = self.screenshot_delay.value()
                self.status_bar.showMessage(f"Wait for {delay}s capturing screenshot", 2000)
                new_thread = Thread(target=self.automation.take_scrolled_screenshot,
                                    args=(self, prefix, suffix, second_instruction, delay))
                new_thread.start()
            else:
                self.display_message_box("Pls check prefix or Run name entered properly")
        except Exception as e:
            print(f"Error: {e}")

    @staticmethod
    def get_instructions():
        instructions = []
        with open("instructions.txt") as f:
            instructions = [instruction.strip() for instruction in f.readlines()]
        return instructions

    def set_instruction(self):
        self.prefixes_combobox.clear()
        if self.instructions:
            if 0 <= self.instructuns_index < len(self.instructions):
                self.instruction_box.setPlainText(self.instructions[self.instructuns_index])
                instruction = self.instructions[self.instructuns_index]
                if instruction in self.config["instructions"]:
                    self.prefixes_combobox.addItems(self.config["instructions"][instruction][0])
                return True
        else:
            self.display_message_box("Instructions are not given")

    def display_message_box(self, message):
        alert_box = QMessageBox(self)
        alert_box.setIcon(QMessageBox.Warning)
        alert_box.setText(message)
        alert_box.setStandardButtons(QMessageBox.Ok)
        alert_box.exec()

    @staticmethod
    def get_urls():
        urls_list = []
        with open("urls.txt") as f:
            urls_list = [url.strip() for url in f.readlines()]
        return urls_list

    def close_browser(self):
        self.status_bar.showMessage(f"Closing application", 2000)
        self.automation.close_browser()

    def record_screen(self):
        prefix = self.prefixes_combobox.currentText()
        suffix = self.run_name.text()
        if prefix and suffix:
            try:
                second_instruction = self.config["instructions"][self.instructions[self.instructuns_index]][1]
            except:
                second_instruction = ""
            delay = self.screenshot_delay.value()
            self.status_bar.showMessage(f"Started video recording", 2000)
            new_thread = Thread(target=self.automation.record_screen,
                                args=(self, prefix, suffix, second_instruction, delay))
            new_thread.start()
            # new_thread.join()
            # self.setEnabled(True)
            self.pause_record_button.setEnabled(True)
            self.stop_record_button.setEnabled(True)
            self.start_record_button.setDisabled(True)
        else:
            self.display_message_box("Pls check prefix or Run name entered properly")

    def stop_record(self):
        time.sleep(2)
        self.automation.record_screen_loop = False
        self.start_record_button.setEnabled(True)
        self.stop_record_button.setDisabled(True)
        self.pause_record_button.setDisabled(True)
        self.showNormal()

    def pause_record(self):
        self.automation.pause_record = not self.automation.pause_record
        if self.automation.pause_record:
            self.pause_record_button.setText("Resume")
            self.status_bar.showMessage(f"video recorder Paused", 3000)
        else:
            self.pause_record_button.setText("Pause")
            self.status_bar.showMessage(f"video recorder Resumed", 3000)

    def add_action(self):
        # if self.record_actions_check_box.isChecked():
        #     elements = self.automation.get_element_at_mouse(self)
        #     if isinstance(elements, list):
        #         element = elements[0]
        #     else:
        #         element = elements
        #     if element == "null" or not element:
        #         self.display_message_box("Not selected any event. after click Ctrl+A move mouse cursor bit")
        #     else:
        #         self.automation.scrollable_actions_list.append([element, self.add_click_check_box.isChecked()])
        #         self.display_message_box(f"Added event {element}")
        # else:
        position = pyautogui.position()
        self.non_scrollable_actions_list.append([position, self.add_click_check_box.isChecked()])
        self.display_message_box(f"Added action")
        self.add_click_check_box.setChecked(False)

    def play_actions(self):
        prefix = self.prefixes_combobox.currentText()
        suffix = self.run_name.text()
        if self.non_scrollable_actions_list:
            if prefix and suffix:
                self.showMinimized()

                def play_actions():
                    if self.record_actions_check_box.isChecked():
                        self.start_record_button.click()
                    delay = self.play_action_delay.value()
                    for position, click_or_not in self.non_scrollable_actions_list:
                        pyautogui.moveTo(position[0], position[1], delay)
                        if click_or_not:
                            pyautogui.click()
                    time.sleep(1)
                    self.stop_record()

                thread = Thread(target=play_actions)
                thread.start()
            else:
                self.display_message_box("Pls check prefix or Run name entered properly")
        else:
            self.display_message_box("Action list is empty")

    def clear_actions(self):
        self.automation.scrollable_actions_list = []
        self.non_scrollable_actions_list = []

    def user_detail_name_changed(self, value):
        name = self.details_name_combobox.currentText()
        self.details_list_combobox.clear()
        if name in self.user_details:
            self.details_list_combobox.addItems(self.user_details[name])

    def click_on_element(self):
        self.automation.click_on_element()

    def enter_user_details(self):
        self.automation.enter_user_details(self.details_list_combobox.currentText())

    def toggle_window_on_top(self):
        flags = Qt.WindowFlags()
        if self.keep_on_top_check_box.isChecked():
            flags |= QtCore.Qt.WindowStaysOnTopHint
        self.setWindowFlags(flags)
        self.show()

    def fill_form_automatically(self):
        prefix = self.prefixes_combobox.currentText()
        suffix = self.run_name.text()
        if prefix and suffix:
            self.showMinimized()
            if self.record_actions_check_box.isChecked():
                if self.start_record_button.isEnabled():
                    self.start_record_button.click()
            selected_form = self.config["user_form"][self.auto_fill_form_combobox.currentText()]
            thread = Thread(target=self.automation.fill_form_automatically,
                            args=[self, selected_form, self.stop_record_button])
            thread.start()
        else:
            self.display_message_box("Pls check prefix or Run name entered properly")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = BotWindow()
    app.aboutToQuit.connect(window.close_browser)
    app.exec()
