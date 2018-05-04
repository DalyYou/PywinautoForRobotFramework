'''
# @Time     : 4/24/2018 3:26 PM
# @Author   : Daly.You
# @Email    : daly.you@wesoft.com
# @File     : WindowsUpdatesHelper.py
# @Project  : WindowsUpdatesLib
# @Python   : 2.7.14
# @Software : PyCharm Community Edition
'''
#!/usr/bin/python
# -*- coding:utf-8 -*-

import pywinauto
from pywinauto import Desktop, Application, keyboard
import threading
import _winreg
import re

class WindowsUpdatesHelper():
    def __init__(self):
        self.WUPU = WindowsUpdatesUtilities()
        self.Platform = self.WUPU.get_win_info()["ProductName"]
    def install_updates(self):
        is_win10 = re.search("Windows 10",self.Platform,re.IGNORECASE)
        is_win2016 = re.search("Windows Server 2016",self.Platform,re.IGNORECASE)
        print bool(is_win10)
        print bool(is_win2016)
        self.WUPU.right_click_start_button()
        if bool(is_win10) or bool(is_win2016):
            self.WUPU.click_context_menu("Settings")
            threading._sleep(3)
            self.WUPU.install_updates_via_settingspage()
            self.WUPU.close_windowsupdatesettingspage()
        else:
            self.WUPU.click_context_menu("Control Panel")
            threading._sleep(3)
            self.WUPU.install_updates_via_controlpanel()
            self.WUPU.close_windowsupdatecontrolpanel()

class WindowsUpdatesUtilities():
    def __init__(self):
        pass
    def get_win_info(self):
        win_info = {}
        reg_key = _winreg.OpenKey(_winreg.HKEY_LOCAL_MACHINE, "SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion")
        ProductName = _winreg.QueryValueEx(reg_key, "ProductName")[0] or None
        EditionId = _winreg.QueryValueEx(reg_key, "EditionId")[0] or None
        CurrentBuild = _winreg.QueryValueEx(reg_key, "CurrentBuild")[0] or None
        win_info["ProductName"] = ProductName.decode("utf-8")
        win_info["EditionId"] = EditionId.decode("utf-8")
        win_info["CurrentBuild"] = CurrentBuild.decode("utf-8")
        return win_info
    def is_win10_and_above(self, platform):
        is_win10 = re.search("Windows 10",platform,re.IGNORECASE)
        is_win2016 = re.search("Windows Server 2016",platform,re.IGNORECASE)
        print bool(is_win10)
        print bool(is_win2016)
        if bool(is_win10) or bool(is_win2016):
            return True
        else:
            return False
    def right_click_start_button(self):
        task_bar_app = Application().Connect(title=u'', class_name='Shell_TrayWnd')
        task_bar_class = task_bar_app.Shell_TrayWnd
        start_button = task_bar_class.Start
        start_button.set_focus()
        start_button.click_input()
        start_button.right_click_input()
        right_click_start_result = False
        while(right_click_start_result == False):
            try:
                right_click_startbutton_context = Desktop(backend='uia').window(title="Context", control_type="Menu").wait(wait_for='visible',timeout=5)
                if right_click_startbutton_context:
                    right_click_start_result = True
            except pywinauto.timings.TimeoutError, err:
                print err.message
    def click_context_menu(self, select_menu):
        right_click_startbutton_context = Desktop(backend='uia').window(title="Context", control_type="Menu")
        right_click_startbutton_context.wait('visible')
        #print context.print_control_identifiers()
        control_panel_menu = right_click_startbutton_context.window(title=select_menu, control_type="MenuItem")#Control Panel
        if control_panel_menu.exists():
            control_panel_menu.set_focus()
            control_panel_menu.click_input()

    def install_updates_via_controlpanel(self):
        #Pass on Win2012R2 Std
        available_button_exist = False
        threading._sleep(3)
        control_panel_app = Application().Connect(title_re=u'.*Control Panel.*', class_name='CabinetWClass', visible_only=True, enabled_only=True)
        control_panel_class = control_panel_app.CabinetWClass
        control_panel_class.wait('ready')
        address_band = control_panel_class[u'Address Band Root']
        address_band.ClickInput()
        address_band.type_keys("Control Panel\All Control Panel Items\Windows Update{ENTER}", with_spaces=True)
        threading._sleep(2)
        windows_update_app = Application().Connect(title=u'Windows Update', class_name='CabinetWClass', visible_only=True, enabled_only=True, active_only=True)
        windows_update_class = windows_update_app.window(class_name=u'CabinetWClass', visible_only=True, enabled_only=True, active_only=True)
        windows_update_class.wait('visible')
        try:

            install_updates_button = windows_update_class.window(title="&Install updates", class_name="Button", visible_only=True, enabled_only=True).wait("visible",timeout=5)
            print type(install_updates_button)
            if install_updates_button is None:
                print "No install updates button"
            else:
                available_button_exist = True
                install_updates_button.click()
        except pywinauto.timings.TimeoutError, err:
            print err.message
        try:
            restart_button = windows_update_class.window(title="&Restart now", class_name="Button", visible_only=True, enabled_only=True).wait("visible", timeout=5)
            if restart_button is None:
                print "No restart button"
            else:
                available_button_exist = True
                restart_button.click()
                threading._sleep(3)
                restart_warning_dialog = Application(backend='uia').connect(title=u'Windows',class_name='#32770')
                warning_dialog = restart_warning_dialog.Dialog
                #print warning_dialog.print_control_identifiers()
                if warning_dialog.exists():
                    warning_dialog.Yes.set_focus()
        except pywinauto.timings.TimeoutError, err:
            print err.message
        if available_button_exist == False:
            print "No Windows updates available"

    def close_windowsupdatecontrolpanel(self):
        control_panel_app = Application().connect(title=u'Windows Update', class_name='CabinetWClass', visible_only=True, enabled_only=True, active_only=True)
        control_panel_class = control_panel_app.CabinetWClass
        control_panel_class.wait('ready')
        control_panel_class.type_keys('%{F4}')

    def install_updates_via_settingspage(self):
        #Pass on Win10 Pro
        available_button_exist = False
        #settings main page
        settings_main_app = Application(backend='uia').connect(title=u'Settings', class_name='ApplicationFrameWindow', visible_only=True, enabled_only=True)
        settings_class = settings_main_app.Settings
        settings_class.wait('ready')
        #Dialog
        settings_dialog = settings_main_app.Dialog
        #ListBox
        settings_listBox = settings_dialog.ListBox
        #Click Update list Item
        update_listItem = settings_listBox.child_window(title="Update & Security", control_type="ListItem")
        update_listItem.wait('ready')
        update_listItem.click_input()
        threading._sleep(3)
        settings_class.wait('ready')
        #Setting Pane
        settings_pane = settings_dialog.Pane
        #print settings_pane.print_control_identifiers()
        #Check Check for updates button and click it
        try:
            check_updates_button = settings_pane.child_window(title="Check for updates", control_type="Button", visible_only=True, enabled_only=True).wait("visible",timeout=5)
            #print check_updates_button.exists()
            if check_updates_button is None:
                print "No install updates button"
            else:
                print "exist"
                available_button_exist = True
                check_updates_button.click_input()
        except pywinauto.timings.TimeoutError, err:
            print err.message
        except pywinauto.findbestmatch.MatchError, err:
            print err.message
        #Check Install now button exist and click it
        try:
            install_updates_button = settings_pane.child_window(title="Updates are ready to install", control_type="Button", visible_only=True, enabled_only=True).wait("visible",timeout=5)
            if install_updates_button is None:
                print "No install updates button"
            else:
                print "exist"
                available_button_exist = True
                install_updates_button.click_input()
        except pywinauto.timings.TimeoutError, err:
            print err.message
        except pywinauto.findbestmatch.MatchError, err:
            print err.message
        #Check Restart now button exist and click it
        try:
            restart_button = settings_pane.child_window(title="Restart now", control_type="Button", visible_only=True, enabled_only=True).wait("visible",timeout=5)
            if restart_button is None:
                print "No install updates button"
            else:
                print "exist"
                available_button_exist = True
                restart_button.click_input()
        except pywinauto.timings.TimeoutError, err:
            print err.message
        except pywinauto.findbestmatch.MatchError, err:
            print err.message
        if available_button_exist == False:
            print "No Windows updates available"

    def close_windowsupdatesettingspage(self):
        settings_main_app = Application(backend='uia').connect(title=u'Settings', class_name='ApplicationFrameWindow', visible_only=True, enabled_only=True)
        settings_class = settings_main_app.Settings
        settings_class.wait('ready')
        settings_class.type_keys('%{F4}')
