from __future__ import print_function
from selenium import webdriver
from selenium.webdriver import ActionChains
import time

driver = webdriver.Remote( command_executor='http://127.0.0.1:4723/wd/hub', desired_capabilities={'browserName':'AutoIt'})
print("Desired Capabilities returned by server:\n")
print(driver.desired_capabilities)
print("")

# demo adapted from AutoItX VBScript example that comes with AutoIt installation
driver.get("notepad.exe")
driver.switch_to_window("Untitled - Notepad")
time.sleep(1)
action1 = ActionChains(driver)
action2 = ActionChains(driver)
action3 = ActionChains(driver)
action4 = ActionChains(driver)
action1.send_keys("Hello, this is line 1{ENTER}").perform()
time.sleep(1)
action2.send_keys("This is line 2{ENTER}This is line 3").perform()
time.sleep(1)
action3.send_keys("!{F4}").perform()
time.sleep(1)
action4.send_keys("!n").perform()
driver.quit()