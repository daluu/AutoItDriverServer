from __future__ import print_function
from selenium import webdriver
from selenium.webdriver import ActionChains
import time

driver = webdriver.Remote( command_executor='http://127.0.0.1:4723/wd/hub', desired_capabilities={'browserName':'AutoIt'})
print("Desired Capabilities returned by server:\n")
print(driver.desired_capabilities)
print("")

# demo adapted from 
# http://www.joecolantonio.com/2014/07/02/selenium-autoit-how-to-automate-non-browser-based-functionality/
driver.get("calc.exe")
driver.switch_to_window("Calculator")
time.sleep(1)
driver.find_element_by_id("133").click() # 3
time.sleep(1)
driver.find_element_by_id("93").click() # +
time.sleep(1)
driver.find_element_by_id("133").click() # 3
time.sleep(1)
driver.find_element_by_id("121").click() # =
time.sleep(1)
if driver.find_element_by_id("150").text != "6":
	print("3 + 3 did not produce 6 as expected.")
driver.find_element_by_id("81").click() # Clear "C" button
time.sleep(1)

# demo adapted from AutoItX VBScript example that comes with AutoIt installation
action1 = ActionChains(driver)
action2 = ActionChains(driver)
action3 = ActionChains(driver)
action1.send_keys("2*2=").perform()
time.sleep(1)
if driver.find_element_by_id("150").text != "4":
	print("2 x 2 did not produce 4 as expected.")
driver.find_element_by_id("81").click() # Clear "C" button
time.sleep(1)
action2.send_keys("4*4=").perform()
time.sleep(1)
if driver.find_element_by_id("150").text != "16":
	print("4 x 4 did not produce 16 as expected.")
driver.find_element_by_id("81").click() # Clear "C" button
time.sleep(1)
action3.send_keys("8*8=").perform()
time.sleep(1)
if driver.find_element_by_id("150").text != "64":
	print("8 x 8 did not produce 64 as expected.")
driver.find_element_by_id("81").click() # Clear "C" button
time.sleep(1)
driver.close()
driver.quit()