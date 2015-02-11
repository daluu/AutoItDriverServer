from __future__ import print_function
from selenium import webdriver
import os

driver = webdriver.Remote( command_executor='http://127.0.0.1:4723/wd/hub', desired_capabilities={'browserName':'AutoIt'})
print("Desired Capabilities returned by server:\n")
print(driver.desired_capabilities)
print("")

# execute an AutoIt script file (rather than call specific AutoItX API commands)
# supply path to AutoIt script file followed by optional arguments
driver.execute_script("C:\\PathOnAutoItDriverServerHostMachineTo\\demo.au3","Hello","World")

# or if using compiled binary option 
# (but need to set AutoItScriptExecuteScriptAsCompiledBinary to True in autoit_options.cfg first
# before starting the server).
#driver.execute_script("C:\\PathOnAutoItDriverServerHostMachineTo\\demo.exe","Hello","World")

driver.quit()