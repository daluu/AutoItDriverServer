from __future__ import print_function
from selenium import webdriver
from selenium.webdriver import ActionChains
import time

wd = webdriver.Firefox()
ad = webdriver.Remote( command_executor='http://127.0.0.1:4723/wd/hub', desired_capabilities={'browserName':'AutoIt'})
print("Desired Capabilities returned by server:\n")
print(ad.desired_capabilities)
print("")

action1 = ActionChains(ad)

### HTTP authentication dialog popup demo ###
wd.get("http://www.httpwatch.com/httpgallery/authentication/")
time.sleep(1)
# check state that img is "unauthenticated" at start
img_src = wd.find_element_by_id("downloadImg").get_attribute("src")
if not img_src.endswith("/images/spacer.gif"):
	print("HTTP demo test fail because test site not started with correct default unauthenticated state.")

# now test authentication
wd.find_element_by_id("displayImage").click() # trigger the popup
time.sleep(5) # wait for popup to appear
ad.switch_to_window("Authentication Required")
action1.send_keys("httpwatch{TAB}AutoItDriverServerAndSeleniumIntegrationDemo{TAB}{ENTER}").perform()
time.sleep(5)

# now check img is authenticated or changed
img_src = wd.find_element_by_id("downloadImg").get_attribute("src")
if img_src.endswith("/images/spacer.gif"):
	print("HTTP demo failed, image didn't authenticate/change after logging in.")

### file upload demo, also adapted from sample code of the test/target site ###
wd.get("http://www.toolsqa.com/automation-practice-form")
# wd.find_element_by_id("photo").click() # this doesn't seem to trigger file upload to popup
elem = wd.find_element_by_id("photo")
wd.execute_script("arguments[0].click();",elem)
time.sleep(10) # wait for file upload dialog to appear
ad.switch_to_window("File Upload")

# FYI, on FF, Opera, (Windows Safari) filename field control ID may be 1152
# but on IE, Chrome, and Windows in general, should be control ID should be 1148

ad.find_element_by_class_name("Edit1").send_keys("C:\\ReplaceWith\\PathTo\\AnActualFile.txt")
# due to remote file upload (via local file detector), you may notice actual path typed
# may differ from what you fill in above, but that path is still valid for the original file
# it's just a copy in a temp directory. This is the nature of file uploads over RemoteWebDriver
# instead of a local driver.
time.sleep(2)
ad.find_element_by_class_name("Button1").click()
#actions.send_keys("{ENTER}").perform() # another option to "click" Open/Upload
#actions.send_keys("!o").perform() # another option to invoke Open/Upload via keyboard shortcut ALT + O
time.sleep(5)

### handle browser's print dialog popup demo? to come... ###
# Print dialog popup may hang Selenium code,
# so may want to fire off a separate thread for AutoIt(DriverServer)
# to monitor for popup when it shows up (i.e. Selenium then hung)
# and have AutoIt print or cancel the popup to then resume Selenium main thread code

# file download popup handling demo? to come...

# handle Adobe Acrobat PDF loaded inside browser demo? e.g. click save, print, etc. to come...

# handle Flash or embedded streaming video controls demo? e.g. play, stop, etc. to come...

wd.quit()
ad.quit()