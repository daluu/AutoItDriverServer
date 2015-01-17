AutoItDriverServer
=========

AutoItDriverServer is a server interface wrapper to AutoIt that provides a Selenium WebDriver API via the webdriver JSON  wire protocol to drive AutoIt (using AutoItX COM/DLL API). 

There are 3 benefits to testing with AutoItDriverServer:

1: you are able to write your test in your choice of programming language, using the Selenium WebDriver API and language-specific client libraries. 

2: is remote execution enabled, and perhaps Selenium Grid compatible in the future. See [Why use AutoItDriverServer](https://github.com/daluu/AutoItDriverServer/wiki/Why-use-AutoItDriverServer) for more details.

3: uses the well known AutoIt tool for Windows GUI manipulation. If you already use AutoIt, this will be a nice benefit.

Quick Start & server/implementation notes
------------------------------------------

The server is currently implemented in Python, calling AutoItX over COM (or optionally/alternatively DLL). There is a plan to do a .NET/C# version that is more standalone than Python in the future. Both implementations have the goal of working with all off the shelf Selenium client libraries.

Python implementation is adapted from the old Appium server Python implementation (https://github.com/hugs/appium-old), and uses the [Bottle micro web-framework](http://www.bottlepy.org). The .NET/C# implementation will be adapted from Strontium server (https://github.com/jimevans/strontium).

To get started, clone the repo:<br />
`git clone git://github.com/daluu/AutoItDriverServer`

Next, change into the 'autoitdriverserver_python' directory, and install dependencies:<br />
`pip install -r partial_requirements.txt`

The partial_requirements file doesn't necessarily specify all requirements. You'll need to have the Python win32com.client module (or Python for Windows extensions, win32 extensions for Python, etc.). That may already be installed with your Python installation. Or you may have to separately install that yourself. You can test first yourself to see if executing "import win32com.client" will return an exception or not in Python, with no errors meaning it's already installed. Alternatively, the server code was initially written to also work with https://github.com/jacexh/pyautoit as well (if you swap out the commented code with the current code).

Additionally, you'll need to have AutoIt installed, or register the appropriate version of AutoItX DLL (x86 vs x64). When installing AutoIt or registering DLL, the version to register/use depends on the platform you're using with. For Python it would depend on whether you run the 32 or 64 bit version of Python not whether your OS is 32 or 64 bit.

To launch a webdriver-compatible server, run:<br />
`python server.py` <br />

For additional parameter info, append the `--help` parameter

Example WebDriver API/client test usage against this server tool can be found in `sample-code` folder. To run the test, startup the server (with customized parameters as needed, recommend set address to 127.0.0.1) and review the example files' code before executing those scripts. Examples use Python WebDriver binding, but any language binding will actually do.

NOTES/Caveats
-------------

AutoItDriverServer is simply a WebDriver server interface to AutoIt. Issues you experience may usually be the result of an issue with AutoIt rather than AutoItDriverServer itself. It would be wise to test your issue/scenario using AutoIt (via native AutoIt script compiled to executable or not) or AutoItX to confirm whether the issue is with AutoIt or AutoItDriverServer. The source code of AutoItDriverServer will point you to the appropriate AutoItX API or see the wiki documentation here.

WebDriver API/command support and mapping to AutoItX API
-------------------------------------------------------

See [WebDriver API command support and mapping to AutoItX API](https://github.com/daluu/AutoItDriverServer/wiki/WebDriver-API-command-support-and-mapping-to-AutoItX-API)

Contributing
------------

Fork the project, make a change, and send a pull request!

Or as a user, try it out, provide your feedback, submit bugs/enhancement requests.
