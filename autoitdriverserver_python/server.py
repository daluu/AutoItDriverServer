#    Source code is modified from and based off of 
#    old/original Appium Python implementation at
#
#    https://github.com/hugs/appium-old
#
#    Licensed to the Apache Software Foundation (ASF) under one
#    or more contributor license agreements.  See the NOTICE file
#    distributed with this work for additional information
#    regarding copyright ownership.  The ASF licenses this file
#    to you under the Apache License, Version 2.0 (the
#    "License"); you may not use this file except in compliance
#    with the License.  You may obtain a copy of the License at
#
#    http://www.apache.org/licenses/LICENSE-2.0
#
#    Unless required by applicable law or agreed to in writing,
#    software distributed under the License is distributed on an
#    "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
#    KIND, either express or implied.  See the License for the
#    specific language governing permissions and limitations
#    under the License.

from __future__ import print_function
from bottle import Bottle, request, response, redirect
from bottle import run, static_file
import json
import socket
import sys
import platform
import os
import subprocess
import tempfile
import base64
import urllib
#import autoit
import win32com.client
from time import time
from time import sleep
try:
    import configparser
except ImportError:
    import ConfigParser as configparser

app = Bottle()

@app.get('/favicon.ico')
def get_favicon():
    return static_file('favicon.ico', root='.')

def get_platform():
    if sys.platform == "win32":
        if platform.release() == "Vista":
            wd_platform = "VISTA"
        elif platform.release() == "XP": #?
            wd_platform = "XP"
        else:
            wd_platform = "WINDOWS"
    else:
        wd_platform = "WINDOWS"
    return wd_platform

@app.route('/wd/hub/status', method='GET')
def status():
    status = {'sessionId': app.SESSION_ID if app.started else None,
              'status': 0,
              'value': {'build': {'version': 'AutoItDriverServer 0.1'}, 
              'os': {'arch':platform.machine(),'name':'windows','version':platform.release()}}}
    return status

@app.route('/wd/hub/session', method='POST')
def create_session():
    #app.oAutoItX = win32com.client.Dispatch("AutoItX3.Control")

    #process desired capabilities
    request_data = request.body.read()
    dc = json.loads(request_data.decode()).get('desiredCapabilities')
    if dc is not None:
        app.caretCoordMode = dc.get('caretCoordMode') if dc.get('caretCoordMode') is not None else app.caretCoordMode
        app.expandEnvStrings = dc.get('expandEnvStrings') if dc.get('expandEnvStrings') is not None else app.expandEnvStrings
        app.mouseClickDelay = dc.get('mouseClickDelay') if dc.get('mouseClickDelay') is not None else app.mouseClickDelay
        app.mouseClickDownDelay = dc.get('mouseClickDownDelay') if dc.get('mouseClickDownDelay') is not None else app.mouseClickDownDelay
        app.mouseClickDragDelay = dc.get('mouseClickDragDelay') if dc.get('mouseClickDragDelay') is not None else app.mouseClickDragDelay
        app.mouseCoordinateMode = dc.get('mouseCoordinateMode') if dc.get('mouseCoordinateMode') is not None else app.mouseCoordinateMode
        app.sendAttachMode = dc.get('sendAttachMode') if dc.get('sendAttachMode') is not None else app.sendAttachMode
        app.sendCapslockMode = dc.get('sendCapslockMode') if dc.get('sendCapslockMode') is not None else app.sendCapslockMode
        app.sendKeyDelay = dc.get('sendKeyDelay') if dc.get('sendKeyDelay') is not None else app.sendKeyDelay
        app.sendKeyDownDelay = dc.get('sendKeyDownDelay') if dc.get('sendKeyDownDelay') is not None else app.sendKeyDownDelay
        app.winDetectHiddenText = dc.get('winDetectHiddenText') if dc.get('winDetectHiddenText') is not None else app.winDetectHiddenText
        app.winSearchChildren = dc.get('winSearchChildren') if dc.get('winSearchChildren') is not None else app.winSearchChildren
        app.winTextMatchMode = dc.get('winTextMatchMode') if dc.get('winTextMatchMode') is not None else app.winTextMatchMode
        app.winTitleMatchMode = dc.get('winTitleMatchMode') if dc.get('winTitleMatchMode') is not None else app.winTitleMatchMode
        app.winWaitDelay = dc.get('winWaitDelay') if dc.get('winWaitDelay') is not None else app.winWaitDelay
        
        #autoit.option.caret_coord_mode = app.caretCoordMode
        app.oAutoItX.Opt("CaretCoordMode", app.caretCoordMode)
        #autoit.option.expand_env_strings = app.expandEnvStrings
        app.oAutoItX.Opt("ExpandEnvStrings", app.expandEnvStrings)
        #autoit.option.mouse_click_delay = app.mouseClickDelay
        app.oAutoItX.Opt("MouseClickDelay", app.mouseClickDelay)
        #autoit.option.mouse_click_down_delay = app.mouseClickDownDelay
        app.oAutoItX.Opt("MouseClickDownDelay", app.mouseClickDownDelay)
        #autoit.option.mouse_click_drag_delay = app.mouseClickDragDelay
        app.oAutoItX.Opt("MouseClickDragDelay", app.mouseClickDragDelay)
        #autoit.option.mouse_coordinate_mode = app.mouseCoordinateMode
        app.oAutoItX.Opt("MouseCoordinateMode", app.mouseCoordinateMode)
        #autoit.option.send_attach_mode = app.sendAttachMode
        app.oAutoItX.Opt("SendAttachMode", app.sendAttachMode)
        #autoit.option.send_capslock_mode = app.sendCapslockMode
        app.oAutoItX.Opt("SendCapslockMode", app.sendCapslockMode)
        #autoit.option.send_key_delay = app.sendKeyDelay
        app.oAutoItX.Opt("SendKeyDelay", app.sendKeyDelay)
        #autoit.option.send_key_down_delay = app.sendKeyDownDelay
        app.oAutoItX.Opt("SendKeyDownDelay", app.sendKeyDownDelay)
        #autoit.option.win_detect_hidden_text = app.winDetectHiddenText
        app.oAutoItX.Opt("WinDetectHiddenText", app.winDetectHiddenText)
        #autoit.option.win_search_children = app.winSearchChildren
        app.oAutoItX.Opt("WinSearchChildren", app.winSearchChildren)
        #autoit.option.win_text_match_mode = app.winTextMatchMode
        app.oAutoItX.Opt("WinTextMatchMode", app.winTextMatchMode)
        #autoit.option.win_title_match_mode = app.winTitleMatchMode
        app.oAutoItX.Opt("WinTitleMatchMode", app.winTitleMatchMode)
        #autoit.option.win_wait_delay = app.winWaitDelay
        app.oAutoItX.Opt("WinWaitDelay", app.winWaitDelay)

    #setup session
    app.started = True
    redirect('/wd/hub/session/%s' % app.SESSION_ID)

@app.route('/wd/hub/session/<session_id>', method='GET')
def get_session(session_id=''):
    wd_platform = get_platform()
    app_response = {'sessionId': session_id,
                'status': 0,
                'value': {"version":"0.1",
                          "browserName":"AutoIt",
                          "platform":wd_platform,
                          "takesScreenshot":False,                          
                          "caretCoordMode":app.caretCoordMode,
                          "expandEnvStrings":app.expandEnvStrings,
                          "mouseClickDelay":app.mouseClickDelay,
                          "mouseClickDownDelay":app.mouseClickDownDelay,
                          "mouseClickDragDelay":app.mouseClickDragDelay,
                          "mouseCoordinateMode":app.mouseCoordinateMode,
                          "sendAttachMode":app.sendAttachMode,
                          "sendCapslockMode":app.sendCapslockMode,
                          "sendKeyDelay":app.sendKeyDelay,
                          "sendKeyDownDelay":app.sendKeyDownDelay,
                          "winDetectHiddenText":app.winDetectHiddenText,
                          "winSearchChildren":app.winSearchChildren,
                          "winTextMatchMode":app.winTextMatchMode,
                          "winTitleMatchMode":app.winTitleMatchMode,
                          "winWaitDelay":app.winWaitDelay}}
    return app_response

@app.route('/wd/hub/session/<session_id>', method='DELETE')
def delete_session(session_id=''):
    app.started = False
    # any need to dispose of/clean up COM connection to AutoIt for win32com?
    # if we instantiated from create session (rather than at server startup)
    #app.oAutoItX = None # for example
    app_response = {'sessionId': session_id,
                'status': 0,
                'value': {}}
    return app_response

@app.route('/wd/hub/session/<session_id>/execute', method='POST')
def execute_script(session_id=''):
    request_data = request.body.read()
    try:
        script = json.loads(request_data.decode()).get('script')
        args = json.loads(request_data.decode()).get('args')

        if config.get("AutoIt Options",'AutoItScriptExecuteScriptAsCompiledBinary') == "False":
            if platform.machine() == "AMD64":
                if config.get("AutoIt Options",'AutoIt64BitOSOnInstallUse32Bit') == "True":
                    au3Runner = config.get("AutoIt Options",'AutoIt64BitOS32BitExecutablePath')
                else:
                    au3Runner = config.get("AutoIt Options",'AutoIt64BitOS64BitExecutablePath')
            else: # platform.machine() == "i386"
                au3Runner = config.get("AutoIt Options",'AutoIt32BitExecutablePath')
            script_call = "%s %s" % (au3Runner,script)
        else:
            script_call = script
        if args is not None:
            for arg in args:
                script_call = "%s %s" % (script_call,arg)
        print("script2exec: ",script_call)
        os.system(script_call)
    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 13, 'value': str(sys.exc_info()[1])}

    app_response = {'sessionId': session_id,
        'status': 0,
        'value': {}}
    return app_response

@app.route('/wd/hub/session/<session_id>/element/<element_id>/click', method='POST')
def element_click(session_id='', element_id=''):
    try:
        element = decode_value_from_wire(element_id)
        #result = autoit.control_click("[active]", element)
        result = app.oAutoItX.ControlClick("[active]", "", element)
        if result == 0:
            raise Exception("AutoIt failed to click element %s." % element)
    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 13, 'value': str(sys.exc_info()[1])}

    app_response = {'sessionId': session_id,
        'status': 0,
        'value': {}}
    return app_response

@app.route('/wd/hub/session/<session_id>/click', method='POST')
def mouse_click(session_id=''):
    request_data = request.body.read()
    if request_data == None or request_data == '' or request_data == "{}":
        button = 0
    else:
        button = json.loads(request_data.decode()).get('button')
    try:
        if button == 1:
            btn_type = "middle"
        elif button == 2:
            btn_type = "right"
        else:
            btn_type = "left"
        #autoit.mouse_click(btn_type)
        app.oAutoItX.MouseClick(btn_type)
    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 13, 'value': str(sys.exc_info()[0])}

    app_response = {'sessionId': session_id,
        'status': 0,
        'value': {}}
    return app_response

@app.route('/wd/hub/session/<session_id>/doubleclick', method='POST')
def double_click(session_id=''):
    try:
        #src = autoit.mouse_get_pos()
        #autoit.mouse_click("left",src.x,src.y,2)
        x = app.oAutoItX.MouseGetPosX()
        y = app.oAutoItX.MouseGetPosY()
        app.oAutoItX.MouseClick("left",x,y,2)
    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 13, 'value': str(sys.exc_info()[0])}

    app_response = {'sessionId': session_id,
        'status': 0,
        'value': {}}
    return app_response

@app.route('/wd/hub/session/<session_id>/buttonup', method='POST')
def mouse_up(session_id=''):
    request_data = request.body.read()
    if request_data == None or request_data == '' or request_data == "{}":
        button = 0
    else:
        button = json.loads(request_data.decode()).get('button')
    try:
        if button == 1:
            btn_type = "middle"
        elif button == 2:
            btn_type = "right"
        else:
            btn_type = "left"
        #autoit.mouse_up(btn_type)
        app.oAutoItX.MouseUp(btn_type)
    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 13, 'value': str(sys.exc_info()[0])}

    app_response = {'sessionId': session_id,
        'status': 0,
        'value': {}}
    return app_response

@app.route('/wd/hub/session/<session_id>/buttondown', method='POST')
def mouse_down(session_id=''):
    request_data = request.body.read()
    if request_data == None or request_data == '' or request_data == "{}":
        button = 0
    else:
        button = json.loads(request_data.decode()).get('button')
    try:
        if button == 1:
            btn_type = "middle"
        elif button == 2:
            btn_type = "right"
        else:
            btn_type = "left"
        #autoit.mouse_down(btn_type)
        app.oAutoItX.MouseDown(btn_type)
    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 13, 'value': str(sys.exc_info()[0])}

    app_response = {'sessionId': session_id,
        'status': 0,
        'value': {}}
    return app_response

@app.route('/wd/hub/session/<session_id>/moveto', method='POST')
def move_to(session_id=''):
    request_data = request.body.read()
    if request_data == None or request_data == '' or request_data == "{}":
        element_id = None
        xoffset = None
        yoffset = None
    else:
        element_id = json.loads(request_data.decode()).get('element')
        xoffset = json.loads(request_data.decode()).get('xoffset')
        yoffset = json.loads(request_data.decode()).get('yoffset')
    try:
        if element_id == None and (xoffset != None or yoffset != None):
            #src = autoit.mouse_get_pos()            
            #autoit.mouse_move(src.x+xoffset,src.y+yoffset)
            x = app.oAutoItX.MouseGetPosX()
            y = app.oAutoItX.MouseGetPosY()
            app.oAutoItX.MouseMove(x+xoffset,y+yoffset)
        else:
            if xoffset != None or yoffset != None:
                control_id = decode_value_from_wire(element_id)
                #pos = autoit.control_get_pos("[active]",control_id)                
                x = app.oAutoItX.ControlGetPosX("[active]","",control_id)
                y = app.oAutoItX.ControlGetPosY("[active]","",control_id)
                #if autoit._has_error():
                if app.oAutoItX.error == 1:
                    raise Exception("AutoIt failed to get element %s location to move to." % control_id)
                #autoit.mouse_move(pos.left+xoffset,pos.top+yoffset)
                app.oAutoItX.MouseMove(x+xoffset,y+yoffset)
            else: # just go to center of element
                control_id = decode_value_from_wire(element_id)
                #pos = autoit.control_get_pos("[active]",control_id)                
                x = app.oAutoItX.ControlGetPosX("[active]","",control_id)
                y = app.oAutoItX.ControlGetPosY("[active]","",control_id)
                width = app.oAutoItX.ControlGetPosWidth("[active]","",control_id)
                height = app.oAutoItX.ControlGetPosHeight("[active]","",control_id)
                #if autoit._has_error():
                if app.oAutoItX.error == 1:
                    raise Exception("AutoIt failed to get element %s location to move to." % control_id)
                #autoit.mouse_move(pos.left+(pos.right/2),pos.top+(pos.bottom/2))
                app.oAutoItX.MouseMove(x+(width/2),y+(height/2))
    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 13, 'value': str(sys.exc_info()[1])}

    app_response = {'sessionId': session_id,
        'status': 0,
        'value': {}}
    return app_response

@app.route('/wd/hub/session/<session_id>/element/<element_id>/value', method='POST')
def set_value(session_id='', element_id=''):
    request_data = request.body.read()
    try:
        value_to_set = json.loads(request_data.decode()).get('value')
        value_to_set = ''.join(value_to_set)
        control_id = decode_value_from_wire(element_id)
        #result = autoit.control_set_text("[active]",control_id,value_to_set)
        result = app.oAutoItX.ControlSetText("[active]","",control_id,value_to_set)
        if result == 0:
            raise Exception("AutoIt failed to set text of element %s, element or window not found." % control_id)
    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 13, 'value': str(sys.exc_info()[1])}

    app_response = {'sessionId': session_id,
        'status': 0,
        'value': {}}
    return app_response

@app.route('/wd/hub/session/<session_id>/element/<element_id>/clear', method='POST')
def clear(session_id='', element_id=''):
    try:
        control_id = decode_value_from_wire(element_id)
        #result = autoit.control_set_text("[active]",control_id,"")
        result = app.oAutoItX.ControlSetText("[active]","",control_id,"")
        if result == 0:
            raise Exception("AutoIt failed to clear text of element %s, element or window not found." % control_id)
    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 13, 'value': str(sys.exc_info()[1])}

    app_response = {'sessionId': session_id,
        'status': 0,
        'value': {}}
    return app_response

@app.route('/wd/hub/session/<session_id>/element/<element_id>/element', method='POST')
def element_find_element(session_id='', element_id=''):
    return _find_element(session_id, element_id)

@app.route('/wd/hub/session/<session_id>/element', method='POST')
def find_element(session_id=''):
    return _find_element(session_id, "root")

def _find_element(session_id, context, many=False):
    try:
        json_request_data = json.loads(request.body.read().decode())
        locator_strategy = json_request_data.get('using')
        value = json_request_data.get('value')

        if locator_strategy == "id":
            control_id = "[ID:%s]" % value
        elif locator_strategy == "link text":
            control_id = "[TEXT:%s]" % value
        elif locator_strategy == "tag name":
            control_id = "[CLASS:%s]" % value
        elif locator_strategy == "class name":
            control_id = "[CLASSNN:%s]" % value
        elif locator_strategy == "name":
            control_id = "[NAME:%s]" % value
        elif locator_strategy == "xpath":
            control_id = "[REGEXPCLASS:%s]" % value
        elif locator_strategy == "css selector":
            control_id = value
        else:
            control_id = value

        # AutoIt has no concept of finding elements/controls and checking if it
        # exists or not. Therefore, we just pass back the location strategy value
        # and let the individual WebElement methods fail when control/element not found
        found_elements = {'ELEMENT':encode_value_4_wire(control_id)}            
        return {'sessionId': session_id, 'status': 0, 'value': found_elements}
    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 13, 'value': str(sys.exc_info()[0])}

@app.route('/wd/hub/session/<session_id>/keys', method='POST')
def keys(session_id=''):
    try:
        request_data = request.body.read()
        wired_keys = json.loads(request_data.decode()).get('value')
        keys = "".join(wired_keys)
        #autoit.send(keys)
        app.oAutoItX.Send(keys)
        return {'sessionId': session_id, 'status': 0}
    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 13, 'value': str(sys.exc_info()[0])}

@app.route('/wd/hub/session/<session_id>/element/<element_id>/location', method='GET')
def element_location(session_id='', element_id=''):
    try:
        control_id = decode_value_from_wire(element_id)
        #pos = autoit.control_get_pos("[active]",control_id)
        #location = {'x': pos.left, 'y': pos.top}
        x = app.oAutoItX.ControlGetPosX("[active]","",control_id)
        y = app.oAutoItX.ControlGetPosY("[active]","",control_id)
        location = {'x': x, 'y': y}
        #if autoit._has_error():
        if app.oAutoItX.error == 1:
            raise Exception("AutoIt failed to get element %s location coordinates." % control_id)
        return {'sessionId': session_id, 'status': 0, 'value': location}
    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 13, 'value': str(sys.exc_info()[1])}
    
@app.route('/wd/hub/session/<session_id>/element/<element_id>/size', method='GET')
def element_size(session_id='', element_id=''):
    try:
        control_id = decode_value_from_wire(element_id)
        #pos = autoit.control_get_pos("[active]",control_id)
        #size = {'width': pos.right, 'height': pos.bottom}
        width = app.oAutoItX.ControlGetPosWidth("[active]","",control_id)
        height = app.oAutoItX.ControlGetPosWidth("[active]","",control_id)
        size = {'width': width, 'height': height}
        #if autoit._has_error():
        if app.oAutoItX.error == 1:
            raise Exception("AutoIt failed to get element %s width & height size." % control_id)
        return {'sessionId': session_id, 'status': 0, 'value': size}
    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 13, 'value': str(sys.exc_info()[1])}

@app.route('/wd/hub/session/<session_id>/element/<element_id>/displayed', method='GET')
def element_displayed(session_id='', element_id=''):
    try:
        control_id = decode_value_from_wire(element_id)
        #result = autoit.control_command("[active]",control_id,"IsVisible")
        result = app.oAutoItX.ControlCommand("[active]","",control_id,"IsVisible")
        displayed = True if result == 1 else False
        #if autoit._has_error():
        if app.oAutoItX.error == 1:
            raise Exception("AutoIt failed to find element %s." % control_id)
        return {'sessionId': session_id, 'status': 0, 'value': displayed}
    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 13, 'value': str(sys.exc_info()[1])}

@app.route('/wd/hub/session/<session_id>/element/<element_id>/enabled', method='GET')
def element_enabled(session_id='', element_id=''):
    try:
        control_id = decode_value_from_wire(element_id)
        #result = autoit.control_command("[active]",control_id,"IsEnabled")
        result = app.oAutoItX.ControlCommand("[active]","",control_id,"IsEnabled")
        enabled = True if result == 1 else False
        #if autoit._has_error():
        if app.oAutoItX.error == 1:
            raise Exception("AutoIt failed to find element %s." % control_id)
        return {'sessionId': session_id, 'status': 0, 'value': enabled}
    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 13, 'value': str(sys.exc_info()[1])}

@app.route('/wd/hub/session/<session_id>/element/<element_id>/selected', method='GET')
def element_selected(session_id='', element_id=''):
    try:
        control_id = decode_value_from_wire(element_id)
        #result = autoit.control_command("[active]",control_id,"IsChecked")
        result = app.oAutoItX.ControlCommand("[active]","",control_id,"IsChecked")
        selected = True if result == 1 else False
        #if autoit._has_error():
        if app.oAutoItX.error == 1:
            raise Exception("AutoIt failed to find element %s." % control_id)
        return {'sessionId': session_id, 'status': 0, 'value': selected}
    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 13, 'value': str(sys.exc_info()[1])}

@app.route('/wd/hub/session/<session_id>/element/<element_id>/text', method='GET')
def get_text(session_id='', element_id=''):
    try:
        control_id = decode_value_from_wire(element_id)
        #text = autoit.control_get_text("[active]",control_id)
        text = app.oAutoItX.ControlGetText("[active]","",control_id)
        #if autoit._has_error() and text == "":
        if app.oAutoItX.error == 1 and text == "":
            raise Exception("AutoIt failed to find element %s, and get it's text" % control_id)
    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 13, 'value': str(sys.exc_info()[1])}

    app_response = {'sessionId': session_id,
        'status': 0,
        'value': text}
    return app_response

@app.route('/wd/hub/session/<session_id>/element/<element_id>/attribute/<attribute>', method='GET')
def get_attribute(session_id='', element_id='', attribute=''):
    try:
        control_id = decode_value_from_wire(element_id)
        #result = autoit.control_command("[active]",control_id,attribute)
        result = app.oAutoItX.ControlCommand("[active]","",control_id,attribute)
        #if autoit._has_error():
        if app.oAutoItX.error == 1:
            raise Exception("AutoIt failed to find element %s or extract the specified attribute/command '%s'." % (control_id,attribute))
    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 13, 'value': str(sys.exc_info()[1])}

    app_response = {'sessionId': session_id,
        'status': 0,
        'value': str(result)}
    return app_response

@app.route('/wd/hub/session/<session_id>/element/<element_id>/css/<property_name>', method='GET')
def get_property(session_id='', element_id='', property_name=''):
    try:
        control_id = decode_value_from_wire(element_id)
        #result = autoit.control_command("[active]",control_id,property_name)
        result = app.oAutoItX.ControlCommand("[active]","",control_id,property_name)
        #if autoit._has_error():
        if app.oAutoItX.error == 1:
            raise Exception("AutoIt failed to find element %s or extract the specified property/command '%s'." % (control_id,property_name))
    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 13, 'value': str(sys.exc_info()[1])}

    app_response = {'sessionId': session_id,
        'status': 0,
        'value': str(result)}
    return app_response

@app.route('/wd/hub/session/<session_id>/title', method='GET')
def get_window_title(session_id=''):
    try:
        #text = autoit.win_get_title("[active]")
        text = app.oAutoItX.WinGetTitle("[active]")
        if text == 0:
           raise Exception("AutoIt failed to get window title of current window.") 
    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 23, 'value': str(sys.exc_info()[1])}

    app_response = {'sessionId': session_id,
        'status': 0,
        'value': text}
    return app_response

@app.route('/wd/hub/session/<session_id>/window_handle', method='GET')
def get_current_window_handle(session_id=''):
    try:
        #handle = autoit.win_get_handle("[active]")
        handle = app.oAutoItX.WinGetHandle("[active]")
        #if autoit._has_error() and handle == "":
        if app.oAutoItX.error == 1 and handle == "":
            raise Exception("AutoIt failed to get current window handle.")
    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 23, 'value': str(sys.exc_info()[1])}

    app_response = {'sessionId': session_id,
        'status': 0,
        'value': handle}
    return app_response

@app.route('/wd/hub/session/<session_id>/window', method='POST')
def select_window(session_id=''):
    request_data = request.body.read()
    try:
        win_name_or_handle = json.loads(request_data.decode()).get('name')
        #try:
            #autoit.win_activate_by_handle(win_name_or_handle)
        #except:
            #autoit.win_activate(win_name_or_handle)
        app.oAutoItX.WinActivate(win_name_or_handle)
    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 23, 'value': str(sys.exc_info()[0])}

    app_response = {'sessionId': session_id,
        'status': 0,
        'value': {}}
    return app_response

@app.route('/wd/hub/session/<session_id>/window', method='DELETE')
def close_window(session_id=''):
    try:
        #autoit.win_close("[active]")
        app.oAutoItX.WinClose("[active]")
    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 23, 'value': str(sys.exc_info()[0])}

    app_response = {'sessionId': session_id,
        'status': 0,
        'value': {}}
    return app_response

@app.route('/wd/hub/session/<session_id>/window/<window_handle>/size', method='POST')
def resize_window(session_id='', window_handle=''):
    try:
        request_data = request.body.read()
        width = json.loads(request_data.decode()).get('width')
        height = json.loads(request_data.decode()).get('height')

        if window_handle == "current":
            window = "[active]"
            #pos = autoit.win_get_pos(window)
            #if autoit._has_error():
                #raise Exception("Window handle %s not found for resizing." % window_handle)
            #autoit.win_move(window,pos.left,pos.top,width,height)
        else:
            window = window_handle
            #pos = autoit.win_get_pos_by_handle(window)
            #if autoit._has_error():
                #raise Exception("Window handle %s not found for resizing." % window_handle)
            #autoit.win_move_by_handle(window,pos.left,pos.top,width,height)
            
        x = app.oAutoItX.WinGetPosX(window)
        y = app.oAutoItX.WinGetPosX(window)
        if app.oAutoItX.error == 1:
            raise Exception("Window handle %s not found for resizing." % window_handle)
        app.oAutoItX.WinMove(window,x,y,width,height)

    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 23, 'value': str(sys.exc_info()[1])}

    app_response = {'sessionId': session_id,
        'status': 0,
        'value': {}}
    return app_response

@app.route('/wd/hub/session/<session_id>/window/<window_handle>/size', method='GET')
def get_window_size(session_id='', window_handle=''):
    try:
        if window_handle == "current":
            window = "[active]"
            #pos = autoit.win_get_pos(window)
        else:
            window = window_handle
            #pos = autoit.win_get_pos_by_handle(window)

        #size = {'width': pos.right, 'height': pos.bottom}
        width = app.oAutoItX.WinGetPosWidth(window)
        height = app.oAutoItX.WinGetPosHeight(window)
        size = {'width': width, 'height': height}
        #if autoit._has_error():
        if app.oAutoItX.error == 1:
            raise Exception("Window handle %s not found for getting window size." % window_handle)        
    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 23, 'value': str(sys.exc_info()[1])}

    app_response = {'sessionId': session_id,
        'status': 0,
        'value': size}
    return app_response

@app.route('/wd/hub/session/<session_id>/window/<window_handle>/position', method='POST')
def move_window(session_id='', window_handle=''):
    try:
        request_data = request.body.read()
        x = json.loads(request_data.decode()).get('x')
        y = json.loads(request_data.decode()).get('y')

        if window_handle == "current":
            window = "[active]"
            #autoit.win_move(window,x,y)
        else:
            window = window_handle
            #autoit.win_move_by_handle(window,x,y)
        
        app.oAutoItX.WinMove(window,x,y)
    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 23, 'value': str(sys.exc_info()[0])}

    app_response = {'sessionId': session_id,
        'status': 0,
        'value': {}}
    return app_response

@app.route('/wd/hub/session/<session_id>/window/<window_handle>/position', method='GET')
def get_window_position(session_id='', window_handle=''):
    try:
        if window_handle == "current":
            window = "[active]"
            #pos = autoit.win_get_pos(window)
        else:
            window = window_handle
            #pos = autoit.win_get_pos_by_handle(window)
        
        #size = {'x': pos.left, 'y': pos.top}
        x = app.oAutoItX.WinGetPosX(window)
        y = app.oAutoItX.WinGetPosY(window)
        size = {'x': x, 'y': y}
        #if autoit._has_error():
        if app.oAutoItX.error == 1:
            raise Exception("Window handle %s not found for getting window position." % window_handle)        
    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 23, 'value': str(sys.exc_info()[1])}

    app_response = {'sessionId': session_id,
        'status': 0,
        'value': size}
    return app_response

@app.route('/wd/hub/session/<session_id>/window/<window_handle>/maximize', method='POST')
def max_window(session_id='', window_handle=''):
    try:
        if window_handle == "current":
            window = "[active]"
            #pos = autoit.win_set_state(window,autoit.properties.SW_MAXIMIZE)
        else:
            window = window_handle
            #pos = autoit.win_set_state_by_handle(window,autoit.properties.SW_MAXIMIZE)
            
        # FYI, @SW_MAXIMIZE = 3, per https://github.com/jacexh/pyautoit/blob/master/autoit/autoit.py#L111
        app.oAutoItX.WinSetState(window,app.oAutoItX.SW_MAXIMIZE)
    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 23, 'value': str(sys.exc_info()[0])}

    app_response = {'sessionId': session_id,
        'status': 0,
        'value': {}}
    return app_response

@app.route('/wd/hub/session/<session_id>/url', method='POST')
def run_autoit_app(session_id=''):
    try:
        request_data = request.body.read()
        app2run = json.loads(request_data.decode()).get('url')
        #autoit.run(app2run)
        app.oAutoItX.Run(app2run)
        #if autoit._has_error():
        if app.oAutoItX.error == 1:
            raise Exception("Failed to run the application/executable: %s." % app2run)
    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 13, 'value': str(sys.exc_info()[1])}

    app_response = {'sessionId': session_id,
        'status': 0,
        'value': {}}
    return app_response

@app.route('/wd/hub/session/<session_id>/file', method='POST')
def upload_file(session_id=''):
    try:
        request_data = request.body.read()
        b64data = json.loads(request_data.decode()).get('file')
        byteContent = base64.b64decode(b64data)
        path = ""
        with tempfile.NamedTemporaryFile(delete=False) as f:
            f.write(byteContent)
            path = f.name
        extracted_files = unzip(path,os.path.dirname(path))        
    except:
        response.status = 400
        return {'sessionId': session_id, 'status': 13, 'value': str(sys.exc_info()[0])}

    # For (remote) file uploads - well currently AutoItDriverServer will always be "remote"
    # we can't formally/technically support multiple file uploads yet, due to Selenium issue 2239
    # as the WebDriver/JSONWireProtocol spec doesn't define how to handle request/response
    # of multiple files uploaded. Therefore, we assume user upload single file for now
    result = "".join(extracted_files)
    app_response = {'sessionId': session_id,
        'status': 0,
        'value': result}
    return app_response

def unzip(source_filename, dest_dir):
    import zipfile,os.path
    files_in_zip = []
    with zipfile.ZipFile(source_filename) as zf:        
        for member in zf.infolist():
            words = member.filename.split('/')
            path = dest_dir
            for word in words[:-1]:
                drive, word = os.path.splitdrive(word)
                head, word = os.path.split(word)
                if word in (os.curdir, os.pardir, ''): continue
                path = os.path.join(path, word)
            zf.extract(member, path)
            unzipped_file = os.path.join(dest_dir,member.filename)
            print("Unzipped a file: ",unzipped_file)
            files_in_zip.append(unzipped_file)
    return files_in_zip

@app.error(404)
def unsupported_command(error):
    response.content_type = 'text/plain'
    return 'Unrecognized command, or AutoItDriverServer does not support/implement this: %s %s' % (request.method, request.path)

def encode_value_4_wire(value):
    try:
        return urllib.parse.quote(base64.b64encode(value.encode("utf-8")))
    except:
        return urllib.quote(base64.b64encode(value.encode("utf-8")))

def decode_value_from_wire(value):
    try:
        return base64.b64decode(urllib.parse.unquote(value)).decode("utf-8")
    except:
        return base64.b64decode(urllib.unquote(value)).decode("utf-8")

if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='AutoItDriverServer - a webdriver-compatible server for use with desktop GUI automation via AutoIt COM/DLL interface.')
    #parser.add_argument('-v', dest='verbose', action="store_true", default=False, help='verbose mode')
    parser.add_argument('-a', '--address', type=str, default=None, help='ip address to listen on')
    parser.add_argument('-p', '--port', type=int, default=4723, help='port to listen on')
    parser.add_argument('-c', '--autoit_options_file',  type=str, default=None, help='config file defining the AutoIt options to use, see default sample config file in the app/server directory')

    args = parser.parse_args()
    
    if args.address is None:
        try:
            args.address = socket.gethostbyname(socket.gethostname())
        except:
            args.address = '127.0.0.1'
    
    if args.autoit_options_file is not None:
        options_file = args.autoit_options_file
    else:
        options_file = os.path.join(os.path.curdir,'autoit_options.cfg')
    config = configparser.RawConfigParser()
    config.read(options_file)
    app.caretCoordMode = config.get("AutoIt Options",'CaretCoordMode')
    app.expandEnvStrings = config.get("AutoIt Options",'ExpandEnvStrings')
    app.mouseClickDelay = config.get("AutoIt Options",'MouseClickDelay')
    app.mouseClickDownDelay = config.get("AutoIt Options",'MouseClickDownDelay')
    app.mouseClickDragDelay = config.get("AutoIt Options",'MouseClickDragDelay')
    app.mouseCoordinateMode = config.get("AutoIt Options",'MouseCoordinateMode')
    app.sendAttachMode = config.get("AutoIt Options",'SendAttachMode')
    app.sendCapslockMode = config.get("AutoIt Options",'SendCapslockMode')
    app.sendKeyDelay = config.get("AutoIt Options",'SendKeyDelay')
    app.sendKeyDownDelay = config.get("AutoIt Options",'SendKeyDownDelay')
    app.winDetectHiddenText = config.get("AutoIt Options",'WinDetectHiddenText')
    app.winSearchChildren = config.get("AutoIt Options",'WinSearchChildren')
    app.winTextMatchMode = config.get("AutoIt Options",'WinTextMatchMode')
    app.winTitleMatchMode = config.get("AutoIt Options",'WinTitleMatchMode')
    app.winWaitDelay = config.get("AutoIt Options",'WinWaitDelay')

    # for win32com, if we choose to instantiate AutoIt COM connection here
    # instead of per session basis
    app.oAutoItX = win32com.client.Dispatch("AutoItX3.Control")

    app.SESSION_ID = "%s:%d" % (args.address, args.port)
    app.started = False
    run(app, host=args.address, port=args.port)