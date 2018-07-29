import json
import os
import socket
import sys

import requests
import servicemanager
import usb.util
import win32event
import win32service
import win32serviceutil
import time


class Device(object):
    vendor = ""
    product = ""
    name = ""
    status = ""

    # The class "constructor" - It's actually an initializer
    def __init__(self, v, p, n):
        self.vendor = v
        self.product = p
        self.name = n
        self.status = 'plugged'


class USBMonitoring(win32serviceutil.ServiceFramework):
    _svc_name_ = "USBMonitoring"
    _svc_display_name_ = "USB Monitoring"
    _run = True
    with open('config.json', 'r') as f:
        _config = json.load(f)

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)
        self._run = True

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
        self._run = False

    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        self._run = True
        self.main()

    def main(self):
        mouse_dongle = Device(0x046D, 0xC531, 'dongle')
        usb_key = Device(0x125F, 0xDB8A, 'key')
        keyboard_test = Device(0x413C, 0x2003, 'keyboard')

        device_list = {usb_key, mouse_dongle, keyboard_test}
        slack_hook = self._config['MAIN']['SLACK_HOOK_URL']
        pc_identifier = os.environ['COMPUTERNAME']

        while self._run:
            for vendor in device_list:
                # find our device
                dev = usb.core.find(idVendor=vendor.vendor, idProduct=vendor.product)

                # was it found?
                if dev is None and vendor.status is 'plugged':
                    vendor.status = 'unplugged'
                    payload = {'channel': '#usb-monitoring', 'username': pc_identifier,
                               'text': 'Peripheral: ' + vendor.name + '\n Status: *' + vendor.status + '*'}
                    r = requests.post(slack_hook, data=json.dumps(payload))

                if dev is not None and vendor.status is 'unplugged':
                    vendor.status = 'plugged'
                    payload = {'channel': '#usb-monitoring', 'username': pc_identifier,
                               'text': 'Peripheral: ' + vendor.name + '\n Status: *' + vendor.status + '*'}
                    r = requests.post(slack_hook, data=json.dumps(payload))
            time.sleep(1)


if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(USBMonitoring)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(USBMonitoring)
