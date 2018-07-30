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
    def __init__(self, vendor, product, name):
        self.vendor = vendor
        self.product = product
        self.name = name
        self.status = 'plugged'


class USBMonitoring(win32serviceutil.ServiceFramework):
    _svc_name_ = "USBMonitoring"
    _svc_display_name_ = "USB Monitoring"
    _run = True
    with open('C:\\Windows\\usbmonitor.json', 'r') as f:
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
        config_device_list = self._config['DEVICES']
        device_list = []
        slack_hook = self._config['MAIN']['SLACK_HOOK_URL']
        pc_identifier = os.environ['COMPUTERNAME']

        for config_device in config_device_list:
            device = Device(
                    vendor=int(config_device['VENDOR_ID'], 16),
                    product=int(config_device['PRODUCT_ID'], 16),
                    name=config_device['DISPLAY_NAME'],
                )
            device_list.append(device)

        while self._run:
            for device in device_list:
                # find our device
                usb_device = usb.core.find(idVendor=device.vendor, idProduct=device.product)

                # was it found?
                if usb_device is None and device.status is 'plugged':
                    device.status = 'unplugged'
                    payload = {'channel': '#usb-monitoring', 'username': pc_identifier,
                               'text': 'Peripheral: ' + device.name + '\nStatus: *' + device.status + '*'}
                    r = requests.post(slack_hook, data=json.dumps(payload))

                if usb_device is not None and device.status is 'unplugged':
                    device.status = 'plugged'
                    payload = {'channel': '#usb-monitoring', 'username': pc_identifier,
                               'text': 'Peripheral: ' + device.name + '\nStatus: *' + device.status + '*'}
                    r = requests.post(slack_hook, data=json.dumps(payload))
            time.sleep(1)


if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(USBMonitoring)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(USBMonitoring)
