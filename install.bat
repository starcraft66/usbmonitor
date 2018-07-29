copy /Y config.json C:\\Windows\\usbmonitor.json
sc stop "USBMonitoring"
copy /Y .\\dist\\usb_monitor.exe C:\\Windows\\usb_monitor.exe
C:\\Windows\\usb_monitor.exe --startup auto install
C:\\Windows\\usb_monitor.exe start
pause
