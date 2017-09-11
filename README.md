# APOD Wallpaper Changer for Windows 7, 10
Python script for APOD (NASA Astronomy Picture Of the Day) Wallpaper Changer for Windows 7, 10

Instructions for use with Windows 7 or Windows 10:

Pre-requisites:
---------------
1. Install Python 3.x. [https://www.python.org/downloads/]
2. Open command prompt (Windows Key + R, Open: cmd) and run the below commands
 	 - pip install pypiwin32
 	 - pip install image

Instructions for running the script:
------------------------------------
1. If you just want to run the script once or want to run the script manually every time:
   - $python apod.py
2. If you want the script register itself with Windows Task Scheduler to run automatically
   - $python apod.py install
3. If you want the script to unregister itself with Windows Task Scheduler
   - $python apod.py uninstall
