# APOD (Astronomy Picture of the Day) Wallpaper Changer for Windows
# --------------------------------------------------------------------
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>
#-------------------------------------------------------------------------------------------------
#
# Instructions for use with Windows 7 or Windows 10:
# Pre-requisites:
# ---------------
# 1.) Install Python 3.x. 
# 2.) Open command prompt and run the below commands
# 		pip install pypiwin32
# 		pip install image
# Instructions for running the script:
# ------------------------------------
# 1.) If you just want to run the script once or want run the script manually every time:
# 	$python apod.py
# 2.) If you want the script register itself with Windows Task Scheduler to run automatically
#   $python apod.py install
# 3.) If you want the script to unregister itself with Windows Task Scheduler
#   $python apod.py uninstall
# 4.) If you want to force an image update, run
#	$python apod.py update

import urllib.request, json
from win32api import GetSystemMetrics
import win32api, win32con
from PIL import Image, ImageFont, ImageDraw
import textwrap, ctypes
import logging, datetime
import os, sys, subprocess
import shutil

class ApodSettings(object):
	# APOD Image Configurations
	apod_url = "https://api.nasa.gov/planetary/apod?api_key=DEMO_KEY"
	apod_url_hd = "https://api.nasa.gov/planetary/apod?api_key=DEMO_KEY&hd=True"
	hd_image = True
	image_path = os.path.dirname(os.path.realpath(__file__)) + "\\apod_wallpaper.jpg"
	processed_image_path = os.path.dirname(os.path.realpath(__file__)) + "\\apod_wallpaper1.jpg"
	json_data = None
	screen_width = GetSystemMetrics(0)
	screen_height = GetSystemMetrics(1)
	img_size = None
	maxsize = (screen_width, screen_height)
	recognized_formats = ("jpg", "jpeg", "bmp", "png")
	
	# Display settings related to Title of the APOD image
	title_start = (0, 0) #Place holder, value will be calculated at runtime
	title_size = 20
	title_font = ImageFont.truetype("arial.ttf", title_size)
	
	# Display settings related to Explanation of the APOD image
	explanation_embed = True
	explanation_start = (0, 0) #Place holder, value will be calculated at runtime
	explanation_size = 12
	explanation_font = ImageFont.truetype("arial.ttf", explanation_size)
	
	# General settings related to display of text
	text_width = 100
	text_border_offset = 25
	
	# Configuration details related to logging
	log_filename = os.path.dirname(os.path.realpath(__file__)) + "\\apod_downloader.log"
	log_last_success = os.path.dirname(os.path.realpath(__file__)) + "\\last_success.log"
	
	# Generic settings
	windows_task_name = "Apod Wallpaper Changer"
	
def hours_since_last_update(settings):
	try:
		f = open(settings.log_last_success, 'r')
	except IOError:
		return True
	
	timestamp = f.read()
	f.close()
	last_run = datetime.datetime.strptime(timestamp, "%Y-%m-%d %H:%M",)
	now = datetime.datetime.now()
	delta = now - last_run
	delta_in_hours = delta.total_seconds()/(60*60)
	
	if delta_in_hours < 0:
		delta_in_hours = 24
	
	return delta_in_hours
	
def update_last_success_timestamp(settings):
	try:
		f = open(settings.log_last_success, 'w')
	except IOError:
		logging.error("Failed to log the last successful run timestamp to %s" % settings.log_last_success)
		return
	timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
	f.write(timestamp)
	f.close()
	
def download_apod_image(settings):
	image_date = datetime.date.today()
	date_query_str = ""

	if settings.hd_image:
		apod_url = settings.apod_url_hd
		url_key = "hdurl"
	else:
		apod_url = settings.apod_url
		url_key = "url"

	apod_url_initial = apod_url
	while True:
		if date_query_str:
			apod_url = apod_url_initial + "&date=" + date_query_str
			
		try:
			with urllib.request.urlopen(apod_url) as url:
				settings.json_data = json.loads(url.read().decode())
				if settings.json_data["media_type"] != "image":
					logging.error("Unsupported media type : %s" % settings.json_data["media_type"])
					image_date = image_date - datetime.timedelta(days=30)
					date_query_str = image_date.strftime("%Y-%m-%d")
					logging.info("Picking image from older date : %s" % date_query_str)
					continue
				file_path = settings.json_data[url_key]		
			if not file_path.lower().endswith(settings.recognized_formats):
				logging.error("Unrecognized file type in %s" % settings.json_data["url"])
				image_date = image_date - timedelta(days=30)
				date_query_str = image_date.strftime("%Y-%m-%d")
				logging.info("Picking image from older date : %s" % date_query_str)
				continue
			urllib.request.urlretrieve(settings.json_data["url"],settings.image_path)
		except urllib.error.URLError  as e:
			logging.error('URLError: %s. Exiting' % (str(e)))
			return False
		except urllib.error.HTTPError as e:
			logging.error('HTTPError: %s. Exiting' % (str(e)))
			return False
		else:
			logging.info('APOD image downloaded. Processing.')
			break; # No need to loop anymore
	return True

def calculate_text_placement(settings, draw):
	text = textwrap.fill(settings.json_data["explanation"], settings.text_width)
	size = draw.multiline_textsize(text, settings.explanation_font)
	settings.text_width = int(round((settings.img_size[0] / size[0]) * 
								settings.text_width))
	while True:
		text = textwrap.fill(settings.json_data["explanation"], settings.text_width)
		size = draw.multiline_textsize(text, settings.explanation_font)
		
		if size[0] < (settings.img_size[0] - (2 * settings.text_border_offset)):
			break
		
		settings.text_width -= 10
		logging.info('New text_width = %d' % settings.text_width)
		continue
		
	title_size = settings.title_font.getsize(settings.json_data["title"])
	x = settings.text_border_offset
	y = settings.img_size[1] - (title_size[1] + size[1] + settings.text_border_offset)
	settings.title_start = (x, y)
	settings.explanation_start = (x, y + title_size[1])
	
def process_apod_image(settings):
	img = Image.open(settings.image_path)
	ratio = min((settings.maxsize[0]/img.size[0]), (settings.maxsize[1]/img.size[1]))
	req_size = (int(round(img.size[0] * ratio)), int(round(img.size[1] * ratio)))
	img = img.resize(req_size, Image.ANTIALIAS)
	settings.img_size = img.size
	
	# Convert to RGB mode
	if img.mode != "RGB":
	    img = img.convert("RGB")
	
	if settings.explanation_embed == True:
		draw = ImageDraw.Draw(img)
		calculate_text_placement(settings, draw)
		draw.text(settings.title_start, settings.json_data["title"],(255,255,255),font=settings.title_font)

		text = textwrap.fill(settings.json_data["explanation"], width=settings.text_width)
		draw.multiline_text(settings.explanation_start,text,(255,255,255), font=settings.explanation_font)

	img.save(settings.processed_image_path)

def set_apod_wallpaper(settings):
	image_path_cbuff = bytes(settings.processed_image_path,encoding='utf-8')
	key = win32api.RegOpenKeyEx(win32con.HKEY_CURRENT_USER,"Control Panel\\Desktop",0,win32con.KEY_SET_VALUE)
	win32api.RegSetValueEx(key, "WallpaperStyle", 0, win32con.REG_SZ, "0")
	win32api.RegSetValueEx(key, "TileWallpaper", 0, win32con.REG_SZ, "0")
	retval = ctypes.windll.user32.SystemParametersInfoA(win32con.SPI_SETDESKWALLPAPER,
			len(image_path_cbuff), image_path_cbuff,
			win32con.SPIF_SENDWININICHANGE | win32con.SPIF_UPDATEINIFILE)
	if retval == 0:
		logging.error("SystemParametersInfoA(SPI_SETDESKWALLPAPER) returned = %d", retval)
		retval = ctypes.GetLastError()
		logging.error("GetLastError= %d", retval)
		return False
	logging.info('APOD image processing complete. Set as wallpaper')
	return True

def manage_installation(settings, action):
	appdata_path = os.getenv('LOCALAPPDATA')
	folder_name = "ApodDownloader"
	file_name = __file__
	dest = appdata_path + "/" + folder_name
	this_script_path = os.path.realpath(__file__)
	python_path = sys.exec_prefix + "\\pythonw.exe"
	
	if action == 'INSTALL':
		try:
			os.makedirs(dest, exist_ok=True)
			shutil.copy(this_script_path, dest)
			create_cmd = ['schtasks', '/Create', '/SC', 'HOURLY', '/MO', '2','/TN',
				settings.windows_task_name, '/TR',
				python_path + " " + dest + "/" + file_name, "/F"]
			completed = subprocess.run(create_cmd)
		except OSError as e:
			logging.error('Installation failed : %s' % (str(e)))
			return False			
	elif action == 'UNINSTALL':
		try:
			shutil.rmtree(dest, ignore_errors=True)
		except OSError as err:
			logging.error('Uninstallation failed : %s' % (str(e)))
			return False
		else:
			delete_cmd = ['schtasks', '/Delete', '/TN',
				settings.windows_task_name, '/F']
			completed = subprocess.run(delete_cmd)
	else:
		return False

	return True

def main():
	settings = ApodSettings()
	done = False
	
	logging.basicConfig(filename=settings.log_filename,level=logging.DEBUG,format='%(asctime)s %(levelname)s %(message)s')
	logging.info('Starting script')
	
	if len(sys.argv) > 1:
		if sys.argv[1] == "install":
			manage_installation(settings, 'INSTALL')
			done = False		# Also download and set APOD wallpaper
		elif sys.argv[1] == "uninstall":
			logging.info('Uninstalling Windows scheduler task : %s' % settings.windows_task_name)
			manage_installation(settings, 'UNINSTALL')
			done = True
		elif sys.argv[1] == "update":
			done = False
	elif hours_since_last_update(settings) <= 24:
		logging.info('No update needed')
		done = True
		
	if not done:
		success = download_apod_image(settings)
		if success:
			logging.info('Successfully downloaded image of the day.')
			process_apod_image(settings)
			logging.info('Processing image complete.')
			set_apod_wallpaper(settings)
			update_last_success_timestamp(settings)		

	logging.info('Exiting script')
	logging.info('=================================')
	
if __name__ == "__main__":
	main()
