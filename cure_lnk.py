import ctypes, sys, win32api, os, win32com.client, shutil, win32con
# Check, if script runs from Administrator account
def is_admin():
	try:
		return ctypes.windll.shell32.IsUserAnAdmin()
	except:
		return False

if not is_admin():
	print("Please run this program as administrator.")
	input("Press Enter to continue...")
	sys.exit()
print("Running scan for .lnk on all drives...")
count = 0
infected = 0
restored = 0
# Get list of disks
drives = win32api.GetLogicalDriveStrings()
drives = drives.split('\000')[:-1]
# Loop disks
for drive in drives:
	print("Scanning drive", drive)
	# Loop all files in drive
	for dirname, dirs, files in os.walk(drive):
		for filename in files:
			filename_without_extension, extension = os.path.splitext(filename)
			# If file is link
			if extension == ".lnk":
				count +=1
				shell = win32com.client.Dispatch("WScript.Shell")
				fname = dirname+'\\'+filename
				shortcut = shell.CreateShortCut(fname)
				# Check link fingerprint
				if "/C attrib -h -s msg.exe & start /b msg.exe & \\MSOCache\\" in shortcut.Arguments:
					infected +=1
					hidden = drive+'MSOCache\\'+shortcut.Arguments.split('\\MSOCache\\')[1]
					originalname = dirname+'\\'+filename_without_extension
					try:
						# Try to restore hidden file
						shutil.move(hidden,originalname)
						# Clear hidden flag
						win32api.SetFileAttributes(originalname, win32con.FILE_ATTRIBUTE_NORMAL)
						# Remove link
						os.remove(fname)
						restored +=1
						print("Restored file",originalname)
					except Exception as e:
						# Eemove link anyway
						print("Removed broken link",fname)
						os.remove(fname)
print('Total',count,'files scanned',infected,'infected',restored,'restored.')
input("Press Enter to continue...")