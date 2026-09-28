[app]

# (str) Title of your application
title = Telekom Ombor

# (str) Package name
package.name = telekomombor

# (str) Package domain (needed for android/ios packaging)
package.domain = org.telekom.ombor

# (str) Source code where the main.py live
source.dir = .

# (list) Source files to include (include your Excel file as well)
source.include_exts = py,png,jpg,kv,atlas,xlsx

# (list) Application requirements
# IMPORTANT: Include openpyxl for Excel handling
requirements = python3,kivy==2.3.0,pandas,openpyxl,requests

# (str) Application versioning
version = 1.0

# (list) Permissions
permissions = READ_EXTERNAL_STORAGE, WRITE_EXTERNAL_STORAGE, INTERNET

# (int) Target Android API
android.api = 33

# (int) Minimum API required
android.minapi = 21

# (str) Android NDK architecture
android.archs = arm64-v8a, armeabi-v7a

# (bool) Accept SDK license automatically
android.accept_sdk_license = True

# (str) Orientation
orientation = portrait

[buildozer]

# (int) Log level (0 = error only, 1 = info, 2 = debug)
log_level = 2

# (int) Display warning if buildozer is run as root (0 = error, 1 = warning)
warn_on_root = 1
