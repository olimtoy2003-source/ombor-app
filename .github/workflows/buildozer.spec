[app]
title = Ombor
package.name = ombor
package.domain = org.test
source.dir = .
source.include_exts = py,png,jpg,dotenv,txt,csv,db,xlsx
version = 0.1
requirements = python3,kivy==2.3.0,sqlite3
orientation = portrait
fullscreen = 0
android.permissions = READ_EXTERNAL_STORAGE, WRITE_EXTERNAL_STORAGE
android.api = 33
android.minapi = 21
android.accept_sdk_license = True
android.archs = arm64-v8a

[buildozer]
log_level = 2
warn_on_root = 1
