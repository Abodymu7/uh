import PyInstaller.__main__
import sys

# تحديد اسم ملف الكود الأساسي الخاص بك (تأكد أن اسمه com.py)
script_name = 'com.py'

PyInstaller.__main__.run([
    script_name,
    '--onefile',
    '--noconsole',
    '--name=GifToPptx'
])