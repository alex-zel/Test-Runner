from distutils.core import setup
import py2exe, sys, os, shutil

sys.argv.append('py2exe')

if os.path.exists(r'./runner'):
    shutil.rmtree(r'./runner')

py2exe_options = dict(
                      bundle_files=2,
                      compressed=True,
                      dist_dir='./runner',
                      excludes=['Tkconstants', 'Tkinter', 'tcl'],
                      )

setup(console=[{'script': 'main.py'}],
      zipfile=None,
      options={'py2exe': py2exe_options}, requires=['openpyxl']
      )

shutil.copy(r'.\data.json', 'runner')
