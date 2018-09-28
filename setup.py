from distutils.core import setup
import py2exe

setup(console=[{'script':'fuel_transform.py'}],
      options={"py2exe":{"includes":["xlrd", "xlsxwriter"]}})
