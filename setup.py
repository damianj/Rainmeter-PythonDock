from distutils.core import setup
import py2exe

setup(console=['getLinkTargets.py'],
      windows=['launcher.py'],
      zipfile=None)
