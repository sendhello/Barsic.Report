pip install virtualenv
virtualenv venv
venv\Scripts\activate
venv\Scripts\pip.exe install -r req.txt
venv\Scripts\pip.exe install --upgrade pip wheel setuptools
venv\Scripts\pip.exe install --upgrade google-api-python-client
venv\Scripts\python.exe INSTALL\KivyMD\setup.py install