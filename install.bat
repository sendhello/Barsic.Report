pip install virtualenv
virtualenv venv
venv\Scripts\pip.exe install -r requirements.txt
venv\Scripts\pip.exe install --upgrade pip wheel setuptools
venv\Scripts\pip.exe install --upgrade google-api-python-client
cd INSTALL\KivyMD
setup.py install