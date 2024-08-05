import subprocess
import sys

'''
IMPORTANT developer notes on PyPDF2 and pypdf:

PyPDF2 is NOT the same package as pypdf. They are developed by the same people but are not the same package. PyPDF2 version 3.0.1 is the final version of that package, thus For this project I install the most recent (as of 07-31-2024) version of pypdf using 'pip install pypdf==4.3.1' not 'pip install pypdf' 

Additionally I use xlsxwriter and openpyxl as openpyxl is a more robust module and is easier to use
'''
def install_packages():
    # Install pypdf
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pypdf==4.3.1'])
    
    # Install openpyxl
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'openpyxl==3.1.5'])
    
    # Install xlsxwriter
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'xlsxwriter==3.2.0'])

    packages = [
        'numpy',
        'colorama',
        'pathlib']
    
    # Install the required packages
    for package in packages:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
        