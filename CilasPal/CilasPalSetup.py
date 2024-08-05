from PackageManager import install_packages

def build_env():
    install_packages()
    
    from colorama import Fore
    print(Fore.WHITE + "...Ensuring Correct Build for CilasPal...")
    
    import pypdf
    import openpyxl
    import xlsxwriter
    pdfversion = pypdf.__version__ == '4.3.1'
    openpyxlversion = openpyxl.__version__ == '3.1.5'
    xlsxversion = xlsxwriter.__version__ == '3.2.0'
    if pdfversion and openpyxlversion and xlsxversion:
        print(Fore.GREEN + "pypdf = 4.3.1 \nopenpyxl = 3.1.5 \nxlsxwriter = 3.2.0 \n" + Fore.WHITE)
        print(Fore.GREEN + "build is correct!" + Fore.WHITE)
        
    else:
        raise Exception("Incorrect build. Ensure that pypdf version 4.3.1, openpyxl version 3.1.5, xlsxwriter version 3.2.0")
    
build_env()