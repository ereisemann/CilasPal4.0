#import sys
#sys.path.append('/Users/eveeisemann/Documents/GitHub/CilasPal4.0/CilasPal')   ## If script can't find PdfReaderObj map your directory here
import PdfReaderObj as p
import pypdf
import xlsxwriter
import openpyxl
import numpy as np
import math
import os
from colorama import Fore

# Global vars
defined_classes_headers = ['0.04', '3.90', '62.00', '88.00', '125.0', '177.0', '250.0', '350.0', '500.0', '710.0', '1000.0',
                                        '1410.0', '2000.0']
standard_classes_headers = ['0.04', '0.07', '0.10', '0.20', '0.30', '0.40', '0.50', '0.60', '0.70', '0.80', '0.90', '1.0', '1.1',
                                        '1.2', '1.3', '1.4', '1.6', '1.8', '2.0', '2.2', '2.4', '2.6', '3.0', '4.0', '5.0', '6.0', '6.5', '7.0',
                                        '7.5', '8.0', '8.5', '9.0', '10.0', '11.0', '12.0', '13.0', '14.0', '15.0', '16.0', '17.0', '18.0',
                                        '19.0', '20.0', '22.0', '25.0', '28.0', '32.0', '36.0', '38.0', '40.0', '45.0', '50.0', '53.0', '56.0',
                                        '63.0', '71.0', '75.0', '80.0', '85.0', '90.0', '95.0', '100.0', '106.0', '112.0', '125.0', '130.0',
                                        '140.0', '145.0', '150.0', '160.0', '170.0', '180.0', '190.0', '200.0', '212.0', '242.0', '250.0',
                                        '300.0', '400.0', '500.0', '600.0', '700.0', '800.0', '900.0', '1000.0', '1100.0', '1200.0', '1300.0',
                                        '1400.0', '1500.0', '1600.0', '1700.0', '1800.0', '1900.0', '2000.0', '2100.0', '2200.0', '2300.0',
                                        '2400.0', '2500.0']

file = None

debug = input("Run in debug mode? (Y/N)")
if debug == "Y" or debug == "y":
    from CilasPalSetup import build_env
    from CilasPalDebugger import debug
    file = "tests/testfile.pdf"
    print(Fore.YELLOW + "...debugging with verbose output...")
    print(Fore.YELLOW + "...rebuilding CilasPal env...")
    build_env()
    print(Fore.YELLOW + "...running test file...")
    debug(file, defined_classes_headers, standard_classes_headers)
else:
    print("Welcome, please paste the path to the data file (.pdf only!)")
    file = input("Full file path is: ")

    pdf = p.PdfReaderObject(file)      ## PDF defined here based on PDF reader object (see PdfReaderObj.py)

    spreadsheet_path = pdf.init_excel()
    print(spreadsheet_path)

    row = 2
    # Now parse info from pdf. I'll keep this as straight forward as possible
    for i in range(0, pdf.num_pages, 2):
        contents1 = pdf.read_content(i)   ## Contents of page 1 of sample i (user defined size classes)
        contents2 = pdf.read_content(i+1) ## Contents of page 2 of sample i (all size classes)
        sample_name = contents1[contents1.index("Sample ref.") + 14:contents1.index("Sample Name")]
        median = contents1[contents1.index('Diameter at 50%') + 18:contents1.index('µmDiameter at 90%')-1]
        mean = contents1[contents1.index('Mean diameter') + 16:contents1.index('FraunhoferDensity')-4]

        # Check index alignment on first sample
        if i == 0:
            print(Fore.YELLOW + "...Testing Index Alignment on sample 1...\n" + Fore.WHITE)
            print(f"name:{sample_name}mean: {mean}\nmedian: {median}")

            if " " in sample_name or " " in mean or " " in median: # If there are spaces before/after
                raise Exception("Misalignment in indexing size metrics. Check that pdf version == or that the sample name does not contain spaces. Or, if you're sure this error is a mistake, delete the conditional on line 45 in CilasPal.py")
            else:
                print(Fore.GREEN + "Correct indexing size metrics ... checking size classes..." + Fore.WHITE)

                contents1 = contents1[620:]    

                defined_class_distrib = []
                for classs in defined_classes_headers:
                    s = contents1.index(classs)
                    try:
                        val = float(contents1[s+6:s+11])     # Catches if output contains characters --> can not convert a misalignment to float
                        defined_class_distrib.append(val)
                    except:
                        print(Fore.RED + f"Misalignment in indexing size classes. Can not cast {contents1[s+6:s+11]} to float" + Fore.WHITE)
                        raise Exception("Misalignment in indexing size classes. Can not cast to float. see above")
                print(Fore.GREEN + "Correct indexing for customer defined size classes ... checking standard classes..." + Fore.WHITE)
                contents2 = contents2[620:]    

                standard_class_distrib = []

                for classs in standard_classes_headers:
                    s = contents2.index(classs)

                    ## TESTING \/\/\/
                    #print(str(s) + ' = index of class heading')
                    #print(classs + ' = class heading')
                    #val = float(contents2[s + 12:s + 16])  ## individual size class value (q3)
                    #print(str(val) + ' = val')
                    #print(str(contents2[s+12:s+16]) + ' = val')
                    ## TESTING /\/\/\

                    try:
                        #val = float(contents2[s+6:s+11])   ## cumulative value (Q3)
                        val = float(contents2[s+12:s+16])   ## individual size class value (q3)
                        print(val)
                        standard_class_distrib.append(val)
                    except:
                        print(Fore.RED + f"Misalignment in indexing size classes. Can not cast {contents2[s+6:s+11]} to float" + Fore.WHITE)
                        raise Exception("Misalignment in indexing size classes. Can not cast to float, see above")
                print(Fore.GREEN + "Correct indexing for standard size classes ... proceeding to parse full data..." + Fore.WHITE)
                # print(standard_class_distrib)


        # Same exact thing as above, still catches misalignment errors but has no verbose output
        contents1 = pdf.read_content(i)
        contents2 = pdf.read_content(i+1)        
        if " " in sample_name or " " in mean or " " in median:
            raise Exception("Misalignment in indexing size metrics. Check that pdf version ==  or that the sample name does not contain spaces. Or, if you're sure this error is a mistake, delete lines 77-79 in CilasPal.py")
        else:
            contents1 = contents1[620:]    
            defined_class_distrib = []
            for classs in defined_classes_headers:
                s = contents1.index(classs)
                try:
                    val = float(contents1[s+6:s+11])
                    defined_class_distrib.append(val)
                except:
                    print(Fore.RED + f"Misalignment in indexing size classes. Can not cast {contents1[s+6:s+11]} to float" + Fore.WHITE)
                    raise Exception("Misalignment in indexing size classes. Can not cast to float. see above")
            contents2 = contents2[620:]    

            standard_class_distrib = []
            for classs in standard_classes_headers:
                s = contents2.index(classs)
                try:
                    #val = float(contents2[s+6:s+11])
                    val = float(contents2[s+12:s+16])   ## See comment above about indexing
                    standard_class_distrib.append(val)
                except:
                    print(Fore.RED + f"Misalignment in indexing size classes. Can not cast {contents2[s+12:s+16]} to float" + Fore.WHITE)
                    raise Exception("Misalignment in indexing size classes. Can not cast to float, see above")

                    
            workbook = openpyxl.load_workbook(spreadsheet_path)
            defined_class_distrib_sheet = workbook[workbook.sheetnames[1]]
            standard_class_distrib_sheet = workbook[workbook.sheetnames[0]]


        standard_class_distrib_sheet.cell(row=row, column=1, value=sample_name)
        standard_class_distrib_sheet.cell(row=row, column=2, value=mean)
        standard_class_distrib_sheet.cell(row=row, column=3, value=median)
        for val, j in zip(standard_class_distrib, range(0, len(standard_class_distrib))):
            standard_class_distrib_sheet.cell(row=row, column=j+4, value=val)
        workbook.save(spreadsheet_path)
        defined_class_distrib_sheet.cell(row=row, column=1, value=sample_name)
        for val, j in zip(defined_class_distrib, range(0, len(defined_class_distrib))):
            defined_class_distrib_sheet.cell(row=row, column=j+2, value=val)
        workbook.save(spreadsheet_path)
        row+=1
            

