import sys
sys.path.append('/Users/eveeisemann/Documents/GitHub/CilasPal4.0/CilasPal')
import PdfReaderObj as p
import CilasPalSetup as setup
import CilasPalDebugger as debug
import pypdf
import xlsxwriter
import openpyxl
import numpy as np
import math
import os
from colorama import Fore

def debug(file, defined_classes_headers, standard_classes_headers):
    pdf = p.PdfReaderObject(file)

    spreadsheet_path = pdf.init_excel()
    print(Fore.YELLOW + f"spreadsheet path is {spreadsheet_path}")
    print(Fore.YELLOW + f"Size classes are hard coded as follows: Standard classes = {standard_classes_headers}, Customer defined classes = {defined_classes_headers}")

    row = 2
    # Now parse info from pdf. I'll keep this as straight forward as possible
    for i in range(0, pdf.num_pages, 2):
        contents1 = pdf.read_content(i)
        contents2 = pdf.read_content(i+1)
        sample_name = contents1[contents1.index("Sample ref.") + 14:contents1.index("Sample Name")]
        median = contents1[contents1.index('Diameter at 50%') + 18:contents1.index('µmDiameter at 90%')-1]
        mean = contents1[contents1.index('Mean diameter') + 16:contents1.index('FraunhoferDensity')-4]

        # Check index alignment on first sample
        if i >= 1000:
            print(Fore.YELLOW + "Whoa thats a lot of samples! ...Quitting Debug Mode...")
        else:
            print(Fore.YELLOW + f"...Testing Index Alignment on sample {i} with ID {sample_name}...\n")
            print(f"name:{sample_name}mean: {mean}\nmedian: {median}")

            if " " in sample_name or " " in mean or " " in median: # If there are spaces before/after
                print(Fore.RED + "Misalignment in indexing size metrics. Check that pdf version == or that the sample name does not contain spaces. ...Continuing with incorrect indexing to shwo error effects...")
            else:
                print(Fore.GREEN + "Correct indexing size metrics ... checking size classes..." + Fore.WHITE)

                contents1 = contents1[620:]    

                defined_class_distrib = []
                print(Fore.YELLOW + "...solving Indexing problem with start=anchor+6, end=anchor+11 assuming Cilas outputs values as length 5 FPNs")
                for classs in defined_classes_headers:
                    s = contents1.index(classs)
                    try:
                        val = float(contents1[s+6:s+11])     # Catches if output contains characters --> can not convert a misallignment to float
                        defined_class_distrib.append(val)
                    except:
                        print(Fore.RED + f"Misalignment in indexing size classes. Can not cast {contents1[s+6:s+11]} to float ...continuing with bad solution...")
                        defined_class_distrib.append(contents1[s+6:s+11])
                print(Fore.YELLOW + f"Customer defined ouput: {defined_class_distrib}")
                #contents2 = contents2[620:]
                contents2 = contents2[674:]   ### Trimming a bit more extraneous text off
                contents2 = contents2.replace('\n', '')  ### stripping newline characters
                contents2 = contents2.replace('xQ3q3', ' ')  ### stripping row names
                contents2 = contents2.strip()  ### removes leading/trailing spaces

                standard_class_distrib = []
                for classs in standard_classes_headers:
                    s = contents2.index(classs)
                    #s = my_string.index(classs)

                    ## TESTING \/\/\/
                    #print(str(s) + ' = index of class heading')
                    #print(classs + ' = class heading')
                    # val = float(contents2[s + 12:s + 16])  ## individual size class value (q3)
                    # print(str(val) + ' = val')
                    #print(str(contents2[s+12:s+17]) + ' = val')
                    ## TESTING /\/\/\


                    try:
                        val = float(contents2[s+12:s+17])   ## ere
                        standard_class_distrib.append(val)
                    except:
                        print(Fore.RED + f"Misalignment in indexing size classes. Can not cast {contents2[s+12:s+17]} to float ...continuing with bad solution...")
                        standard_class_distrib.append(contents2[s+12:s+17])
                print(Fore.YELLOW + f"Standard defined output: {standard_class_distrib}")

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