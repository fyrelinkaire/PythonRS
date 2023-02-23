# -*- coding: utf-8 -*-
"""
Created on Tue Jan 31 14:03:01 2023

@author: ALiu3
"""




from PyPDF2 import PdfReader
import re

pdf_location = r'I:\Physics\aliu3\PatientList.pdf'
    
reader = PdfReader(pdf_location)
page = reader.pages[0]

pdfFile = page.extract_text()

re_numPage = re.findall(r'([A-Z].+$)', pdfFile) #page#

numPage = re_numPage[0][:7]

# re_numQAs = re.findall(r"(?<=Measurements).+", pdfFile) #number of QAs

# numQAs = re_numQAs[0][-2:]

# re_lsNames = re.findall(r"(?<=\n)+(?=\w).+[^\d]+(?<!\n)[A-Z]+[\,].+", pdfFile) #Patient Names

re_lsNames = re.findall(r"(?<!\n_).+", pdfFile)

# test = re.sub(r'\d+', '', re_lsNames)

# re_lsNames = [s.strip() for s in re_lsNames]

lsNames = re_lsNames

re_lsMMRN = re.findall(r'(?<=\()(.*?)(?=\))', pdfFile)

re_lsMMRN = [s.strip() for s in re_lsMMRN]

lsMMRN = re_lsMMRN[1:]

search_mv = re.findall(r'[A-Z].+[=x]+', pdfFile)
find_mv = list(map(lambda test: test.replace('Comment :', ''), search_mv))

mmrn_dict = dict((zip(lsNames,lsMMRN)))

MV_dict = dict((zip(lsNames,find_mv)))



print(re_lsNames)
# print(test)
# print(lsNames)
# print(lsMMRN)
# print(mmrn_dict)
# print(MV_dict)
# print(dict3)
# 