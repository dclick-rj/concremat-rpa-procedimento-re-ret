import os
import sys
import yagmail
import yaml
import logging
from datetime import date, timedelta, datetime
from time import time, sleep
from timeit import default_timer as timer
import ctypes
import re
import signal
import re
import requests
import json
import shutil
import pdfplumber
import keyboard
from collections import defaultdict
import pandas as pd
import openpyxl
import PyPDF2

startTimer = timer()

# path configs
basePath = os.getcwd()
basePath = basePath.replace('\\', '/')
logPath = f'{basePath}/log'
emailsPath = f'{basePath}/emails'
excelFilePath = f'{basePath}/excel'
outputPdfPath = f'{basePath}/pdfs'

# parameter configs
with open(file=f'{basePath}/config.yaml', mode='r', encoding='utf-8') as config:
    config = yaml.safe_load(config)

emailLogin = config['email']['login']
emailPassword = config['email']['password']
emailReceiversError = config['email']['receiversError']
emailReceiversSuccess = config['email']['receiversSuccess']
logLevel = config['log']['logLevel']
regexContratos1 = config['regex']['contratos1']
regexUrlTeams = config['regex']['urlTeams']
regexSharepointFerias = config['regex']['sharepointFerias']
regexContratoPuro = config['regex']['contratoPuro']
regexNamesRelFuncs = config['regex']['namesRelFuncs']
excelFileName = config['name']['excelFile']
sharepointClientId = config['sharepoint']['clientId']
sharepointClientSecret = config['sharepoint']['clientSecret']
sharepointResource = config['sharepoint']['resource']
sharepointGetExcelFileUrl = config['sharepoint']['url']['GetExcelFile']
sharepointGetBearerTokenUrl = config['sharepoint']['url']['GetBearerToken']
sharepointRelatoriosIndividuaisRetUrl = config['sharepoint']['url']['relatoriosIndividuaisRet']
sharepointRelatoriosIndividuaisRetUploadFileUrl = config['sharepoint']['url']['relatoriosIndividuaisRetUploadFile']
sharepointRetriesRequest = config['sharepoint']['retriesRequest']
sharepointFolderNameCodEmpresa002 = config['sharepoint']['folder']['names']['codEmpresa002']
sharepointFolderNameCodEmpresa065 = config['sharepoint']['folder']['names']['codEmpresa065']
sharepointFolderNameDocumentoBancarioRetorno = config['sharepoint']['folder']['names']['documentoRetornoBancario']
sheet1Name = config['name']['sheet1']
sheet2Name = config['name']['sheet2']
sheet5Name = config['name']['sheet5']
sheetPRCEName = config['name']['sheetPRCE']
patternRelFuncsName= config['name']['relFuncs']
summarizedRetName= config['name']['ret']
summarizedReName = config['name']['re']
summarizedRelFuncsName = config['name']['relFuncs']
prefixCmatEngenhariaDocument = config['name']['prefixCmatEngenhariaDocument'] 
prefixCmatServicosDocument = config['name']['prefixCmatServicosDocument']
prefixEquipesDeMontagemDocument = config['name']['prefixEquipesDeMontagemDocument']
unConfig = config['un']

# log configs
today = datetime.today().strftime('%Y-%m-%d')
logFileName = f"fopag_re_ret_{datetime.now().strftime('%H_%M_%S')}"

if os.path.exists(f'{logPath}/{today}') == False:
    os.makedirs(f'{logPath}/{today}')



cbOutputPdfPath = outputPdfPath.replace('/', '\\')

#level=logLevel
logging.basicConfig(level=20,
                    datefmt='%d-%m-%Y %H:%M:%S',
                    format='%(asctime)s; %(levelname)s; %(module)s.%(funcName)s; %(message)s',
                    handlers=[
                        logging.FileHandler(filename=f'{logPath}/{today}/{logFileName}.log', mode='w', encoding='utf-8', delay=False),
                        logging.StreamHandler(sys.stdout)
                    ])

# Obtém a data atual
day = datetime.today()
realMonth = day.month

if realMonth == 1:
    year = (day.year-1)
    month = 12
    
else:
    year = day.year
    month = realMonth-1

month = str(month).zfill(2)

#month = '11'
#year = '2023'

logging.info(f'Month to be inserted as reference: "{month}".')    
logging.info(f'Year to be inserted as reference: "{year}".')    

yag = yagmail.SMTP(user=emailLogin, password=emailPassword)

logging.info('Parameter configs finished.\n')
