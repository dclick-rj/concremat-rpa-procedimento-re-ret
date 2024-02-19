import pdfplumber
import re
import sys
import os

def extractRetPages(pdfPath, inscricaoDoTomadorRe):

    print(f'"Inscrição do Tomador" to be found in RET file: {inscricaoDoTomadorRe}')
    with pdfplumber.open(pdfPath) as pdf:

            for index, page in enumerate(pdf.pages):

                listedPage = page.extract_text().split('\n')
                for i, item in enumerate(listedPage):
                    if i == 2:
                        item3 = item
                    print(f'item: {item}')
                    if 'TOMADOR/OBRA' in item and 'INSCRIÇÃO' in item:
                        with open('result.txt', 'w', encoding='utf-8') as file:
                            file.write(str(listedPage))
                        print(f'Encontrou TOMADOR/OBRA ({page.page_number}): {item}')
                        inscricao = re.search(r'INSCRIÇÃO:(.*) N', item).group(1)
                        print(f'{inscricao.strip()} == {inscricaoDoTomadorRe.strip()}')
                        if inscricao.strip() == inscricaoDoTomadorRe.strip():
                            print(f'First page of the group (found "Inscrição do Tomador"): {page.page_number}')
                            regexFirstAndFinalPages =  re.search(r'(\d{4})\/(\d{4}$)', item3) 
                            blockFirstPage = regexFirstAndFinalPages.group(1)
                            blockFinalPage =  regexFirstAndFinalPages.group(2)
                            numBlockFirstPage = int(blockFirstPage)
                            numBlockFinalPage = int(blockFinalPage)
                            numLastPageInscricao = page.page_number + abs(numBlockFirstPage - numBlockFinalPage)
                            #lastPageInscricao = pdf.pages[numLastPageInscricao-1].extract_text()
                            print(f'Last page of the group: {numLastPageInscricao}')
    
                            return [page.page_number, numLastPageInscricao]


            return []

pdfPath = f'{os.getcwd()}/pdfs/arquivoRet.pdf'
inscricao = '33.146.648/0003-91'


extractRetPages(pdfPath=pdfPath, inscricaoDoTomadorRe=inscricao)