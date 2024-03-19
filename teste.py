from subprograms.functions import *

#33.146.648/0003-91
def testeSharepointGetRetDocumentFile(inscricaoTomadorRe, sheetName):

    summarizedDocument = summarizedRetName 

    if sheetName == sheet1Name:
        prefix = prefixCmatEngenhariaDocument
    elif sheetName == sheet2Name:
        prefix = prefixCmatServicosDocument
    elif sheetName == sheet5Name:
        prefix = prefixEquipesDeMontagemDocument

    foldersNumbers = ['150', '155']    

    for folderNumber in foldersNumbers:

        document = f'{prefix}_{folderNumber}_{summarizedDocument}_{month}{year}.pdf'
        logging.info(f'Expected document: {document}')

        token = sharepointGetBearerToken()

        url_da_pasta = f"https://concrematcorp.sharepoint.com/teams/departamentopessoal/_api/Web/GetFolderByServerRelativePath(decodedurl='/teams/departamentopessoal/Documentos Compartilhados/ARQ_SEFIP_MENSAL/SEFIP_ANO_{year}/{month}_{year}/{sheetName}/{folderNumber}')/Files"
        logging.info(url_da_pasta)

        # Cabeçalhos da solicitação
        headers = {
            "Authorization": f"Bearer {token}",  # Substitua {seu_token} pelo seu token de autenticação real
            'Content-Type': 'application/json;odata=verbose',
            "Accept": "application/json;odata=verbose"
        }

        dataFolder = {
                    "__metadata":
                                    {
                                                    "type": "SP.Folder"
                                                
                                    },
                                    "ServerRelativeUrl": "/teams/departamentopessoal"
        }


        # Faz a solicitação GET para obter as pastas
        response = requests.get(url_da_pasta, headers=headers, json=dataFolder)

        if response.status_code == 200:    
            token = sharepointGetBearerToken()

            # Cabeçalhos da solicitação
            headers = {
            "Authorization": f"Bearer {token}",  # Substitua {seu_token} pelo seu token de autenticação real
            'Content-Type': 'application/json;odata=verbose',
            "Accept": "application/json;odata=verbose"
            }

            logging.warning(f'token: {token}')
            logging.warning(f'url_da_pasta: {url_da_pasta}')

            # Faz a solicitação GET para obter as pastas
            response = requests.get(url_da_pasta, headers=headers)#, json=dataFolder)
            
            responseJson = response.json()
            
            logging.warning(f'responseJson inicial aqui: {str(response.text)}')

            # checa se o resultado de pastas foi vazio, se sim, ele vai procurar o arquivo nessa pasta
            #if '{    "d": {        "results": []    }}' in response.text:
            if '{"d":{"results":[]}}' in response.text:

                logging.warning(f'No files were found.')
                
                return []

            else:

                flagNextFolderNumber = False

                resultsFiles = responseJson['d']['results']

                for resultFile in resultsFiles:
                    
                    token = sharepointGetBearerToken()

                    # Cabeçalhos da solicitação
                    headers = {
                    "Authorization": f"Bearer {token}",  # Substitua {seu_token} pelo seu token de autenticação real
                    'Content-Type': 'application/json;odata=verbose',
                    "Accept": "application/json;odata=verbose"
                    }
                    
                    
                    endpointFile = resultFile['__metadata']['id']

                    if document not in endpointFile:
                        continue

                    else:

                        response = requests.get(f'{endpointFile}/$value', headers=headers)

                        with open(f'{outputPdfPath}/arquivoRET.pdf', 'wb') as file:
                            file.write(response.content)

                        sleep(1)

                        groupOfPages = extractRetPages(pdfPath=f'{outputPdfPath}/arquivoRET.pdf', inscricaoDoTomadorRe=inscricaoTomadorRe)
                        
                        if groupOfPages == []:
                            logging.warning(f'Could not find "Inscrição do Tomador {inscricaoTomadorRe}" in {endpointFile}')
                            os.remove(f'{outputPdfPath}/arquivoRET.pdf')
                            flagNextFolderNumber = True
                            break
                        
                        else:
                            logging.info(f'"Inscrição do Tomador {inscricaoTomadorRe}" found for RET documentType.')
                            return [groupOfPages, f'{outputPdfPath}/arquivoRET.pdf', folderNumber]

                if flagNextFolderNumber == True:
                    continue
            
    return []



                        



def testeExtractRetPages(pdfPath, inscricaoDoTomadorRe):

    logging.info(f'"Inscrição do Tomador" to be found in RET file: {inscricaoDoTomadorRe}')
    with pdfplumber.open(pdfPath) as pdf:

            for index, page in enumerate(pdf.pages):

                listedPage = page.extract_text().split('\n')
                for i, item in enumerate(listedPage):
                    if i == 2:
                        item3 = item

                    logging.info(f'item: {item}')

                    if 'TOMADOR/OBRA :' in item:
                        inscricao = re.search(r'INSCRIÇÃO:(.*) N', item).group(1)
                        if inscricao.strip() == inscricaoDoTomadorRe:
                            #print(inscricao)
                            #print(item3)
                            logging.info(f'First page of the group (found "Inscrição do Tomador"): {page.page_number}')
                            regexFirstAndFinalPages =  re.search(r'(\d{4})\/(\d{4}$)', item3) 
                            blockFirstPage = regexFirstAndFinalPages.group(1)
                            blockFinalPage =  regexFirstAndFinalPages.group(2)
                            numBlockFirstPage = int(blockFirstPage)
                            numBlockFinalPage = int(blockFinalPage)
                            numLastPageInscricao = page.page_number + abs(numBlockFirstPage - numBlockFinalPage)
                            #lastPageInscricao = pdf.pages[numLastPageInscricao-1].extract_text()
                            logging.info(f'Last page of the group: {numLastPageInscricao}')
    
                            return [page.page_number, numLastPageInscricao]


            return []

'''pdfPath = r'C:/Users/Consultor_dclick2/Desktop/projetos/concremat-rpa-procedimento-re-ret/pdfs/10138510/2023/122023/SP/CMAT_150_RET_122023.pdf'
inscricaoTomadorRe = '33.146.648/0003-91'
#testeExtractRetPages(pdfPath, inscricaoDoTomadorRe)
testeSharepointGetRetDocumentFile(inscricaoTomadorRe=inscricaoTomadorRe, sheetName=sheet1Name)'''


basePath = os.getcwd()
print(basePath)