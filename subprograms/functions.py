from subprograms.parameters import *


def show_exception_and_exit(exc_type, exc_value, tb):
    # https:/www.youtube.com/watch?v=8MjfalI4AO8
    # traceback.print_exception(exc_type, exc_value, tb)
    #pidKillFinish(finish=False, exc='')
    logging.error(exc_value, exc_info=(exc_type, exc_value, tb))
    with open(f'{emailsPath}/emailError.html', 'r', encoding='utf=8') as fileError:
        templateError = fileError.read()
        htmlError = templateError.format(exc_value=exc_value, month=month, year=year)
    yag.send(to=str(emailReceiversError).split(','), subject=f"ERRO - BOT Procedimento RE e RET",
    contents=htmlError,
    attachments=f'{logPath}/{today}/{logFileName}.log')
    logging.info("Error e-mail sent.")
    sys.exit(1)


# ------------------------------------------------------------------------------------

def emailSuccess():
    with open(f'{emailsPath}/emailSuccess.html', 'r', encoding='utf=8') as fileSuccess:
        templateSuccess = fileSuccess.read()
        htmlSuccess = templateSuccess.format(month=month, year=year)
    yag.send(to=str(emailReceiversSuccess).split(','), subject=f"SUCESSO - BOT Procedimento RE e RET",
    contents=htmlSuccess)#,
    #attachments=f'{logPath}/{today}/fopag_relatorios_bancarios_{logFileName}.log')
    logging.info("Success e-mail sent.")
    sys.exit(0)


# ------------------------------------------------------------------------------------

def is_capslock_on():
    return True if ctypes.WinDLL("User32.dll").GetKeyState(0x14) else False


# ------------------------------------------------------------------------------------


def turn_capslock_off():
    if is_capslock_on():
        keyboard.press("capslock")
        logging.info('CAPS LOCK is now deactivated.')


# ------------------------------------------------------------------------------------

def is_numlock_off():
    return not ctypes.windll.user32.GetKeyState(0x90) > 0

def turn_numlock_on():
    if is_numlock_off():
        keyboard.press("capslock")
        logging.info('NUM LOCK is now activated.')


# ------------------------------------------------------------------------------------

def sharepointGetBearerToken():
    
    # URL da requisição
    url = sharepointGetBearerTokenUrl

    # Dados do corpo da requisição x-www-form-urlencoded
    data = {
        "grant_type": "client_credentials",
        "client_id": sharepointClientId,
        "client_secret": sharepointClientSecret,
        "resource": sharepointResource
    }

    # Faz a solicitação POST com dados x-www-form-urlencoded
    response = requests.post(url, data=data)

    # Verifica a resposta
    if response.status_code == 200:
        logging.debug("Get Bearer Token request successfull.")
        #print("Resposta:", json.loads(response.text)['access_token'])
        return json.loads(response.text)['access_token']
    else:
        logging.error("Get Bearer Token Request error. Status code:", response.status_code)
        logging.error("Get Bearer Token Request response:", response.text)


# ------------------------------------------------------------------------------------

def sharepointGetExcelFile():
    token = sharepointGetBearerToken()
    # URL da solicitação para o arquivo "Planilha_modelo_CENTROS DE CUSTOS_DOCUMENTACAO.xlsx"
    url_da_pasta = sharepointGetExcelFileUrl

    # Cabeçalhos da solicitação
    headers = {
        "Authorization": f"Bearer {token}",  # Substitua {seu_token} pelo seu token de autenticação real
    }

    # Faz a solicitação GET para obter o arquivo
    response = requests.get(url_da_pasta, headers=headers)

    # Verifica a resposta
    if response.status_code == 200:
        # A resposta contém o arquivo ou seus metadados
        # Você pode salvar o arquivo em disco ou processá-lo conforme necessário
        with open(f'{excelFilePath}/{excelFileName}.xlsx', "wb") as f:
            f.write(response.content)
        logging.info('Request to get "Planilha_modelo_CENTROS DE CUSTOS_DOCUMENTACAO.xlsx" worksheet successfull.')
    else:
        logging.error(f'ERROR in request to get "Planilha_modelo_CENTROS DE CUSTOS_DOCUMENTACAO.xlsx". Response: {response.status_code} - {response.text}')
        show_exception_and_exit(f'ERROR in request to get "Planilha_modelo_CENTROS DE CUSTOS_DOCUMENTACAO.xlsx". Response: {response.status_code} - {response.text}', f'ERROR in request to get "Planilha_modelo_CENTROS DE CUSTOS_DOCUMENTACAO.xlsx". Response: {response.status_code} - {response.text}', f'ERROR in request to get "Planilha_modelo_CENTROS DE CUSTOS_DOCUMENTACAO.xlsx". Response: {response.status_code} - {response.text}')


# ------------------------------------------------------------------------------------

def groupIds(df):
    # Criar uma lista para armazenar os grupos
    grupos = []
    grupo_atual = []

    # Iterar pelas linhas do DataFrame
    for index, row in df.iterrows():
        #valor = row.iloc[0]  # Obtém o valor da primeira coluna
        valor = row.iloc[1]  # Obtém o valor da primeira coluna
        logging.debug(f'row: {row}')

        #if not pd.isna(valor):  # Verifica se o valor não é NaN
        if not pd.isna(valor) and not any(char.isdigit() for char in valor):
            logging.debug(f'teste valor nome: {row.iloc[1]}')
            grupo_atual.append(valor)  # Adiciona o valor ao grupo atual
        else:
            if grupo_atual:  # Verifica se o grupo atual não está vazio
                grupos.append(grupo_atual.copy())  # Adiciona o grupo atual à lista de grupos
                grupo_atual.clear()  # Limpa o grupo atual

    # Adiciona o último grupo (se houver)
    if grupo_atual:
        grupos.append(grupo_atual)

    return grupos


# ------------------------------------------------------------------------------------

def groupEquipes(df):
    # Criar uma lista para armazenar os grupos
    grupos = []
    grupo_atual = []

    # Iterar pelas linhas do DataFrame
    for index, row in df.iterrows():
        valor = row.iloc[1]  # Obtém o valor da primeira coluna
        unidade = row.iloc[2]

        if not pd.isna(valor) and '-' in valor:  # Verifica se o valor não é NaN
            logging.debug(f'unidade: {unidade}')
            grupos.append([valor, unidade])  # Adiciona o valor ao grupo atual


    return grupos


# ------------------------------------------------------------------------------------

def onlyEquipeNumbers(equipes, ids):
    numberEquipes = []

    for i, equipe in enumerate(equipes):
        numberEquipes.append(re.findall(r'\d{8}', equipe[0]))#.group())
    
    # Cria uma lista de dicionários
    equipes_ids_list = []

    # Itera sobre as listas equipes e ids para criar as estruturas desejadas e adicioná-las à lista
    for equipe, id_list, fullEquipe in zip(numberEquipes, ids, equipes):
        #print(equipe, id_list, fullEquipe)
        # Usando o primeiro elemento da lista equipe como chave
        
        chave = equipe[0]
        unidade = fullEquipe[1]
        logging.warning(unidade)
        
        # Usando a lista id como valor
        valor = id_list
        # Cria a estrutura desejada
        equipe_dict = [{(chave, unidade): valor}]
        # Adiciona a estrutura à lista
        equipes_ids_list.append(equipe_dict)


    
    return equipes_ids_list
    

# ------------------------------------------------------------------------------------

def validateOccurrencesNumbersEquipe(equipes_ids_list):
    
    # Dicionário para rastrear o número de ocorrências de cada chave
    ocorrencias = defaultdict(int)
    
    # Lista para armazenar o resultado
    resultado = []

    for dicionario in equipes_ids_list:
        novo_dicionario = {}  # Dicionário modificado para armazenar as chaves com o número de ocorrências
        for subdicionario in dicionario:
            for chave, valores in subdicionario.items():
                ocorrencia_atual = ocorrencias[chave] + 1
                ocorrencias[chave] = ocorrencia_atual
                if ocorrencia_atual > 1:
                    #chave_modificada = f"{chave} ({ocorrencia_atual})"
                    chave_modificada = (f"{chave[0]} ({ocorrencia_atual})", chave[1])
                else:
                    chave_modificada = chave
                novo_dicionario[chave_modificada] = valores
        resultado.append(novo_dicionario)

    return resultado 


# ------------------------------------------------------------------------------------

def joinIdsAndEquipes(df):
    ids = groupIds(df=df)
    equipes = groupEquipes(df=df)
    numberEquipes = onlyEquipeNumbers(equipes=equipes, ids=ids)
    finalListDict = validateOccurrencesNumbersEquipe(equipes_ids_list=numberEquipes)

    return finalListDict


# ------------------------------------------------------------------------------------

def cmatEngenhariaSheetData():
    
    df = pd.read_excel(io=f'{excelFilePath}/{excelFileName}.xlsx', header = 0, sheet_name=sheet1Name).dropna(axis=0, how='any')

    if len(df) == 0:
        logging.warning(f'Empty "{sheet1Name}" sheet. Ignoring...')
        return []
    
    else:
        column1, column2, column3, column4 = df.head(0)
        listOfDictionaries = []
        for index, row in enumerate(df.iterrows()): 
            listOfDictionaries.append({'col1': df[column1].iloc[index], 'col2': df[column2].iloc[index], 'col3': df[column3].iloc[index], 'col4': df[column4].iloc[index]})
        logging.info(f'{sheet1Name} full dict list: {listOfDictionaries}')

        return listOfDictionaries


# ------------------------------------------------------------------------------------

def cmatServicosSheetData():
    
    df = pd.read_excel(io=f'{excelFilePath}/{excelFileName}.xlsx', header = 1, sheet_name=sheet2Name).dropna(axis=0, how='any')

    if len(df) == 0:
        logging.warning(f'Empty "{sheet2Name}" sheet. Ignoring...')
        return []
    
    else:
        column1, column2, column3, column4 = df.head(0)
        listOfDictionaries = []
        for index, row in enumerate(df.iterrows()): 
            listOfDictionaries.append({'col1': df[column1].iloc[index], 'col2': df[column2].iloc[index], 'col3': df[column3].iloc[index], 'col4': df[column4].iloc[index]})
        logging.info(f'{sheet2Name} full dict list: {listOfDictionaries}')
        return listOfDictionaries


# ------------------------------------------------------------------------------------

def equipesDeMontagemSheetData():
    
    df = pd.read_excel(io=f'{excelFilePath}/{excelFileName}.xlsx', header = 0, sheet_name=sheet5Name)

    if len(df) == 0:
        logging.warning(f'Empty "{sheet5Name}" sheet. Ignoring...')
        return []
    
    else:
        listOfDictionaries = joinIdsAndEquipes(df)
        logging.info(f'{sheet5Name} full dict list: {listOfDictionaries}')
        return listOfDictionaries


# ------------------------------------------------------------------------------------

def logProcessingDuration(startTimer):
    endTimer = timer()
    time_difference = endTimer - startTimer
    minutes, seconds = divmod(int(time_difference), 60)
    milliseconds = int((time_difference - int(time_difference)) * 1000)
    logging.info(f"All done in {minutes}:{seconds:02}.{milliseconds:03} minutes.")


# ------------------------------------------------------------------------------------


def sharepointSendLog():
    
    codEmpresas = ['002', '065']

    for codEmpresa in codEmpresas:

        token = sharepointGetBearerToken()

        # URL da biblioteca do SharePoint
        url_biblioteca = "https://concrematcorp.sharepoint.com/teams/departamentopessoal"

        baseDirectory = f'{url_biblioteca}/Documentos Compartilhados/FERIAS'

        if codEmpresa == '002':
            preDirectorySplit = f"{sharepointFolderNameCodEmpresa002}/{year}/{month}{year}/DOCUMENTO RETORNO BANCARIO"
            directorySplit = preDirectorySplit.split('/')
        elif codEmpresa == '065':
            preDirectorySplit = f"{sharepointFolderNameCodEmpresa065}/{year}/{month}{year}/DOCUMENTO RETORNO BANCARIO"
            directorySplit = preDirectorySplit.split('/')

        for folder in directorySplit:

            # Cabeçalhos para autenticação
            headers = {
                'Authorization': f'Bearer {token}',  # Substitua pelo seu token de autenticação
                'Content-Type': 'application/json;odata=verbose',
            }

            # Define o corpo da solicitação para criar a nova pasta
            data = {
                '__metadata': {
                    'type': 'SP.Folder',
                },
                'ServerRelativeUrl': f"{baseDirectory}/{folder}",
            }

            baseDirectory = f"{baseDirectory}/{folder}"

            count = 0

            while count <= sharepointRetriesRequest: 

                logging.info(f'Try {count+1}/{sharepointRetriesRequest} to create the folder "{folder}" in Sharepoint.')


                # Faz a solicitação POST para criar a pasta
                response = requests.post(f"{url_biblioteca}/_api/web/folders", headers=headers, json=data)

                logging.debug(f'fileCreateMultipleFolder response: {response.status_code} - {response.text}')

                #print(response.text)
                # Verifica se a pasta foi criada com sucesso
                if response.status_code == 201:
                    logging.info(f'The folder "{folder}" was created succesfully in Sharepoint.')
                    break
                else:
                    if count == sharepointRetriesRequest:
                        logging.error(f"ERROR. Could not create folder in Sharepoint. Status code: {response.status_code}")
                        logging.error(f"ERROR response text:", response.text)
                        show_exception_and_exit(exc_type=f'Could not created folder "{folder}" in Sharepoint. Status Code: {response.status_code}. Text: {response.text}', exc_value=f'Could not created folder "{folder}" in Sharepoint. Status Code: {response.status_code}. Text: {response.text}', tb=f'Could not created folder "{folder}" in Sharepoint. Status Code: {response.status_code}. Text: {response.text}')
                    else:
                        logging.warning(f'Try {count+1}/{sharepointRetriesRequest} failed.')
                        count+=1
                        sleep(2)
                        continue


        logging.info(f'Directory "{baseDirectory}" created successfully.')
    
        # finalizou a criacao das pastas


        # comecou o processo de upload dos logs

        fileDirectory = f'{logPath}/{today}'
        fileName = f'{logFileName}.log'

        if codEmpresa == '002':
            directory = f"/teams/departamentopessoal/Documentos Compartilhados/FERIAS/{sharepointFolderNameCodEmpresa002}/{year}/{month}{year}/DOCUMENTO RETORNO BANCARIO"
        elif codEmpresa == '065':
            directory = directory = f"/teams/departamentopessoal/Documentos Compartilhados/FERIAS/{sharepointFolderNameCodEmpresa065}/{year}/{month}{year}/DOCUMENTO RETORNO BANCARIO"

        token = sharepointGetBearerToken()

        # URL de destino no SharePoint
        url_destino = f"https://concrematcorp.sharepoint.com/teams/departamentopessoal/_api/Web/GetFolderByServerRelativeUrl('{directory}')/Files/add(url='{fileName}', overwrite=true)"

        # Cabeçalhos da solicitação com token de autenticação
        headers = {
            "Authorization": f"Bearer {token}",  # Substitua pelo seu token de autenticação real
            "Content-Type": "application/json;odata=verbose",
        }

        # Corpo da solicitação com o arquivo a ser enviado
        # Certifique-se de ajustar o nome do arquivo e o caminho do arquivo no disco
        arquivo_a_enviar = open(f'{fileDirectory}/{fileName}', "rb")

        count = 0

        while count <= sharepointRetriesRequest: 

            logging.debug(f'Try {count+1}/{sharepointRetriesRequest} to upload "{fileDirectory}/{fileName}" to Sharepoint.')

            token = sharepointGetBearerToken()

            # Cabeçalhos para autenticação
            headers = {
            'Authorization': f'Bearer {token}',  # Substitua pelo seu token de autenticação
            'Content-Type': 'application/json;odata=verbose',
            }

            # Faz a solicitação POST para enviar o arquivo
            response = requests.post(url_destino, headers=headers, data=arquivo_a_enviar)

            logging.debug(f'fileUploadSharepoint response: {response.status_code} - {response.text}')

            # Verifica a resposta
            if response.status_code == 200:
                logging.info(f'File "{fileName}" uploaded to Sharepoint successfully.')
                break
            else:
                if count == sharepointRetriesRequest:
                    logging.error("ERROR. Could not upload file to Sharepoint. Status code:", response.status_code)
                    logging.error("ERROR response text:", response.text)
                    show_exception_and_exit(exc_type=f'Could not upload file "{fileDirectory}/{fileName}" to Sharepoint. Status Code: {response.status_code}. Text: {response.text}', exc_value=f'Could not upload file "{fileDirectory}/{fileName}" to Sharepoint. Status Code: {response.status_code}. Text: {response.text}', tb=f'Could not upload file "{fileDirectory}/{fileName}" to Sharepoint. Status Code: {response.status_code}. Text: {response.text}')
                else:
                    logging.warning(f'Try {count+1}/{sharepointRetriesRequest} failed. {response.status_code} - {response.text}')
                    count+=1
                    sleep(2)
                    continue


# ------------------------------------------------------------------------------------

def extractRetData(retPath, personName):
    limitedPersonName = personName[:30]
    # Nome do arquivo
    #filename = 'ret.ret'

    with open(retPath, 'r', encoding='utf-8') as file:
        # Iterar sobre as linhas do arquivo
        for line_number, line in enumerate(file, start=1):
            # Verificar se limitedPersonName está na linha
            if limitedPersonName in line:
                logging.debug(f'The row {line_number} contains "{limitedPersonName}".')
                #print(line)
                nextLine = next(file)
                #print(nextLine)
                break

    #brokenLine = line.split('              ')
    brokenLine = line
    brokenNextLine = nextLine.split('   ')
    #print(brokenLine)
    #print(brokenNextLine)

    preCpf = brokenNextLine[1][1:][:14]
    #print(preCpf)
    cpf = preCpf[-11:]
    logging.debug(f'CPF: {cpf}')


    #preDataDoCredito = re.search(r'(.*)BRL', brokenLine[1]).group(1)
    preDataDoCredito = re.search(r'(.*)BRL', brokenLine).group(1)
    #preDataDoCredito = re.search(r'(.*)BRL', brokenLine[0]).group(1)
    #print(preDataDoCredito)
    preDataDoCredito = preDataDoCredito[-8:]
    #print(preDataDoCredito)
    preDataDoCredito = re.match(r"(\d{2})(\d{2})(\d{4})", preDataDoCredito)
    dataDoCredito = f'{preDataDoCredito.group(1)}/{preDataDoCredito.group(2)}/{preDataDoCredito.group(3)}'
    logging.debug(f'Data do Crédito: {dataDoCredito}')


    #preAgenciaAndCC = re.search(fr'(.*?) {limitedPersonName}', brokenLine[0]).group(1)
    preAgenciaAndCC = re.search(fr'(.*?) {limitedPersonName}', brokenLine).group(1)
    preAgenciaAndCC = re.search(r'A(.*)', preAgenciaAndCC).group(1)
    preAgenciaAndCC = preAgenciaAndCC[10:]
    #print(preAgenciaAndCC)
    agencia = preAgenciaAndCC[:4]
    digitoDaAgencia = preAgenciaAndCC[4]
    fullAgencia = f'{agencia}-{digitoDaAgencia}'
    preCC = preAgenciaAndCC[6:]
    contaCorrente = preCC[:-1]
    digitoVerificador = preCC[-1]
    fullContaCorrente = f'{contaCorrente}-{digitoVerificador}'
    logging.debug(f'Agência: {agencia}')
    logging.debug(f'Dígito da Agência: {digitoDaAgencia}')
    logging.debug(f'Conta Corrente: {contaCorrente}')
    logging.debug(f'Dígito Verificador: {digitoVerificador}')


    #preValor = re.search(r'BRL(.*?)PGIT', brokenLine[1]).group(1)
    preValor = re.search(r'BRL(.*?)PGIT', brokenLine).group(1)
    valor = str(float(preValor[:-2] + '.' + preValor[-2:]))
    logging.debug(f'Valor: {valor}')

    #autenticacao = re.search(r'PGIT\d{12}', brokenLine[1]).group(0)
    autenticacao = re.search(r'PGIT\d{12}', brokenLine).group(0)
    logging.debug(f'Autenticação: {autenticacao}')


    completeList = [dataDoCredito, agencia, digitoDaAgencia, contaCorrente, digitoVerificador, \
                valor, autenticacao]


    return completeList


# ------------------------------------------------------------------------------------

def checkNameInPdf(personName, pdfPath):

    with pdfplumber.open(pdfPath) as pdf:

        for page in pdf.pages:

            logging.debug(f'page.extract_text(): {page.extract_text()}')

            if personName in page.extract_text():
                pageNumber = page.page_number

                return pageNumber
            
        return None


# ------------------------------------------------------------------------------------

def legacySharepointGetDocumentFile(filial, contratoPuro, personName, documentType):

    if documentType == 'REL FUNCS':
        document = summarizedRelFuncsName
    
    elif document == 'RET':
        document = summarizedRetName

    elif document == 'RE':
        document = summarizedReName

    token = sharepointGetBearerToken()
    # URL da solicitação para o arquivo "Planilha_modelo_CENTROS DE CUSTOS_DOCUMENTACAO.xlsx"
    #url_da_pasta = f"https://concrematcorp.sharepoint.com/teams/departamentopessoal/_api/Web/GetFileByServerRelativePath(decodedurl='/teams/departamentopessoal/Documentos%20Compartilhados/DOCUMENTOS%20MEDI%C3%87%C3%83O%20(SHAREPOINT)/{contratoPuro}/{year}/{month}{year}')/Folders"
    #url_da_pasta_com_filial = f"https://concrematcorp.sharepoint.com/teams/departamentopessoal/_api/Web/GetFileByServerRelativePath(decodedurl='/teams/departamentopessoal/Documentos%20Compartilhados/DOCUMENTOS%20MEDI%C3%87%C3%83O%20(SHAREPOINT)/{contratoPuro}/{filial}/{year}/{month}{year}')/Folders"
    url_da_pasta = f"https://concrematcorp.sharepoint.com/teams/departamentopessoal/_api/Web/GetFolderByServerRelativePath(decodedurl='/teams/departamentopessoal/Documentos%20Compartilhados/DOCUMENTOS%20MEDI%C3%87%C3%83O%20(SHAREPOINT)/{contratoPuro}/{year}/{month}{year}')/Folders"
    url_da_pasta_com_filial = f"https://concrematcorp.sharepoint.com/teams/departamentopessoal/_api/Web/GetFolderByServerRelativePath(decodedurl='/teams/departamentopessoal/Documentos%20Compartilhados/DOCUMENTOS%20MEDI%C3%87%C3%83O%20(SHAREPOINT)/{contratoPuro}/{filial}/{year}/{month}{year}')/Folders"


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
    response = requests.get(url_da_pasta_com_filial, headers=headers, json=dataFolder)

    
    

    if response.status_code == 404:
        
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
        if '{    "d": {        "results": []    }}' in response.text:
    
            url_dos_arquivos = url_da_pasta.replace('Folders', 'Files')
            response = requests.get(url_dos_arquivos, headers=headers)
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

                    with open(f'{outputPdfPath}/{personName}.pdf', 'wb') as file:
                        file.write(response.content)

                    sleep(1)

                    pageNumber = checkNameInPdf(personName=personName, pdfPath=f'{outputPdfPath}/{personName}.pdf')
                    
                    if pageNumber == None:
                        os.remove(f'{outputPdfPath}/{personName}.pdf')
                        continue
                    
                    else:
                        logging.info(f'"{personName}" found for {documentType}.')
                        return pageNumber, f'{outputPdfPath}/{personName}.pdf'

        else:
            
            resultFolders = responseJson['d']['results']

            endpointFilesList = []
            for result in resultFolders:
                endpointFiles = result['__metadata']['Files']
                endpointFilesList.append(endpointFiles)

            for endpointFile in endpointFilesList:

                token = sharepointGetBearerToken()

                # Cabeçalhos da solicitação
                headers = {
                "Authorization": f"Bearer {token}",  # Substitua {seu_token} pelo seu token de autenticação real
                'Content-Type': 'application/json;odata=verbose',
                "Accept": "application/json;odata=verbose"
                }


                response = requests.get(f'{endpointFile}', headers=headers)

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

                        with open(f'{outputPdfPath}/{personName}.pdf', 'wb') as file:
                            file.write(response.content)

                        sleep(1)

                        pageNumber = checkNameInPdf(personName=personName, pdfPath=f'{outputPdfPath}/{personName}.pdf')
                        
                        if pageNumber == None:
                            os.remove(f'{outputPdfPath}/{personName}.pdf')
                            continue
                        
                        else:
                            logging.info(f'"{personName}" found for {documentType}.')
                            return pageNumber, f'{outputPdfPath}/{personName}.pdf'

    else:

        token = sharepointGetBearerToken()

        # Cabeçalhos da solicitação
        headers = {
        "Authorization": f"Bearer {token}",  # Substitua {seu_token} pelo seu token de autenticação real
        'Content-Type': 'application/json;odata=verbose',
        "Accept": "application/json;odata=verbose"
        }

        # Faz a solicitação GET para obter as pastas
        response = requests.get(url_da_pasta_com_filial, headers=headers, json=dataFolder)
        
        responseJson = response.json()
        
        # checa se o resultado de pastas foi vazio, se sim, ele vai procurar o arquivo nessa pasta
        if '{    "d": {        "results": []    }}' in response.text:
    
            url_dos_arquivos = url_da_pasta_com_filial.replace('Folders', 'Files')
            response = requests.get(url_da_pasta_com_filial, headers=headers)
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

                    with open(f'{outputPdfPath}/{personName}.pdf', 'wb') as file:
                        file.write(response.content)

                    sleep(1)

                    pageNumber = checkNameInPdf(personName=personName, pdfPath=f'{outputPdfPath}/{personName}.pdf')
                    
                    if pageNumber == None:
                        os.remove(f'{outputPdfPath}/{personName}.pdf')
                        continue
                    
                    else:
                        logging.info(f'"{personName}" found for {documentType}.')
                        return pageNumber, f'{outputPdfPath}/{personName}.pdf'

        else:
            
            resultFolders = responseJson['d']['results']

            endpointFilesList = []
            for result in resultFolders:
                endpointFiles = result['__metadata']['Files']
                endpointFilesList.append(endpointFiles)

            for endpointFile in endpointFilesList:

                token = sharepointGetBearerToken()

                # Cabeçalhos da solicitação
                headers = {
                "Authorization": f"Bearer {token}",  # Substitua {seu_token} pelo seu token de autenticação real
                'Content-Type': 'application/json;odata=verbose',
                "Accept": "application/json;odata=verbose"
                }


                response = requests.get(f'{endpointFile}', headers=headers)

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

                        with open(f'{outputPdfPath}/{personName}.pdf', 'wb') as file:
                            file.write(response.content)

                        sleep(1)

                        pageNumber = checkNameInPdf(personName=personName, pdfPath=f'{outputPdfPath}/{personName}.pdf')
                        
                        if pageNumber == None:
                            os.remove(f'{outputPdfPath}/{personName}.pdf')
                            continue
                        
                        else:
                            logging.info(f'"{personName}" found for {documentType}.')
                            return pageNumber, f'{outputPdfPath}/{personName}.pdf'


#----------------------------------------------------------------------





def legacySharepointGetRelFuncsFile(filial, contratoPuro):

    #document = summarizedRelFuncsName
    document = 'FUNCS'
    
    token = sharepointGetBearerToken()
    # URL da solicitação para o arquivo "Planilha_modelo_CENTROS DE CUSTOS_DOCUMENTACAO.xlsx"
    #url_da_pasta = f"https://concrematcorp.sharepoint.com/teams/departamentopessoal/_api/Web/GetFileByServerRelativePath(decodedurl='/teams/departamentopessoal/Documentos%20Compartilhados/DOCUMENTOS%20MEDI%C3%87%C3%83O%20(SHAREPOINT)/{contratoPuro}/{year}/{month}{year}')/Folders"
    url_da_pasta = f"https://concrematcorp.sharepoint.com/teams/departamentopessoal/_api/Web/GetFolderByServerRelativePath(decodedurl='/teams/departamentopessoal/Documentos%20Compartilhados/DOCUMENTOS%20MEDI%C3%87%C3%83O%20(SHAREPOINT)/{contratoPuro}/{year}/{month}{year}/{filial}')/Folders"
    #url_da_pasta = f"https://concrematcorp.sharepoint.com/teams/departamentopessoal/_api/Web/GetFolderByServerRelativePath(decodedurl='/teams/departamentopessoal/Documentos%20Compartilhados/DOCUMENTOS%20MEDI%C3%87%C3%83O%20(SHAREPOINT)/{contratoPuro}/{year}/{month}{year}')/Folders"
    #url_da_pasta_com_filial = f"https://concrematcorp.sharepoint.com/teams/departamentopessoal/_api/Web/GetFolderByServerRelativePath(decodedurl='/teams/departamentopessoal/Documentos%20Compartilhados/DOCUMENTOS%20MEDI%C3%87%C3%83O%20(SHAREPOINT)/{contratoPuro}/{filial}/{year}/{month}{year}')/Folders"


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

    
    

    #if response.status_code == 404:

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
        
        logging.warning(f'response.text inicial aqui: {str(response.text)}')

        logging.warning(f'response.json() inicial aqui: {str(response.json())}')
        
        #sys.exit()

        # checa se o resultado de pastas foi vazio, se sim, ele vai procurar o arquivo nessa pasta
        #if '{    "d": {        "results": []    }}' in response.text:
        if '{"d":{"results":[]}}' in response.text:

            logging.warning('Entrou no if resposta vazia')

            url_dos_arquivos = url_da_pasta.replace('Folders', 'Files')
            logging.warning(f'url_dos_arquivos: {url_dos_arquivos}')
            response = requests.get(url_dos_arquivos, headers=headers)
            responseJson = response.json()
            
            resultsFiles = responseJson['d']['results']

            logging.warning(f'lista de resultsFiles: {resultsFiles}')

            for resultFile in resultsFiles:
                
                token = sharepointGetBearerToken()

                # Cabeçalhos da solicitação
                headers = {
                "Authorization": f"Bearer {token}",  # Substitua {seu_token} pelo seu token de autenticação real
                'Content-Type': 'application/json;odata=verbose',
                "Accept": "application/json;odata=verbose"
                }
                
                
                endpointFile = resultFile['__metadata']['id']

                logging.debug(f'endpointFile: {endpointFile}')

                if document not in endpointFile:
                    continue

                else:

                    response = requests.get(f'{endpointFile}/$value', headers=headers)

                    with open(f'{outputPdfPath}/{contratoPuro}.pdf', 'wb') as file:
                        file.write(response.content)

                    sleep(1)

                    logging.info(f'"{outputPdfPath}/{contratoPuro}.pdf" created. (REL FUNCS)')

                    return f'{outputPdfPath}/{contratoPuro}.pdf'

                    '''
                    pageNumber = checkNameInPdf(personName=personName, pdfPath=f'{outputPdfPath}/{personName}.pdf')
                    
                    if pageNumber == None:
                        logging.warning(f'Could not find "{personName}" in {endpointFile}')
                        os.remove(f'{outputPdfPath}/{personName}.pdf')
                        continue
                    
                    else:
                        logging.info(f'"{personName}" found for {document}')
                        return pageNumber, f'{outputPdfPath}/{personName}.pdf'

                    '''
                        
        else:
            
            '''
            url_dos_arquivos = url_da_pasta.replace('Folders', 'Files')
            logging.debug(url_dos_arquivos)
            response = requests.get(url_dos_arquivos, headers=headers)
            responseJson = response.json()
            
            resultsFiles = responseJson['d']['results']

            logging.warning(f'lista de resultsFiles: {resultsFiles}')

            for resultFile in resultsFiles:
                
                token = sharepointGetBearerToken()

                # Cabeçalhos da solicitação
                headers = {
                "Authorization": f"Bearer {token}",  # Substitua {seu_token} pelo seu token de autenticação real
                'Content-Type': 'application/json;odata=verbose',
                "Accept": "application/json;odata=verbose"
                }
                
                
                endpointFile = resultFile['__metadata']['id']

                logging.debug(f'endpointFile: {endpointFile}')

                if document not in endpointFile:
                    continue

                else:

                    response = requests.get(f'{endpointFile}/$value', headers=headers)

                    with open(f'{outputPdfPath}/{contratoPuro}.pdf', 'wb') as file:
                        file.write(response.content)

                    sleep(1)

                    logging.info(f'"{outputPdfPath}/{contratoPuro}.pdf" created. (REL FUNCS)')

                    return f'{outputPdfPath}/{contratoPuro}.pdf'
            '''

            resultFolders = responseJson['d']['results']

            #logging.warning('Entrou no else resposta nao vazia')
            #logging.warning(f'lista de resultFolders: {resultFolders}')

            endpointFilesList = []
            for result in resultFolders:
                endpointFiles = result['Files']['__deferred']['uri']
                endpointFilesList.append(endpointFiles)


            for endpointFile in endpointFilesList:

                token = sharepointGetBearerToken()

                logging.warning(f'endpointFile: {endpointFile}')

                # Cabeçalhos da solicitação
                headers = {
                "Authorization": f"Bearer {token}",  # Substitua {seu_token} pelo seu token de autenticação real
                'Content-Type': 'application/json;odata=verbose',
                "Accept": "application/json;odata=verbose"
                }


                response = requests.get(f'{endpointFile}', headers=headers)
                responseJson = response.json()

                logging.warning(f'requisicaoFile: {response.text}')

                resultsFiles = responseJson['d']['results']

                logging.warning(f'resultsFiles: {resultsFiles}')


                for resultFile in resultsFiles:
                    
                    token = sharepointGetBearerToken()

                    # Cabeçalhos da solicitação
                    headers = {
                    "Authorization": f"Bearer {token}",  # Substitua {seu_token} pelo seu token de autenticação real
                    'Content-Type': 'application/json;odata=verbose',
                    "Accept": "application/json;odata=verbose"
                    }
                    
                    
                    endpointFile = resultFile['__metadata']['id']

                    logging.warning(f'esse aqui {endpointFile}')

                    if document not in endpointFile:
                        continue

                    else:

                        response = requests.get(f'{endpointFile}/$value', headers=headers)

                        with open(f'{outputPdfPath}/{contratoPuro}.pdf', 'wb') as file:
                            file.write(response.content)

                        sleep(1)

                        logging.info(f'"{outputPdfPath}/{contratoPuro}.pdf" created. (REL FUNCS)')

                        return f'{outputPdfPath}/{contratoPuro}.pdf'

                        '''
                        pageNumber = checkNameInPdf(personName=personName, pdfPath=f'{outputPdfPath}/{personName}.pdf')
                        
                        if pageNumber == None:
                            logging.warning(f'Could not find "{personName}" in {endpointFile}')
                            os.remove(f'{outputPdfPath}/{personName}.pdf')
                            continue
                        
                        else:
                            logging.info(f'"{personName}" found for {document}.')
                            return pageNumber, f'{outputPdfPath}/{personName}.pdf'
                        '''
        
    return None


    '''
    else:

        token = sharepointGetBearerToken()

        # Cabeçalhos da solicitação
        headers = {
        "Authorization": f"Bearer {token}",  # Substitua {seu_token} pelo seu token de autenticação real
        'Content-Type': 'application/json;odata=verbose',
        "Accept": "application/json;odata=verbose"
        }

        # Faz a solicitação GET para obter as pastas
        response = requests.get(url_da_pasta_com_filial, headers=headers, json=dataFolder)
        
        responseJson = response.json()
        
        # checa se o resultado de pastas foi vazio, se sim, ele vai procurar o arquivo nessa pasta
        if '{    "d": {        "results": []    }}' in response.text:
    
            url_dos_arquivos = url_da_pasta_com_filial.replace('Folders', 'Files')
            response = requests.get(url_da_pasta_com_filial, headers=headers)
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

                    with open(f'{outputPdfPath}/{personName}.pdf', 'wb') as file:
                        file.write(response.content)

                    sleep(1)

                    pageNumber = checkNameInPdf(personName=personName, pdfPath=f'{outputPdfPath}/{personName}.pdf')
                    
                    if pageNumber == None:
                        os.remove(f'{outputPdfPath}/{personName}.pdf')
                        continue
                    
                    else:
                        logging.info(f'"{personName}" found for {documentType}.')
                        return pageNumber, f'{outputPdfPath}/{personName}.pdf'

        else:
            
            resultFolders = responseJson['d']['results']

            endpointFilesList = []
            for result in resultFolders:
                endpointFiles = result['__metadata']['Files']
                endpointFilesList.append(endpointFiles)

            for endpointFile in endpointFilesList:

                token = sharepointGetBearerToken()

                # Cabeçalhos da solicitação
                headers = {
                "Authorization": f"Bearer {token}",  # Substitua {seu_token} pelo seu token de autenticação real
                'Content-Type': 'application/json;odata=verbose',
                "Accept": "application/json;odata=verbose"
                }


                response = requests.get(f'{endpointFile}', headers=headers)

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

                        with open(f'{outputPdfPath}/{personName}.pdf', 'wb') as file:
                            file.write(response.content)

                        sleep(1)

                        pageNumber = checkNameInPdf(personName=personName, pdfPath=f'{outputPdfPath}/{personName}.pdf')
                        
                        if pageNumber == None:
                            os.remove(f'{outputPdfPath}/{personName}.pdf')
                            continue
                        
                        else:
                            logging.info(f'"{personName}" found for {documentType}.')
                            return pageNumber, f'{outputPdfPath}/{personName}.pdf'

        '''


#----------------------------------------------------------------------

def sharepointGetRelFuncsFile(filial, contratoPuro, edm):

    #document = summarizedRelFuncsName
    document = 'FUNCS'
    
    token = sharepointGetBearerToken()
    # URL da solicitação para o arquivo "Planilha_modelo_CENTROS DE CUSTOS_DOCUMENTACAO.xlsx"
    #url_da_pasta = f"https://concrematcorp.sharepoint.com/teams/departamentopessoal/_api/Web/GetFileByServerRelativePath(decodedurl='/teams/departamentopessoal/Documentos%20Compartilhados/DOCUMENTOS%20MEDI%C3%87%C3%83O%20(SHAREPOINT)/{contratoPuro}/{year}/{month}{year}')/Folders"
    if edm == False:
        url_da_pasta = f"https://concrematcorp.sharepoint.com/teams/departamentopessoal/_api/Web/GetFolderByServerRelativePath(decodedurl='/teams/departamentopessoal/Documentos%20Compartilhados/DOCUMENTOS%20MEDI%C3%87%C3%83O%20(SHAREPOINT)/{contratoPuro}/{year}/{month}{year}/{filial}')/Folders"
    else:
        url_da_pasta = f"https://concrematcorp.sharepoint.com/teams/departamentopessoal/_api/Web/GetFolderByServerRelativePath(decodedurl='/teams/departamentopessoal/Documentos%20Compartilhados/DOCUMENTOS%20MEDI%C3%87%C3%83O%20(SHAREPOINT)/{contratoPuro}/{year}/{month}{year}')/Folders"

    #url_da_pasta = f"https://concrematcorp.sharepoint.com/teams/departamentopessoal/_api/Web/GetFolderByServerRelativePath(decodedurl='/teams/departamentopessoal/Documentos%20Compartilhados/DOCUMENTOS%20MEDI%C3%87%C3%83O%20(SHAREPOINT)/{contratoPuro}/{year}/{month}{year}')/Folders"
    #url_da_pasta_com_filial = f"https://concrematcorp.sharepoint.com/teams/departamentopessoal/_api/Web/GetFolderByServerRelativePath(decodedurl='/teams/departamentopessoal/Documentos%20Compartilhados/DOCUMENTOS%20MEDI%C3%87%C3%83O%20(SHAREPOINT)/{contratoPuro}/{filial}/{year}/{month}{year}')/Folders"


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

    
    

    #if response.status_code == 404:

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
        
        logging.warning(f'response.text inicial aqui: {str(response.text)}')

        logging.warning(f'response.json() inicial aqui: {str(response.json())}')
        
        #sys.exit()

        # checa se o resultado de pastas foi vazio, se sim, ele vai procurar o arquivo nessa pasta
        #if '{    "d": {        "results": []    }}' in response.text:
        if '{"d":{"results":[]}}' in response.text:

            logging.warning('Entrou no if resposta vazia')

            url_dos_arquivos = url_da_pasta.replace('Folders', 'Files')
            logging.warning(f'url_dos_arquivos: {url_dos_arquivos}')
            response = requests.get(url_dos_arquivos, headers=headers)
            responseJson = response.json()
            
            resultsFiles = responseJson['d']['results']

            logging.warning(f'lista de resultsFiles: {resultsFiles}')

            for resultFile in resultsFiles:
                
                token = sharepointGetBearerToken()

                # Cabeçalhos da solicitação
                headers = {
                "Authorization": f"Bearer {token}",  # Substitua {seu_token} pelo seu token de autenticação real
                'Content-Type': 'application/json;odata=verbose',
                "Accept": "application/json;odata=verbose"
                }
                
                
                endpointFile = resultFile['__metadata']['id']

                logging.debug(f'endpointFile: {endpointFile}')

                if document not in endpointFile:
                    continue

                else:

                    response = requests.get(f'{endpointFile}/$value', headers=headers)

                    with open(f'{outputPdfPath}/{contratoPuro}.pdf', 'wb') as file:
                        file.write(response.content)

                    sleep(1)

                    logging.info(f'"{outputPdfPath}/{contratoPuro}.pdf" created. (REL FUNCS)')

                    return f'{outputPdfPath}/{contratoPuro}.pdf'

                    '''
                    pageNumber = checkNameInPdf(personName=personName, pdfPath=f'{outputPdfPath}/{personName}.pdf')
                    
                    if pageNumber == None:
                        logging.warning(f'Could not find "{personName}" in {endpointFile}')
                        os.remove(f'{outputPdfPath}/{personName}.pdf')
                        continue
                    
                    else:
                        logging.info(f'"{personName}" found for {document}')
                        return pageNumber, f'{outputPdfPath}/{personName}.pdf'

                    '''
                        
        else:
            
            '''
            url_dos_arquivos = url_da_pasta.replace('Folders', 'Files')
            logging.debug(url_dos_arquivos)
            response = requests.get(url_dos_arquivos, headers=headers)
            responseJson = response.json()
            
            resultsFiles = responseJson['d']['results']

            logging.warning(f'lista de resultsFiles: {resultsFiles}')

            for resultFile in resultsFiles:
                
                token = sharepointGetBearerToken()

                # Cabeçalhos da solicitação
                headers = {
                "Authorization": f"Bearer {token}",  # Substitua {seu_token} pelo seu token de autenticação real
                'Content-Type': 'application/json;odata=verbose',
                "Accept": "application/json;odata=verbose"
                }
                
                
                endpointFile = resultFile['__metadata']['id']

                logging.debug(f'endpointFile: {endpointFile}')

                if document not in endpointFile:
                    continue

                else:

                    response = requests.get(f'{endpointFile}/$value', headers=headers)

                    with open(f'{outputPdfPath}/{contratoPuro}.pdf', 'wb') as file:
                        file.write(response.content)

                    sleep(1)

                    logging.info(f'"{outputPdfPath}/{contratoPuro}.pdf" created. (REL FUNCS)')

                    return f'{outputPdfPath}/{contratoPuro}.pdf'
            '''

            resultFolders = responseJson['d']['results']

            #logging.warning('Entrou no else resposta nao vazia')
            #logging.warning(f'lista de resultFolders: {resultFolders}')

            endpointFilesList = []
            for result in resultFolders:
                endpointFiles = result['Files']['__deferred']['uri']
                endpointFilesList.append(endpointFiles)


            for endpointFile in endpointFilesList:

                token = sharepointGetBearerToken()

                logging.warning(f'endpointFile: {endpointFile}')

                # Cabeçalhos da solicitação
                headers = {
                "Authorization": f"Bearer {token}",  # Substitua {seu_token} pelo seu token de autenticação real
                'Content-Type': 'application/json;odata=verbose',
                "Accept": "application/json;odata=verbose"
                }


                response = requests.get(f'{endpointFile}', headers=headers)
                responseJson = response.json()

                logging.warning(f'requisicaoFile: {response.text}')

                resultsFiles = responseJson['d']['results']

                logging.warning(f'resultsFiles: {resultsFiles}')


                for resultFile in resultsFiles:
                    
                    token = sharepointGetBearerToken()

                    # Cabeçalhos da solicitação
                    headers = {
                    "Authorization": f"Bearer {token}",  # Substitua {seu_token} pelo seu token de autenticação real
                    'Content-Type': 'application/json;odata=verbose',
                    "Accept": "application/json;odata=verbose"
                    }
                    
                    
                    endpointFile = resultFile['__metadata']['id']

                    logging.warning(f'esse aqui {endpointFile}')

                    if document not in endpointFile:
                        continue

                    else:

                        response = requests.get(f'{endpointFile}/$value', headers=headers)

                        with open(f'{outputPdfPath}/{contratoPuro}.pdf', 'wb') as file:
                            file.write(response.content)

                        sleep(1)

                        logging.info(f'"{outputPdfPath}/{contratoPuro}.pdf" created. (REL FUNCS)')

                        return f'{outputPdfPath}/{contratoPuro}.pdf'

                        '''
                        pageNumber = checkNameInPdf(personName=personName, pdfPath=f'{outputPdfPath}/{personName}.pdf')
                        
                        if pageNumber == None:
                            logging.warning(f'Could not find "{personName}" in {endpointFile}')
                            os.remove(f'{outputPdfPath}/{personName}.pdf')
                            continue
                        
                        else:
                            logging.info(f'"{personName}" found for {document}.')
                            return pageNumber, f'{outputPdfPath}/{personName}.pdf'
                        '''
        
    return None


    '''
    else:

        token = sharepointGetBearerToken()

        # Cabeçalhos da solicitação
        headers = {
        "Authorization": f"Bearer {token}",  # Substitua {seu_token} pelo seu token de autenticação real
        'Content-Type': 'application/json;odata=verbose',
        "Accept": "application/json;odata=verbose"
        }

        # Faz a solicitação GET para obter as pastas
        response = requests.get(url_da_pasta_com_filial, headers=headers, json=dataFolder)
        
        responseJson = response.json()
        
        # checa se o resultado de pastas foi vazio, se sim, ele vai procurar o arquivo nessa pasta
        if '{    "d": {        "results": []    }}' in response.text:
    
            url_dos_arquivos = url_da_pasta_com_filial.replace('Folders', 'Files')
            response = requests.get(url_da_pasta_com_filial, headers=headers)
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

                    with open(f'{outputPdfPath}/{personName}.pdf', 'wb') as file:
                        file.write(response.content)

                    sleep(1)

                    pageNumber = checkNameInPdf(personName=personName, pdfPath=f'{outputPdfPath}/{personName}.pdf')
                    
                    if pageNumber == None:
                        os.remove(f'{outputPdfPath}/{personName}.pdf')
                        continue
                    
                    else:
                        logging.info(f'"{personName}" found for {documentType}.')
                        return pageNumber, f'{outputPdfPath}/{personName}.pdf'

        else:
            
            resultFolders = responseJson['d']['results']

            endpointFilesList = []
            for result in resultFolders:
                endpointFiles = result['__metadata']['Files']
                endpointFilesList.append(endpointFiles)

            for endpointFile in endpointFilesList:

                token = sharepointGetBearerToken()

                # Cabeçalhos da solicitação
                headers = {
                "Authorization": f"Bearer {token}",  # Substitua {seu_token} pelo seu token de autenticação real
                'Content-Type': 'application/json;odata=verbose',
                "Accept": "application/json;odata=verbose"
                }


                response = requests.get(f'{endpointFile}', headers=headers)

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

                        with open(f'{outputPdfPath}/{personName}.pdf', 'wb') as file:
                            file.write(response.content)

                        sleep(1)

                        pageNumber = checkNameInPdf(personName=personName, pdfPath=f'{outputPdfPath}/{personName}.pdf')
                        
                        if pageNumber == None:
                            os.remove(f'{outputPdfPath}/{personName}.pdf')
                            continue
                        
                        else:
                            logging.info(f'"{personName}" found for {documentType}.')
                            return pageNumber, f'{outputPdfPath}/{personName}.pdf'

        '''



#----------------------------------------------------------------------



def sharepointGetReDocumentFile(personName, sheetName, documentType):

    logging.info(f'sheetName: {sheetName}')


    if documentType == 'RET':
        summarizedDocument = summarizedRetName 
    
    elif documentType == 'RE':
        summarizedDocument = summarizedReName

    if sheetName == sheet1Name:
        prefix = prefixCmatEngenhariaDocument
    elif sheetName == sheet2Name:
        prefix = prefixCmatServicosDocument
    elif sheetName == sheet5Name:
        #prefix = prefixEquipesDeMontagemDocument
        prefix = prefixCmatEngenhariaDocument
        logging.info(f'prefix: {prefix}')

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

                        with open(f'{outputPdfPath}/{personName}.pdf', 'wb') as file:
                            file.write(response.content)

                        sleep(1)

                        pageNumber = checkNameInPdf(personName=personName, pdfPath=f'{outputPdfPath}/{personName}.pdf')
                        
                        if pageNumber == None:
                            logging.warning(f'Could not find "{personName}" in {endpointFile}')
                            os.remove(f'{outputPdfPath}/{personName}.pdf')
                            flagNextFolderNumber = True
                            break
                        
                        else:
                            logging.info(f'"{personName}" found for {documentType} documentType.')
                            return [pageNumber, f'{outputPdfPath}/{personName}.pdf', folderNumber]

                if flagNextFolderNumber == True:
                    continue
            
    return []


#----------------------------------------------------------------------

def extractRelFuncsPdfNames(pdfPath):
    with pdfplumber.open(pdfPath) as pdf:
        
        fullListNames = []

        for page in pdf.pages:
            logging.debug(f'page.extract_text(): {page.extract_text()}')
            with open('resultado.txt', 'w', encoding='utf-8') as file:
                file.write(page.extract_text())

            names = re.findall(fr'{regexNamesRelFuncs}', page.extract_text())
            for name in names:
                fullListNames.append(name)

        logging.info(fullListNames)

        return fullListNames

            
#----------------------------------------------------------------------

def extractNumeroInscricaoFromRe(pageNumber, pdfPath):
    with pdfplumber.open(pdfPath) as pdf:

        for page in pdf.pages:
            
            if page.page_number == pageNumber:
                listedPage = page.extract_text().split('\n')
                for index, list in enumerate(listedPage):
                    if 'TOMADOR/OBRA' in list:
                        inscricao = re.search(r'INSCRIÇÃO:(.*)', list).group(1)
                        logging.info(f'Inscrição: {inscricao}')
                        return inscricao

#----------------------------------------------------------------------

def extractBothNumeroInscricaoFromRe(pageNumber, pdfPath):
    with pdfplumber.open(pdfPath) as pdf:

        inscricoes = []

        #for page in pdf.pages:
            
        #    if page.page_number == pageNumber:
        page = pdf.pages[pageNumber]
        logging.debug(f'PAGE EXTRACT TEXT BOTH{page.extract_text()}')
        listedPage = page.extract_text().split('\n')
        for index, list in enumerate(listedPage):
            if 'TOMADOR/OBRA' in list:
                inscricaoTomador = re.search(r'INSCRIÇÃO:(.*)', list).group(1)
                logging.debug(f'Inscrição do Tomador: {inscricaoTomador}')
                inscricoes.append(inscricaoTomador)

            if 'EMPRESA:' in list:
                    inscricaoFilial = re.search(r'INSCRIÇÃO:(.*)', list).group(1)
                    logging.debug(f'Inscrição da Filial: {inscricaoFilial}')
                    inscricoes.append(inscricaoFilial)

        return inscricoes

#----------------------------------------------------------------------

def pagesWithSameInscricaoTomadorRe(pdfPath, inscricao):
    
    listPagesWithSameInscricao = []

    with pdfplumber.open(pdfPath) as pdf:

        for page in pdf.pages:
            
            #if page.page_number == pageNumber:
            listedPage = page.extract_text().split('\n')
            for index, list in enumerate(listedPage):
                if 'TOMADOR/OBRA' in list:
                    inscricaoFromPage = re.search(r'INSCRIÇÃO:(.*)', list).group(1)
                    logging.debug(f'Inscrição from page: {inscricaoFromPage} ; Inscrição from parameter: {inscricao}')
                    if inscricaoFromPage.strip() == inscricao.strip():
                        listPagesWithSameInscricao.append(page.page_number)
                        logging.debug(f'page {page.page_number} - listaAqui: {list}')
                        logging.debug('Same "Inscrição" Tomador.')
        
        logging.info(f'len(listPagesWithSameInscricao): {len(listPagesWithSameInscricao)}')
        logging.info(f'listPagesWithSameInscricao: {listPagesWithSameInscricao}')
        return listPagesWithSameInscricao


#----------------------------------------------------------------------

def pagesWithSameInscricaoFilialRe(pdfPath, inscricao):

    
    listPagesWithSameInscricao = []

    with pdfplumber.open(pdfPath) as pdf:

        for page in pdf.pages:
            
            listedPage = page.extract_text().split('\n')
            for index, list in enumerate(listedPage):
                #if 'TOMADOR/OBRA' in list:
                
                if 'EMPRESA:' in list:
                    inscricaoFromPage = re.search(r'INSCRIÇÃO:(.*)', list).group(1)
                    logging.info(f'Inscrição from page: {inscricaoFromPage} ; Inscrição from parameter: {inscricao}')
                    if inscricaoFromPage.strip() == inscricao.strip():
                        listPagesWithSameInscricao.append(page.page_number)
                        logging.debug(f'page {page.page_number} - listaAqui: {list}')
                        logging.debug('Same "Inscrição" Filial.')

        logging.info(f'len(listPagesWithSameInscricao): {len(listPagesWithSameInscricao)}')
        logging.info(f'listPagesWithSameInscricao: {listPagesWithSameInscricao}')


        listPagesResumoDoFechamentoEmpresa = []

        with pdfplumber.open(pdfPath) as pdf:

            for pageWithSameInscricao in listPagesWithSameInscricao:

                #for page in pdf.pages:

                    #if page.page_number == pageWithSameInscricao:
                page = pdf.pages[pageWithSameInscricao]
                
                listedPage = page.extract_text().split('\n')
                for index, list in enumerate(listedPage):
                    if 'RESUMO DO FECHAMENTO - EMPRESA' in list:
                        #listPagesResumoDoFechamentoEmpresa.append(page.page_number)
                        logging.debug(f'"RESUMO DO FECHAMENTO - EMPRESA" found in page {pageWithSameInscricao}.')
                        listPagesResumoDoFechamentoEmpresa.append(pageWithSameInscricao)


        logging.info(f'len(listPagesResumoDoFechamentoEmpresa): {len(listPagesResumoDoFechamentoEmpresa)}')
        logging.info(f'listPagesResumoDoFechamentoEmpresa: {listPagesResumoDoFechamentoEmpresa}')

        
        return listPagesResumoDoFechamentoEmpresa

#----------------------------------------------------------------------

def newPagesWithSameInscricaoFilialRe(pdfPath, inscricaoFilial):
    
    listPagesWithSameInscricao = []

    with pdfplumber.open(pdfPath) as pdf:

        for page in pdf.pages:
            
            listedPage = page.extract_text().split('\n')
            for index, list in enumerate(listedPage):
                #if 'TOMADOR/OBRA' in list:
                
                if 'EMPRESA:' in list:
                    inscricaoFromPage = re.search(r'INSCRIÇÃO:(.*)', list).group(1)
                    logging.debug(f'Inscrição from page: {inscricaoFromPage} ; Inscrição from parameter: {inscricaoFilial}')
                    if inscricaoFromPage.strip() == inscricaoFilial.strip():
                        listPagesWithSameInscricao.append(page.page_number)
                        logging.debug(f'page {page.page_number} - listaAqui: {list}')
                        logging.debug('Same "Inscrição" Filial.')

        logging.info(f'len(listPagesWithSameInscricao): {len(listPagesWithSameInscricao)}')
        logging.info(f'listPagesWithSameInscricao: {listPagesWithSameInscricao}')


        listPagesResumoDoFechamentoEmpresa = []

        with pdfplumber.open(pdfPath) as pdf:

            for pageWithSameInscricao in listPagesWithSameInscricao:
                
                logging.debug(f'pageWithSameInscricao: {pageWithSameInscricao} - {listPagesWithSameInscricao}')
                page = pdf.pages[pageWithSameInscricao-1]
                
                listedPage = page.extract_text().split('\n')
                for index, list in enumerate(listedPage):
                    if 'RESUMO DO FECHAMENTO - EMPRESA' in list:
                        #listPagesResumoDoFechamentoEmpresa.append(page.page_number)
                        logging.info(f'"RESUMO DO FECHAMENTO - EMPRESA" found in page {pageWithSameInscricao}.')
                        listPagesResumoDoFechamentoEmpresa.append(pageWithSameInscricao)


            logging.info(f'len(listPagesResumoDoFechamentoEmpresa): {len(listPagesResumoDoFechamentoEmpresa)}')
            logging.info(f'listPagesResumoDoFechamentoEmpresa: {listPagesResumoDoFechamentoEmpresa}')

        if listPagesResumoDoFechamentoEmpresa == []:
            return listPagesResumoDoFechamentoEmpresa
        
        else:
            numberToAdd = int(listPagesResumoDoFechamentoEmpresa[-1]+1)
            listPagesResumoDoFechamentoEmpresa.append(numberToAdd)
            return listPagesResumoDoFechamentoEmpresa

#----------------------------------------------------------------------

def printPdfPages(pdfPath, outputPath, fileName, pagesList):
    
    # Abre o arquivo PDF de entrada
    with open(pdfPath, 'rb') as input_file:
        pdf_reader = PyPDF2.PdfReader(input_file)

        # Cria um novo PDF para imprimir
        pdf_writer = PyPDF2.PdfWriter()

        # Adiciona as páginas desejadas ao novo PDF
        for pagina_num in pagesList:
            if 1 <= pagina_num <= len(pdf_reader.pages):  # Garante que a página existe no PDF
                pagina_desejada = pdf_reader.pages[pagina_num - 1]
                pdf_writer.add_page(pagina_desejada)

        # Salva o novo PDF
        with open(f'{outputPath}/{fileName}', 'wb') as output_file:
            pdf_writer.write(output_file)

    logging.info(f'The pages {pagesList} were created in a single PDF file and saved in "{outputPath}/{fileName}".')

#----------------------------------------------------------------------

#def sharepointCreateMultipleFolder(directory, fileDirectory, fileName):
def sharepointCreateMultipleFolder(directory, fileDirectory, fileName, unParam):

    directorySplit = directory.split('/')

    token = sharepointGetBearerToken()

    # URL da biblioteca do SharePoint
    #url_biblioteca = "https://concrematcorp.sharepoint.com/teams/departamentopessoal"
    url_biblioteca = f"https://concrematcorp.sharepoint.com/teams/{unParam}"

    #baseDirectory = f'{url_biblioteca}/Documentos Compartilhados/DOCUMENTOS MEDIÇÃO (SHAREPOINT)'
    baseDirectory = f'{url_biblioteca}/Documentos Compartilhados'

    for folder in directorySplit:

        # Cabeçalhos para autenticação
        headers = {
            'Authorization': f'Bearer {token}',  # Substitua pelo seu token de autenticação
            'Content-Type': 'application/json;odata=verbose',
        }

        # Define o corpo da solicitação para criar a nova pasta
        data = {
            '__metadata': {
                'type': 'SP.Folder',
            },
            'ServerRelativeUrl': f"{baseDirectory}/{folder}",
        }

        baseDirectory = f"{baseDirectory}/{folder}"

        count = 0

        while count <= sharepointRetriesRequest: 

            logging.info(f'Try {count+1}/{sharepointRetriesRequest} to create the folder "{folder}" in Sharepoint.')


            # Faz a solicitação POST para criar a pasta
            response = requests.post(f"{url_biblioteca}/_api/web/folders", headers=headers, json=data)

            logging.debug(f'fileCreateMultipleFolder response: {response.status_code} - {response.text}')

            #print(response.text)
            # Verifica se a pasta foi criada com sucesso
            if response.status_code == 201:
                logging.info(f'The folder "{folder}" was created succesfully in Sharepoint.')
                break
            else:
                if count == sharepointRetriesRequest:
                    logging.error(f"ERROR. Could not create folder in Sharepoint. Status code: {response.status_code}")
                    logging.error(f"ERROR response text:", response.text)
                    show_exception_and_exit(exc_type=f'Could not created folder "{folder}" in Sharepoint. Status Code: {response.status_code}. Text: {response.text}', exc_value=f'Could not created folder "{folder}" in Sharepoint. Status Code: {response.status_code}. Text: {response.text}', tb=f'Could not created folder "{folder}" in Sharepoint. Status Code: {response.status_code}. Text: {response.text}')
                else:
                    logging.warning(f'Try {count+1}/{sharepointRetriesRequest} failed.')
                    count+=1
                    sleep(2)
                    continue


    logging.info(f'Directory "{baseDirectory}" created successfully.')
    baseDirectory = re.search(regexUrlTeams, baseDirectory).group(0)
    return [baseDirectory, fileDirectory, fileName]

#----------------------------------------------------------------------


#def sharepointSendFile(directory, fileDirectory, fileName):
def sharepointSendFile(directory, fileDirectory, fileName, unParam):

    token = sharepointGetBearerToken()

    #url_destino = f"https://concrematcorp.sharepoint.com/teams/departamentopessoal/_api/Web/GetFolderByServerRelativeUrl('{directory}')/Files/add(url='{fileName}', overwrite=true)"
    url_destino = f"https://concrematcorp.sharepoint.com/teams/{unParam}/_api/Web/GetFolderByServerRelativeUrl('{directory}')/Files/add(url='{fileName}', overwrite=true)"

    logging.info(f'url_destino sharepointSendFile: {url_destino}')

    # Cabeçalhos da solicitação com token de autenticação
    headers = {
        "Authorization": f"Bearer {token}",  # Substitua pelo seu token de autenticação real
        "Content-Type": "application/json;odata=verbose",
    }

    # Corpo da solicitação com o arquivo a ser enviado
    # Certifique-se de ajustar o nome do arquivo e o caminho do arquivo no disco
    arquivo_a_enviar = open(f'{fileDirectory}/{fileName}', "rb")


    count = 0

    while count <= sharepointRetriesRequest: 

        logging.debug(f'Try {count+1}/{sharepointRetriesRequest} to upload "{fileDirectory}/{fileName}" to Sharepoint.')

        # Faz a solicitação POST para enviar o arquivo
        response = requests.post(url_destino, headers=headers, data=arquivo_a_enviar)

        logging.debug(f'fileUploadSharepoint response: {response.status_code} - {response.text}')

        # Verifica a resposta
        if response.status_code == 200:
            logging.info(f'File "{fileName}" uploaded to Sharepoint successfully.\n')
            break
        else:
            if count == sharepointRetriesRequest:
                logging.error("ERROR. Could not upload file to Sharepoint. Status code:", response.status_code)
                logging.error("ERROR response text:", response.text)
                show_exception_and_exit(exc_type=f'Could not upload file "{fileDirectory}/{fileName}" to Sharepoint. Status Code: {response.status_code}. Text: {response.text}', exc_value=f'Could not upload file "{fileDirectory}/{fileName}" to Sharepoint. Status Code: {response.status_code}. Text: {response.text}', tb=f'Could not upload file "{fileDirectory}/{fileName}" to Sharepoint. Status Code: {response.status_code}. Text: {response.text}')
            else:
                logging.warning(f'Try {count+1}/{sharepointRetriesRequest} failed.')
                count+=1
                sleep(2)
                continue

#----------------------------------------------------------------------

def sharepointCreateFoldersAndUploadFile(directory, fileDirectory, fileName, unParam):
    returnCmf = sharepointCreateMultipleFolder(directory, fileDirectory, fileName, unParam)
    directory, fileDirectory, fileName = returnCmf
    sharepointSendFile(directory, fileDirectory, fileName, unParam)


#----------------------------------------------------------------------

def extractRetPages(pdfPath, inscricaoDoTomadorRe):

    logging.info(f'"Inscrição do Tomador" to be found in RET file: {inscricaoDoTomadorRe}')
    with pdfplumber.open(pdfPath) as pdf:

            for index, page in enumerate(pdf.pages):

                listedPage = page.extract_text().split('\n')
                for i, item in enumerate(listedPage):
                    if i == 2:
                        item3 = item
                    logging.debug(f'item: {item}')
                    if 'TOMADOR/OBRA' in item and 'INSCRIÇÃO' in item:
                        logging.debug(f'Encontrou TOMADOR/OBRA: {item}')
                        inscricao = re.search(r'INSCRIÇÃO:(.*) N', item).group(1)
                        logging.debug(f'{inscricao.strip()} == {inscricaoDoTomadorRe.strip()}')
                        if inscricao.strip() == inscricaoDoTomadorRe.strip():
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




#----------------------------------------------------------------------

def sharepointGetRetDocumentFile(inscricaoTomadorRe, sheetName):

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

                        #groupOfPages = extractRetPages(pdfPath=f'{outputPdfPath}/arquivoRET.pdf', inscricaoDoTomadorRe=inscricaoTomadorRe)
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


# ------------------------------------------------------------------------------------

def deletePdfDirectories():
  
    diretorio = outputPdfPath

    try:
        for root, dirs, files in os.walk(diretorio, topdown=False):
            for file in files:
                file_path = os.path.join(root, file)
                os.remove(file_path)
                #logging.debug(f'File "{file_path}" deleted.')
            for dir in dirs:
                dir_path = os.path.join(root, dir)
                shutil.rmtree(dir_path)
                #logging.debug(f'Directory "{dir_path}" deleted.')

        logging.info(f'All files and directories inside "{diretorio}" deleted.')
    except Exception as e:
        logging.info(f'An error occurred: {e}')