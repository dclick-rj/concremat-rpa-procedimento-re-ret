from subprograms.functions import *

#----------------------------------------------------------------


def processCmat(servicos):

    if servicos ==  True:
        logging.info('==== Starting CMAT SERVIÇOS process ====...\n')
        sheetName = sheet2Name
    else:
        logging.info('==== Starting CMAT ENGENHARIA process ====...\n')
        sheetName = sheet1Name

    if servicos ==  True:
        fullListCmat = cmatServicosSheetData()
    else:
        fullListCmat = cmatEngenhariaSheetData()

    flagIncompleteProcessingTemp = False
    flagIncompleteProcessingPerm = False

    for index, lista in enumerate(fullListCmat):
        
        if flagIncompleteProcessingTemp == True:
            flagIncompleteProcessingPerm = True

        filial = lista['col3']
        contrato = lista['col1']
        contratoPuro = re.search(fr'{regexContratoPuro}', contrato).group(1)

        unOriginal = lista['col4']
        unParam = unConfig[unOriginal]
        logging.info(f'unOriginal: {unOriginal}')
        logging.info(f'unParam: {unParam}')

        logging.info(f'--- Currently working on {index+1}° row. ({filial} - {contrato}) ---')

        relFuncsFile = sharepointGetRelFuncsFile(filial, contratoPuro, edm=False)

        if relFuncsFile == None:
            logging.warning(f'Could not find Rel Funcs file for "{contratoPuro}" contract.')
            flagIncompleteProcessingTemp = True
            continue
        
  
        logging.warning(f'FOUND Rel Funcs file for "{contratoPuro}" contract.')
    
        namesList = extractRelFuncsPdfNames(pdfPath=relFuncsFile)

        if namesList == []:
            logging.warning(f'No names were found in "{relFuncsFile}" file.')
            flagIncompleteProcessingTemp = True
            continue

        logging.info(f'namesList: {namesList}')


        #for name in namesList:
        
        name = namesList[0]

        
        logging.info(f'--RE process started for {contratoPuro}--')
        returnReFunction = sharepointGetReDocumentFile(personName=name, sheetName=sheetName, documentType='RE')
        logging.info(f'returnList: {returnReFunction}')
        if returnReFunction == []:
            logging.warning(f'Retornou vazio para "{contratoPuro}". Skipping register...')
            flagIncompleteProcessingTemp = True
            continue
        
        pageNumberFuncionario, pdfPath, folderNumber = returnReFunction
        
        logging.info(f'Page wich contains {name} - {pageNumberFuncionario}')
        logging.info(f'{type(pageNumberFuncionario)} - {pageNumberFuncionario}')
        if pdfPath == None:
            logging.info('Skipping register...')
            flagIncompleteProcessingTemp = True
            continue

        listPageNumberFuncionario = [pageNumberFuncionario]

        inscricoes = extractBothNumeroInscricaoFromRe(pageNumberFuncionario, pdfPath)
        if inscricoes == [] or len(inscricoes) != 2:
            logging.info(f'"inscricoes" not correct: "{inscricoes}". Skipping register...')
            flagIncompleteProcessingTemp = True
            continue
           
        #inscricaoTomadorRe, inscricaoFilialRe = inscricoes
        inscricaoFilialRe, inscricaoTomadorRe = inscricoes
        logging.info(f'inscricaoFilialRe: {inscricaoFilialRe}')
        logging.info(f'inscricaoTomadorRe: {inscricaoTomadorRe}')

        pagsComNumeroDeInscricaoTomadorRe = pagesWithSameInscricaoTomadorRe(pdfPath=pdfPath, inscricao=inscricaoTomadorRe)
        logging.info(f'{len(pagsComNumeroDeInscricaoTomadorRe)} pages with same "Inscrição de Tomador"  number were found: {pagsComNumeroDeInscricaoTomadorRe}')
        if pagsComNumeroDeInscricaoTomadorRe == []:
            logging.info('Skipping register...')
            flagIncompleteProcessingTemp = True
            continue


        #pagsComNumeroDeInscricaoFilialPlusResumoDoFechamentoEmpresaRe = pagesWithSameInscricaoFilialRe(pdfPath=pdfPath, inscricao=inscricaoTomadorRe)
        pagsComNumeroDeInscricaoFilialPlusResumoDoFechamentoEmpresaRe = newPagesWithSameInscricaoFilialRe(pdfPath=pdfPath, inscricaoFilial=inscricaoFilialRe)
        logging.info(f'{len(pagsComNumeroDeInscricaoFilialPlusResumoDoFechamentoEmpresaRe)} pages with same "Inscrição de Filial" + "Resumo do Fechamento - Empresa" number were found: {pagsComNumeroDeInscricaoFilialPlusResumoDoFechamentoEmpresaRe}')
        if pagsComNumeroDeInscricaoFilialPlusResumoDoFechamentoEmpresaRe == []:
            logging.info('Skipping register...')
            flagIncompleteProcessingTemp = True
            continue

        # Converter as listas para conjuntos
        conjunct1 = set(listPageNumberFuncionario)
        conjunct2 = set(pagsComNumeroDeInscricaoTomadorRe)
        conjunct3 = set(pagsComNumeroDeInscricaoFilialPlusResumoDoFechamentoEmpresaRe)

        # Unir os conjuntos sem repetir valores
        unitedConjunct = conjunct1.union(conjunct2, conjunct3)
        #unitedConjunct = conjunct2.union(conjunct3)
        # Converter de volta para uma lista
        pagesListToPrintPdfRe = list(unitedConjunct)
        pagesListToPrintPdfRe.sort()
        #pagesListToPrintPdfRe.remove(pageNumberFuncionario)
        #pagesListToPrintPdfRe.insert(0, pageNumberFuncionario)

        logging.info(f'RE - pagesListToPrintPdfSortedRe: {pagesListToPrintPdfRe}')

        if os.path.exists(f'{outputPdfPath}/{contratoPuro}/{year}/{month}{year}/{filial}') == False:
            os.makedirs(f'{outputPdfPath}/{contratoPuro}/{year}/{month}{year}/{filial}')           

        #funcSPdirectory = f'teste/{contratoPuro}/{year}/{month}{year}/{filial}'
        funcSPdirectory = f'{contratoPuro}/{year}/{month}{year}/{filial}'
        outputPath = f'{outputPdfPath}/{contratoPuro}/{year}/{month}{year}/{filial}'
        funcSPfileName = f'CMAT_{folderNumber}_RE_{month}{year}.pdf'

        printPdfPages(pdfPath=pdfPath, outputPath=outputPath, fileName=funcSPfileName, pagesList=pagesListToPrintPdfRe)
 
        sharepointCreateFoldersAndUploadFile(directory=funcSPdirectory, fileDirectory=outputPath, fileName=funcSPfileName, unParam=unParam)

        logging.info(f'--RE process finished for {contratoPuro}--')
        
        logging.info(f'--RET process started for {contratoPuro}--')

        returnRetFunction = sharepointGetRetDocumentFile(inscricaoTomadorRe=inscricaoTomadorRe, sheetName=sheetName)
        
        if returnRetFunction == []:
            logging.warning(f'No RET files were found for "Inscrição do Tomador": {inscricaoTomadorRe}. Skipping register...')
            flagIncompleteProcessingTemp = True
            continue
        
        groupOfPagesRetFile, pdfPath, folderNumber = returnRetFunction

        #funcSPdirectory = f'teste/{contratoPuro}/{year}/{month}{year}/{filial}'
        funcSPdirectory = f'{contratoPuro}/{year}/{month}{year}/{filial}'
        outputPath = f'{outputPdfPath}/{contratoPuro}/{year}/{month}{year}/{filial}'
        funcSPfileName = f'CMAT_{folderNumber}_RET_{month}{year}.pdf'

        printPdfPages(pdfPath=pdfPath, outputPath=outputPath, fileName=funcSPfileName, pagesList=groupOfPagesRetFile)
 
        sharepointCreateFoldersAndUploadFile(directory=funcSPdirectory, fileDirectory=outputPath, fileName=funcSPfileName, unParam=unParam)

        logging.info(f'--RET process finished for {contratoPuro}--')

        logging.info(f'Finished full process for contract "{contratoPuro}".\n')

    return flagIncompleteProcessingPerm

        #deletePdfDirectories()
        #return










#----------------------------------------------------------------


def processEquipesDeMontagem():

    logging.info('Starting EQUIPES DE MONTAGEM process...')

    sheetName = sheet1Name

    fullListEdm = equipesDeMontagemSheetData()

    flagIncompleteProcessingTemp = False
    flagIncompleteProcessingPerm = False

    for index, lista in enumerate(fullListEdm):

        if flagIncompleteProcessingTemp == True:
            flagIncompleteProcessingPerm = True

        resultPagesRe = []
        resultPagesRet = []
        pdfPathRe = None
        pdfPathRet = None

        logging.info(f'lista: {lista}')

        key = list(lista.keys())[0]

        contratoPuro = key[0]
        unOriginal = key[1]       
        unParam = unConfig[unOriginal]
        namesExcel = list(lista.values())[0]

        logging.info(f'Currently working on {index+1}° row. (Contrato: {contratoPuro})')

        logging.info(f'contratoPuro: {contratoPuro}')
        logging.info(f'unOriginal: {unOriginal}')
        logging.info(f'unParam: {unParam}')
        logging.info(f'grupo names: {namesExcel}')

        for indexNameExcel, nameExcel in enumerate(namesExcel):

            if flagIncompleteProcessingTemp == True:
                flagIncompleteProcessingPerm = True

            logging.info(f'--RE process start for name {nameExcel}--')

            relFuncsFile = sharepointGetRelFuncsFile(filial=None, contratoPuro=f'{contratoPuro} - EQUIPE', edm=True)

            if relFuncsFile == None:
                logging.warning(f'Could not find Rel Funcs file for "{contratoPuro}" contract. Skipping register...')
                flagIncompleteProcessingTemp = True
                break

            namesList = extractRelFuncsPdfNames(pdfPath=relFuncsFile)

            if namesList == []:
                logging.warning(f'No names were found in "{relFuncsFile}" file. Skipping register...')
                flagIncompleteProcessingTemp = True
                break

            logging.info(f'namesList: {namesList}')

            flagNameNotFound = True
            for nameRelFuncs in namesList:
                if nameRelFuncs  in nameExcel:
                    flagNameNotFound = False
                    break
                
            if flagNameNotFound == True:
                logging.info(f'Name from Excel "{nameExcel}" not found in names from Rel Funcs "{namesList}". Skipping register...')
                flagIncompleteProcessingTemp = True
                break
            
            
            returnReFunction = sharepointGetReDocumentFile(personName=nameRelFuncs, sheetName=sheetName, documentType='RE')
            logging.info(f'returnList: {returnReFunction}')
            
            if returnReFunction == []:
                logging.warning(f'Retornou vazio para "{contratoPuro}". Skipping register...')
                flagIncompleteProcessingTemp = True
                break
        
        
            pageNumberFuncionario, pdfPath, folderNumber = returnReFunction
            
            logging.info(f'Page wich contains {nameExcel} - {pageNumberFuncionario}')
            logging.info(f'{type(pageNumberFuncionario)} - {pageNumberFuncionario}')
            if pdfPath == None:
                flagIncompleteProcessingTemp = True
                break
            
            pdfPathRe = pdfPath

            listPageNumberFuncionario = [pageNumberFuncionario]

            inscricoes = extractBothNumeroInscricaoFromRe(pageNumberFuncionario, pdfPath)
            if inscricoes == [] or len(inscricoes) != 2:
                logging.info(f'"inscricoes" not correct: "{inscricoes}". Skipping register...')
                flagIncompleteProcessingTemp = True
                continue
                #sys.exit()
            #inscricaoTomadorRe, inscricaoFilialRe = inscricoes
            inscricaoFilialRe, inscricaoTomadorRe = inscricoes
            logging.info(f'inscricaoFilialRe: {inscricaoFilialRe}')
            logging.info(f'inscricaoTomadorRe: {inscricaoTomadorRe}')

            pagsComNumeroDeInscricaoTomadorRe = pagesWithSameInscricaoTomadorRe(pdfPath=pdfPath, inscricao=inscricaoTomadorRe)
            logging.info(f'{len(pagsComNumeroDeInscricaoTomadorRe)} pages with same "Inscrição de Tomador"  number were found: {pagsComNumeroDeInscricaoTomadorRe}.')
            if pagsComNumeroDeInscricaoTomadorRe == []:
                logging.info('Skipping register...')
                flagIncompleteProcessingTemp = True
                continue


            #pagsComNumeroDeInscricaoFilialPlusResumoDoFechamentoEmpresaRe = pagesWithSameInscricaoFilialRe(pdfPath=pdfPath, inscricao=inscricaoTomadorRe)
            pagsComNumeroDeInscricaoFilialPlusResumoDoFechamentoEmpresaRe = newPagesWithSameInscricaoFilialRe(pdfPath=pdfPath, inscricaoFilial=inscricaoFilialRe)
            logging.info(f'{len(pagsComNumeroDeInscricaoFilialPlusResumoDoFechamentoEmpresaRe)} pages with same "Inscrição de Filial" + "Resumo do Fechamento - Empresa" number were found: {pagsComNumeroDeInscricaoFilialPlusResumoDoFechamentoEmpresaRe}')
            if pagsComNumeroDeInscricaoFilialPlusResumoDoFechamentoEmpresaRe == []:
                logging.info('Skipping register...')
                flagIncompleteProcessingTemp = True
                continue
            
            # Converter as listas para conjuntos
            conjunct1 = set(listPageNumberFuncionario)
            conjunct2 = set(pagsComNumeroDeInscricaoTomadorRe)
            conjunct3 = set(pagsComNumeroDeInscricaoFilialPlusResumoDoFechamentoEmpresaRe)

            # Unir os conjuntos sem repetir valores
            unitedConjunct = conjunct1.union(conjunct2, conjunct3)
            #unitedConjunct = conjunct2.union(conjunct3)
            # Converter de volta para uma lista
            pagesListToPrintPdfRe = list(unitedConjunct)
            pagesListToPrintPdfRe.sort()
            #pagesListToPrintPdfRe.remove(pageNumberFuncionario)
            #pagesListToPrintPdfRe.insert(0, pageNumberFuncionario)

            resultPagesRe.append(pagesListToPrintPdfRe)

            '''
            logging.info(f'RE - pagesListToPrintPdfSortedRe: {pagesListToPrintPdfRe}')

            if os.path.exists(f'{outputPdfPath}/{contratoPuro} - EQUIPE/{year}/{month}{year}') == False:
                os.makedirs(f'{outputPdfPath}/{contratoPuro} - EQUIPE/{year}/{month}{year}')           

            funcSPdirectory = f'{contratoPuro} - EQUIPE/{year}/{month}{year}'
            outputPath = f'{outputPdfPath} - EQUIPE/{contratoPuro}/{year}/{month}{year}'
            funcSPfileName = f'CMAT_{folderNumber}_RE_{month}{year}.pdf'

            printPdfPages(pdfPath=pdfPath, outputPath=outputPath, fileName=funcSPfileName, pagesList=pagesListToPrintPdfRe)
            '''

            #sharepointCreateFoldersAndUploadFile(directory=funcSPdirectory, fileDirectory=outputPath, fileName=funcSPfileName)

            logging.info(f'RE - groupOfPages: {pagesListToPrintPdfRe}')

            logging.info(f'--RE process finished for name {nameExcel}--')
        
            logging.info(f'--RET process started for name {nameExcel}--')

            returnRetFunction = sharepointGetRetDocumentFile(inscricaoTomadorRe=inscricaoTomadorRe, sheetName=sheetName)
            
            if returnRetFunction == []:
                logging.warning(f'No RET files were found for "Inscrição do Tomador": {inscricaoTomadorRe}. Skipping register...')
                flagIncompleteProcessingTemp = True
                break
            
            
            groupOfPagesRetFile, pdfPath, folderNumber = returnRetFunction

            pdfPathRet = pdfPath

            resultPagesRet.append(groupOfPagesRetFile)

            '''
            funcSPdirectory = f'{contratoPuro} - EQUIPE/{year}/{month}{year}'
            outputPath = f'{outputPdfPath} - EQUIPE/{contratoPuro}/{year}/{month}{year}'
            funcSPfileName = f'CMAT_{folderNumber}_RET_{month}{year}.pdf'

            printPdfPages(pdfPath=pdfPath, outputPath=outputPath, fileName=funcSPfileName, pagesList=groupOfPagesRetFile)
 
            #sharepointCreateFoldersAndUploadFile(directory=funcSPdirectory, fileDirectory=outputPath, fileName=funcSPfileName)
            
            '''

            logging.info(f'RET - groupOfPages: {groupOfPagesRetFile}')

            logging.info(f'RET process finished for name {nameExcel}')

            logging.info(f'Finished process for name: {nameExcel}. Name {indexNameExcel+1} / {len(namesExcel)}.')


        logging.info(f'RE process started for contract {contratoPuro}')

        logging.info(f'RE - resultPagesRe (com duplicatas): {resultPagesRe}')

        if resultPagesRe == []:
            logging.warning('RE - "resultPagesRe" is empty. Skipping to next contract...')
            flagIncompleteProcessingTemp = True
            continue
        
        resultPagesRe = sum(resultPagesRe, [])
        resultPagesRe = list(set(resultPagesRe))
        resultPagesRe.sort()

        logging.info(f'RE - resultPagesRe (sem duplicatas): {resultPagesRe}')

        if os.path.exists(f'{outputPdfPath}/{contratoPuro} - EQUIPE/{year}/{month}{year}') == False:
            os.makedirs(f'{outputPdfPath}/{contratoPuro} - EQUIPE/{year}/{month}{year}')           

        #funcSPdirectory = f'teste/{contratoPuro} - EQUIPE/{year}/{month}{year}'
        funcSPdirectory = f'{contratoPuro} - EQUIPE/{year}/{month}{year}'
        outputPath = f'{outputPdfPath}/{contratoPuro} - EQUIPE/{year}/{month}{year}'
        funcSPfileName = f'CMAT_{folderNumber}_RE_{month}{year}.pdf'

        printPdfPages(pdfPath=pdfPathRe, outputPath=outputPath, fileName=funcSPfileName, pagesList=resultPagesRe)
        
        sharepointCreateFoldersAndUploadFile(directory=funcSPdirectory, fileDirectory=outputPath, fileName=funcSPfileName, unParam=unParam)

        logging.info(f'--RE process finished for contract {contratoPuro}--')



        logging.info(f'--RET process started for contract {contratoPuro}')

        logging.info(f'RET - resultPagesRet (com duplicatas): {resultPagesRet}')


        if resultPagesRet == []:
            logging.warning('RET - "resultPagesRet" is empty. Skipping to next contract...')
            flagIncompleteProcessingTemp = True
            continue

        resultPagesRet = sum(resultPagesRet, [])
        resultPagesRet = list(set(resultPagesRet))
        resultPagesRet.sort()

        logging.info(f'RET - resultPagesRet (sem duplicatas): {resultPagesRet}')

        
        #funcSPdirectory = f'teste/{contratoPuro} - EQUIPE/{year}/{month}{year}'
        funcSPdirectory = f'{contratoPuro} - EQUIPE/{year}/{month}{year}'
        outputPath = f'{outputPdfPath}/{contratoPuro} - EQUIPE/{year}/{month}{year}'
        funcSPfileName = f'CMAT_{folderNumber}_RET_{month}{year}.pdf'

        printPdfPages(pdfPath=pdfPathRet, outputPath=outputPath, fileName=funcSPfileName, pagesList=resultPagesRet)

        sharepointCreateFoldersAndUploadFile(directory=funcSPdirectory, fileDirectory=outputPath, fileName=funcSPfileName, unParam=unParam)
        
        logging.info(f'RET process finished for contract {contratoPuro}')

        logging.info(f'Finished full process for names {namesExcel} in contract "{contratoPuro}".\n')

        #deletePdfDirectories()
    return flagIncompleteProcessingPerm