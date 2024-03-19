from subprograms.process import *

def main():

    logging.info(f'----- Starting CMAT ENGENHARIA process... -----')
    flagIncompleteProcessingPermCmatEng = processCmat(servicos=False)
    logging.warning(f'flagIncompleteProcessingPermCmatEng: {flagIncompleteProcessingPermCmatEng}')
    logging.info(f'----- Finished CMAT ENGENHARIA process. -----')

    logging.info(f'----- Starting CMAT SERVIÇOS process... -----')
    flagIncompleteProcessingPermCmatServ = processCmat(servicos=True)
    logging.warning(f'flagIncompleteProcessingPermCmatServ: {flagIncompleteProcessingPermCmatServ}')
    logging.info(f'----- Finished CMAT SERVIÇOS process. -----')

    logging.info(f'----- Starting EQUIPES DE MONTAGEM process... -----')
    flagIncompleteProcessingPermEdm = processEquipesDeMontagem()
    logging.warning(f'flagIncompleteProcessingPermCmatEdm: {flagIncompleteProcessingPermEdm}')
    logging.info(f'----- Finished EQUIPES DE MONTAGEM process. -----')

    if flagIncompleteProcessingPermCmatEng == True or flagIncompleteProcessingPermCmatServ == True or flagIncompleteProcessingPermEdm == True:
        return True
    else:
        return False

if __name__ == '__main__':
    sys.excepthook = show_exception_and_exit
    sharepointGetExcelFile()
    flagIncompleteProcessing = main()
    logProcessingDuration(startTimer=startTimer)
    sharepointSendLog()
    if flagIncompleteProcessing == False:
        emailSuccess()
    else:
        emailIncompleteProcessing()
    sys.exit(0)