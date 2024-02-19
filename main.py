from subprograms.process import *

def main():

    #processCmatEngenharia()

    logging.info(f'----- Starting CMAT ENGENHARIA process... -----')
    processCmat(servicos=False)
    logging.info(f'----- Finished CMAT ENGENHARIA process. -----')

    logging.info(f'----- Starting CMAT SERVIÇOS process... -----')
    processCmat(servicos=True)
    logging.info(f'----- Finished CMAT SERVIÇOS process. -----')

    logging.info(f'----- Starting EQUIPES DE MONTAGEM process... -----')
    processEquipesDeMontagem()
    logging.info(f'----- Finished EQUIPES DE MONTAGEM process. -----')
    

if __name__ == '__main__':
    sys.excepthook = show_exception_and_exit
    sharepointGetExcelFile()
    main()
    logProcessingDuration(startTimer=startTimer)
    sharepointSendLog()
    emailSuccess()
    sys.exit(0)