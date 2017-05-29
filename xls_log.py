import logging



def init_xls_logger():
        logger=logging.getLogger('xlslogger')
        logger.setLevel(logging.DEBUG)
    
        #create a handler to write the log to file 
        fh=logging.FileHandler('test1.log')
        fh.setLevel(logging.DEBUG)

        #creat a handler to output to console
        ch=logging.StreamHandler()
        ch.setLevel(logging.INFO)

        #define the format of output
        formatter_fh=logging.Formatter('%(asctime)s - %(levelname)s - %(funcName)s - %(message)s')                                
        formatter_ch=logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        fh.setFormatter(formatter_fh)
        ch.setFormatter(formatter_ch)

        #assgign the handler to logger
        logger.addHandler(fh)
        logger.addHandler(ch)

        return logger
 
def get_xls_logger():
    return logging.getLogger('xlslogger')
