import logging

def config_log(log_file = "kishore"):
    logging.basicConfig(filename= log_file +'.log',format = '%(asctime)s-%(levelname)s : %(message)s',
                                filemode = 'w')
    logger = logging.getLogger()

    logger.setLevel(logging.DEBUG)
    return logger




