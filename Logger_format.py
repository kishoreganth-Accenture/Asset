import logging

def config_log(name,log_file = "IDS Auto.log"):

    formatter =  logging.Formatter('%(asctime)s-%(levelname)s : %(message)s')
    handler = logging.FileHandler(log_file)
    handler.setFormatter(formatter)
    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)
    logger.addHandler(handler)

    return logger




