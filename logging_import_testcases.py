import logging

logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s %(levelname)-8s [%(filename)s:%(lineno)d]: %(message)s',
                    datefmt='%d-%m-%Y %H:%M:%S',
                    filename='import_testcases.log')
logger = logging.getLogger('import_testcases_logger')

a = 1
b = "test"
c = 7.98
logger.debug('test debug')
str_debug = f"a= {a} b= {b} c= {c:.2f}"
logger.debug(str_debug)
str_debug = f"a= {a:.<20} b= {b:.<20} c= {c:>7.2f}"
logger.debug(str_debug)
