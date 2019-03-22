# -*- coding: UTF-8 -*-
import logging

_LOG_FILE_NAME = 'default.log'
_LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
_DATE_FORMAT = "%m/%d/%Y %H:%M:%S %p"
LOG = logging
LOG.basicConfig(filename=_LOG_FILE_NAME, level=logging.INFO, format=_LOG_FORMAT, datefmt=_DATE_FORMAT)
__ALL__ = ['LOG']
