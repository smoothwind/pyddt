# -*- coding: UTF-8 -*-
import logging

LOG = logging

file_op = open("default.log", encoding="utf-8", mode="a+")

LOG.basicConfig(filename="default.log", level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s",
                datefmt="%m/%d/%Y %H:%M:%S %p")

__ALL__ = ['LOG']
