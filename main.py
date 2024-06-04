#!/usr/bin/python

from src.utils import *
import sys
import os
import logging

act_dir = os.path.dirname(os.path.abspath(__file__))
proyect_dir_src = os.path.join(act_dir, 'src')
sys.path.append(proyect_dir_src)


if __name__ == "__main__":

    saveattachemnts()
    load()
