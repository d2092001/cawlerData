# -*- coding: utf-8 -*-
"""
Created on Thu Nov 24 14:11:18 2022

@author: NM DUC
"""

import numpy as np
from selenium import webdriver
from time import sleep
import random
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException
from selenium.webdriver.common.by import By
import pandas as pd
import re
import openpyxl

import ctypes  # An included library with Python install.


# Declare browser
driver = webdriver.Chrome(executable_path='./chromedriver.exe')
driver.get('http://helpocm.mshopkeeper.vn/')