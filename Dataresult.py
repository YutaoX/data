#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Sep 25 23:29:24 2019

@author: yutao
"""

import os, sys
import shutil  
# Python清空指定文件夹下所有文件的方法shutil
# 创建的目录
#from DataGetFunc import Get_future_data_day
#from DataAnsysFunc import Data_ansys
###
from DataFunc import *

######################################
future_code="SR0"
path = "../data"
ret = os.access( path, os.F_OK)
if ret:
    shutil.rmtree( path)
#    os.mkdir( path)  
os.makedirs( path, 0o777 );
os.chdir( path )
#    retval = os.getcwd()
Datacode(future_code)
    