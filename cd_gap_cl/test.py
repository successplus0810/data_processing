import json
import pandas as pd
import snowflake.connector as sf
import os
import xlwings as xw
from xlwings.constants import DeleteShiftDirection
import datetime
import math

dictt = {(1,2):[3,4]}
print(dictt)

for key,value in dictt.items():
    a,b = key
    c,d = value
    print(a,b,c,d)

