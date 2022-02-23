
import pandas as pd
import numpy as np
import sys, getopt, os
import logging
import test_called_script


base = "base-la.csv"

#**********************************************
#    M A I N   P R O G R A M   S T A R T      *
#**********************************************
#
def main():
    global base
    
    
    logging.basicConfig(format='%(asctime)s %(levelname)-8s %(message)s',
            level=logging.INFO,
            datefmt='%Y-%m-%d %H:%M:%S')

adtable = ["AD2", "AD3", "AD4"]

for ad in adtable:
    test_called_script("echo hello" , ad )

    

if __name__ == '__main__':
    
    main()
    