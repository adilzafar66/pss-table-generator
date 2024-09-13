#***********************
#
# Copyright (c) 2021-2023, Operation Technology, Inc.
#
# THIS PROGRAM IS CONFIDENTIAL AND PROPRIETARY TO OPERATION TECHNOLOGY, INC. 
# ANY USE OF THIS PROGRAM IS SUBJECT TO THE PROGRAM SOFTWARE LICENSE AGREEMENT, 
# EXCEPT THAT THE USER MAY MODIFY THE PROGRAM FOR ITS OWN USE. 
# HOWEVER, THE PROGRAM MAY NOT BE REPRODUCED, PUBLISHED, OR DISCLOSED TO OTHERS 
# WITHOUT THE PRIOR WRITTEN CONSENT OF OPERATION TECHNOLOGY, INC.
#
#***********************


import sys
from scipy import signal
import json
import numpy as np

def butterworth_calc():

    sampling_frequency = float(sys.argv[1])

    K = float(sys.argv[2])
    lambd = float(sys.argv[3])
    omega1 = float(sys.argv[4])
    omega2 = float(sys.argv[5])
    omega3 = float(sys.argv[6])
    omega4 = float(sys.argv[7])

    num1 = np.array([K * omega1, 0])
    den1 = np.array([1, 2 * lambd, omega1*omega1])
    num2 = np.array([1 / omega2, 1])
    den2 = np.array([1 / (omega3 * omega4), 1 / omega3 + 1 / omega4, 1])
    bilinear_num, bilinear_den = signal.bilinear(np.convolve(num1, num2), np.convolve(den1, den2), sampling_frequency)
    sos_bilinear = signal.tf2sos(bilinear_num, bilinear_den)       

    binear_filter_coefficients = {}
    binear_filter_coefficients['Bilinear'] = str(sos_bilinear.tolist())
    json_data = json.dumps(binear_filter_coefficients)
    
    print(json_data)
    return(json_data)

if __name__ == '__main__':
    result = butterworth_calc()
    