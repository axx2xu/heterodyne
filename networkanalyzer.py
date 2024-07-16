import skrf as rf
import numpy as np
from scipy.interpolate import interp1d

# Function to read the header of the s2p file and detect units
def read_s2p_header(filepath):
    with open(filepath, 'r') as file:
        lines = file.readlines()
    
    freq_unit = 'Hz'
    data_format = 'dB'
    
    for line in lines:
        if line.startswith('#'):
            parts = line.split()
            for i, part in enumerate(parts):
                if part.lower() in ['hz', 'khz', 'mhz', 'ghz']:
                    freq_unit = part.lower()
                elif part.lower() in ['ri', 'db', 'ma']:
                    data_format = part.upper()
            break
    
    return freq_unit, data_format

# Path to the .s2p file
filepath = 'C:/Users/Tommy/Downloads/DD_00_08.s2p'

# Read the header to detect the units
freq_unit, data_format = read_s2p_header(filepath)

print(freq_unit, data_format)