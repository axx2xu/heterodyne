import skrf as rf
import numpy as np

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
filepath = 'C:/Users/Tommy/Downloads/1_2_1mm_cable_BW_20240306_ft4xx.s2p'

# Read the header to detect the units
freq_unit, data_format = read_s2p_header(filepath)
print(f"Frequency Unit: {freq_unit}, Data Format: {data_format}")

# Load the .s2p file
s2p_file = rf.Network(filepath)

# Verify the Network object
print(f"Network frequency unit: {s2p_file.frequency.unit}")
print(f"Number of frequency points: {len(s2p_file.f)}")

# Print the scattering parameter matrix shape
print(f"Scattering parameter matrix shape: {s2p_file.s.shape}")

# Gather s12 and s21 data
s12 = s2p_file.s_db[:,0,1]
s21 = s2p_file.s_db[:,1,0]
s_avg = (s12 + s21) / 2

# Verify the extracted data
print("s12:", s12[:10])  # Print first 10 values for inspection
print("s21:", s21[:10])  # Print first 10 values for inspection
print("s_avg:", s_avg[:10])  # Print first 10 values for inspection
