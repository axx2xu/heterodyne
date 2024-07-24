# Include all necessary imports from the equipment and the pymeasure library. The pymeasure library (GitHub linked https://github.com/pymeasure/pymeasure/tree/master) includes drivers
# for various instruments and provides a simple interface for controlling them. Pyvisa is a Python wrapper for the VISA library, which is used to communicate with instruments over GPIB, USB, and other interfaces. 

# import the necessary libraries
from pymeasure.instruments.keithley import Keithley2400
from pymeasure.adapters import VISAAdapter
import time

# Initialize the VISA resource manager
import pyvisa
rm = pyvisa.ResourceManager()
################################################################################################################################################################################
#                               ****MAKE SURE TO ALWAYS CHECK THE CONNECTED DEVICE PORTS BEFORE RUNNING THE CODE****

# Port addresses may change depending on the connected devices, if anything has been moved around, if the computer has been reset, etc. Make sure to always check, in case running the 
# code for the wrong instrument could cause damage. 
################################################################################################################################################################################

# To verify the connected devices, run the following command:
print("Connected devices:", rm.list_resources())

# Connect the instruments to the correct I/O Ports using adapters. To find the correct ports there is necessary NI MAX software installed on the computer. After opening NI MAX, 
# the GPIB address of connected instruments will be shown under the "Devices and Interfaces" tab. The GPIB address is used to connect the instruments to the computer.

# Create a VISA adapter for the Keithley 2400
keithley_adapter = VISAAdapter('GPIB0::24::INSTR')  # Update with your actual GPIB address

# Create a VISA adapter for the ECL Laser
ecl_adapter = VISAAdapter('GPIB0::1::INSTR')  # Update with your actual GPIB address

# Create a VISA adapter for the ESA Spectrum Analyzer
esa_adapter = VISAAdapter('GPIB0::2::INSTR')  # Update with your actual GPIB address

# Create a VISA adapter for the RSL Power Meter
rsl_adapter = VISAAdapter('GPIB0::3::INSTR')  # Update with your actual GPIB address

# Create a VISA adapter for the Optical Attentuator
attenuator_adapter = VISAAdapter('GPIB0::4::INSTR')  # Update with your actual GPIB address

# Create a VISA adapter for the optical wavelength meter
wavelength_adapter = VISAAdapter('GPIB0::5::INSTR')  # Update with your actual GPIB address

################################################################################################################################################################################
#                                                               ****USER INPUTS****
################################################################################################################################################################################
while True:
    print("Would you like to manually set each parameter (M) or input a parameter list (L)?")
    user_input = input("Enter 'M' for manual input or 'L' for list input: ").strip().upper()
    if user_input == 'M':
        # Manual input of parameters

        # **** these need to be updated with the correct inputs, not sure if these are correct for the instruments being used ****

        print("Please enter the parameters for the measurement:")
        laser3_WL = float(input("Enter the starting WL for laser 3 (nm): "))
        laser3_step = float(input("Enter the WL step size for laser 3 (nm): "))
        final_freq = float(input("Enter the final frequency (GHz): "))
        freq_threshold = float(input("Enter frequency threshold (GHz): "))
        operating_current = float(input("Enter the operating current (mA): "))
        current_threshold = float(input("Enter the current threshold (mA): "))
        power_measurement_wait = float(input("Enter the power measurement wait time (ms): "))
        sample_num = float(input("Enter the number of samples for averaging: "))
        #VFI PM5 Range from Spencer Labview?
        break

    elif user_input == 'L': # Input a parameter list to repeat settings, instead of manually inputting each time - separate by comma
        # Input a parameter list
        parameter_list = input("Enter the parameter list in the following format: laser3_WL,laser3_step,final_freq,freq_threshold,operating_current,current_threshold,power_measurement_wait,sample_num: ")
        try:
            laser3_WL, laser3_step, final_freq, freq_threshold, operating_current, current_threshold, power_measurement_wait, sample_num = map(float, parameter_list.split(","))
            break
        except ValueError:
            print("Invalid input format. Please enter the parameters correctly.")
    else:
        print("Invalid input. Please enter 'M' or 'L'.")


################################################################################################################################################################################
#                                                               ****SET UP THE ANRITSU OSICS ECL****
################################################################################################################################################################################
# Initialize the ECL Laser

# Set the initial wavelength for lasers 3 and 4
# Set the step size for laser 3