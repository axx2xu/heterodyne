from pymeasure.adapters import VISAAdapter
import time
import pyvisa

################################################################################################################################################################################
#                               ****THIS IS A TEST CODE TO TEST THE CONNECTION TO THE WAVELENGTH METER, successfully reads freq and WL for one time****

################################################################################################################################################################################

# Initialize the VISA resource manager
rm = pyvisa.ResourceManager()

# Open a connection to the wavelength meter (update with your actual GPIB address)
wavelength_meter = rm.open_resource('GPIB0::20::INSTR')

# Set timeout to 10 seconds (10000 milliseconds)
wavelength_meter.timeout = 10000  # in milliseconds

# Test communication by sending an *IDN? command
try:
    idn = wavelength_meter.query("*IDN?")
    print("Instrument ID:", idn)
except Exception as e:
    print("Error reading instrument ID:", e)

# Initiate a measurement
try:
    wavelength_meter.write(":INITiate:SCALar:POWer")
    time.sleep(1)  # Wait for the measurement to complete
except Exception as e:
    print("Error initiating measurement:", e)

# Fetch frequency
try:
    frequency = wavelength_meter.query(":FETCH:SCALar:POWer:FREQuency?")
    print("Frequency:", frequency)
except Exception as e:
    print("Error reading frequency:", e)

# Fetch wavelength
try:
    wavelength = wavelength_meter.query(":FETCH:SCALar:POWer:WAVelength?")
    print("Wavelength:", wavelength)
except Exception as e:
    print("Error reading wavelength:", e)

# Close the connection
wavelength_meter.close()


# # Initialize the VISA resource manager
# rm = pyvisa.ResourceManager()

# # List all connected VISA devices (Optional: To verify connections)
# print("Connected devices:", rm.list_resources())

# # Create a VISA adapter for the Wavelength Meter
# wl_meter_adapter = rm.open_resource('GPIB0::20::INSTR')  # Update with your actual GPIB address

# time.sleep(1) # Sleep for 1 seconds

# try:
#     frequency = wl_meter_adapter.query(":MEASure:ARRay:POWer:FREQuency?")
#     print("Frequency:", frequency)
# except Exception as e:
#     print("Error reading frequency:", e)

# try:
#     wavelength = wl_meter_adapter.query(":MEASure:ARRay:POWer:WAVelength?")
#     print("Wavelength:", wavelength)
# except Exception as e:
#     print("Error reading wavelength:", e)


#  #MEASure:ARRay:POWer:WAVelength?        query wavelength measurement command
#  #MEASure:ARRay:POWer:FREQuency?            query frequency measurement command