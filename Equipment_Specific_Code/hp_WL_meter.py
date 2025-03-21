import pyvisa
import time

rm = pyvisa.ResourceManager()
wavelength_meter = rm.open_resource('GPIB0::20::INSTR')
wavelength_meter.timeout = 5000

wavelength_meter.write(":CALCulate3:PRESet")
time.sleep(1)
wavelength_meter.write("*OPC?")
if wavelength_meter.read().strip() == "1":
    print("CALCulate3 states cleared.")
else:
    print("Warning: Operation did not complete as expected.")

wavelength_meter.write(":CALCulate3:DELTa:REFerence:WAVelength MIN")
time.sleep(1)

wavelength_meter.write(":CALCulate3:DELTa:WAVelength ON")
time.sleep(1)
wavelength_meter.write("*OPC?")
if wavelength_meter.read().strip() == "1":
    print("Delta wavelength mode successfully turned ON.")
else:
    print("Delta wavelength mode command did not complete as expected.")


wavelength_meter.close()
rm.close()


# from pymeasure.adapters import VISAAdapter
# import time
# import pyvisa

# ################################################################################################################################################################################
# #                               ****THIS IS A TEST CODE TO TEST THE CONNECTION TO THE WAVELENGTH METER, successfully reads freq and WL for one time****

# ################################################################################################################################################################################

# # Initialize the VISA resource manager
# rm = pyvisa.ResourceManager()

# # Open a connection to the wavelength meter (update with your actual GPIB address)
# wavelength_meter = rm.open_resource('GPIB0::20::INSTR')

# # Send the command to turn delta wavelength mode ON
# wavelength_meter.write(":CALCulate3:DELTa:WAVelength OFF")
# time.sleep(1)


# # # Set timeout to 10 seconds (10000 milliseconds)
# # wavelength_meter.timeout = 10000  # in milliseconds

# # # Test communication by sending an *IDN? command
# # try:
# #     idn = wavelength_meter.query("*IDN?")
# #     print("Instrument ID:", idn)
# # except Exception as e:
# #     print("Error reading instrument ID:", e)

# # # Initiate a measurement
# # try:
# #     wavelength_meter.write(":INITiate:SCALar:POWer")
# #     time.sleep(1)  # Wait for the measurement to complete
# # except Exception as e:
# #     print("Error initiating measurement:", e)

# # # Fetch frequency
# # try:
# #     frequency = wavelength_meter.query(":FETCH:SCALar:POWer:FREQuency?")
# #     print("Frequency:", frequency)
# # except Exception as e:
# #     print("Error reading frequency:", e)

# # # Fetch wavelength
# # try:
# #     wavelength = wavelength_meter.query(":FETCH:SCALar:POWer:WAVelength?")
# #     print("Wavelength:", wavelength)
# # except Exception as e:
# #     print("Error reading wavelength:", e)

# # Close the connection
# wavelength_meter.close()


# # # Initialize the VISA resource manager
# # rm = pyvisa.ResourceManager()

# # # List all connected VISA devices (Optional: To verify connections)
# # print("Connected devices:", rm.list_resources())

# # # Create a VISA adapter for the Wavelength Meter
# # wl_meter_adapter = rm.open_resource('GPIB0::20::INSTR')  # Update with your actual GPIB address

# # time.sleep(1) # Sleep for 1 seconds

# # try:
# #     frequency = wl_meter_adapter.query(":MEASure:ARRay:POWer:FREQuency?")
# #     print("Frequency:", frequency)
# # except Exception as e:
# #     print("Error reading frequency:", e)

# # try:
# #     wavelength = wl_meter_adapter.query(":MEASure:ARRay:POWer:WAVelength?")
# #     print("Wavelength:", wavelength)
# # except Exception as e:
# #     print("Error reading wavelength:", e)


# #  #MEASure:ARRay:POWer:WAVelength?        query wavelength measurement command
# #  #MEASure:ARRay:POWer:FREQuency?            query frequency measurement command