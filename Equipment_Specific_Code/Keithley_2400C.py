import time
import pyvisa

# Initialize the VISA resource manager
rm = pyvisa.ResourceManager()

# List all connected VISA devices (Optional: To verify connections)
print("Connected devices:", rm.list_resources())

# # Create a VISA adapter for the Keithley 2400
# keithley = rm.open_resource('GPIB0::24::INSTR')  # Update with your actual GPIB address

# # Allow some time for the measurement
# time.sleep(1)

# # Measure the current
# response = keithley.query(":MEASure:CURRent?")
# current_values = response.split(',')

# if len(current_values) > 1:
#     current = current_values[1]
#     print("Measured Current:", current)
# else:
#     print("Unexpected response format:", response)

# # Set the Keithley to local mode
# keithley.write(":SYSTem:LOCal")

# # Close the connection
# keithley.close()
