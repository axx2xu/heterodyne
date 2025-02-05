import pyvisa
import time

# Initialize the instrument using the VISA resource string
rm = pyvisa.ResourceManager()
instrument = rm.open_resource('GPIB0::25::INSTR')

# Set the timeout to 10 seconds
instrument.timeout = 10000  # Set timeout to 10 seconds (optional, but useful for longer operations)

# Test communication with the IDN query
print("Device IDN:", instrument.query('*IDN?'))

# Set auto-ranging for the power meter in slot 2, channel 1
# instrument.write(':SENSe2:CHAN1:POWer:RANGe:AUTO 1')

# # Set the power unit to dBm
# instrument.write(':SENSe2:CHAN1:POWer:UNIT DBM')

# Wait for a second to allow settings to take effect
time.sleep(1)

# Initialize the measurement
# 

# Wait for 1 second to allow the measurement to complete
time.sleep(1)

# Loop to fetch 10 measurements
for x in range(10):
    # Fetch the power reading in dBm
    instrument.write('INITiate2:CHANnel1:IMMediate')
    power_reading = instrument.query(':READ2:CHAN1:POWer?')
    
    # Print the result in dBm
    print(f'Optical Power: {power_reading} dBm')
    
    # Wait for 1 second between measurements
    time.sleep(1)
