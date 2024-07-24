import pyvisa
import time
import math

################################################################################################################################################################################
#                               ****MAKE SURE TO ALWAYS CHECK THE CONNECTED DEVICE PORTS BEFORE RUNNING THE CODE****

# This works to the extent that it sets some settings and returns a value, has not been verified
# https://docs.alltest.net/manual/Alltest-Rohde-and-Schwarz-NRP-Z58-UserManual.pdf NRZ-Z58 Manual for SCPI commands
# The NRP-NI-VISA-PASSPORT driver needs to be installed from R&S, the NRP Toolkit needs to be installed from R&S
# This device is not detected by NI MAX but through the pyvisa list resources command
 
################################################################################################################################################################################

def main():
    # Initialize the VISA resource manager
    rm = pyvisa.ResourceManager()

    # List all connected instruments
    instruments = rm.list_resources()
    print("Connected instruments:", instruments)

    # Use the correct VISA address for your RS NRP-Z58 sensor
    instrument_address = 'RSNRP::0x00a8::100940::INSTR'
    
    try:
        # Open a connection to the sensor
        sensor = rm.open_resource(instrument_address)
        
        # Identify the instrument
        #print("Querying instrument identification...")
        #response = sensor.query('*IDN?')
        #print("Sensor response:", response)
        
        # Reset the instrument
        #print("Resetting the instrument...")
        #sensor.write('*RST')

        # Calibrate the sensor (uncomment if needed)
        # print("Calibrating the sensor...")
        # sensor.write('CAL:ZERO:AUTO ONCE')
        # time.sleep(10) # wait 10 seconds

        # Configure the measurement settings
        print("Configuring measurement settings...")
        sensor.write('INIT:CONT OFF')
        sensor.write('SENS:FUNC "POW:AVG"')
        sensor.write('SENS:FREQ 1e9')
        sensor.write('SENS:AVER:COUN:AUTO ON')
        sensor.write('SENS:AVER:COUN 16')
        sensor.write('SENS:AVER:STAT ON')
        sensor.write('SENS:AVER:TCON REP')
        sensor.write('SENS:POW:AVG:APER 5e-3')
        sensor.write('INIT:IMM')
        
        # Query immediate trigger
        print("Triggering and reading data...")
        output = sensor.query('TRIG:IMM')
        output = output.split(',')[0]
        output_dbm = math.log10(float(output)) * 10 + 30
        print("Output:", output_dbm, "dBm")

        # Close the connection
        sensor.close()
        print("Connection closed.")
       
    except pyvisa.errors.VisaIOError as e:
        print(f"A VISA error occurred: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    main()
