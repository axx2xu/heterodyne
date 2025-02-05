import pyvisa
import time

rm = pyvisa.ResourceManager()

# List all connected VISA devices (Optional: To verify connections)
print("Connected devices:", rm.list_resources())
voa_GPIB = 'GPIB0::26::INSTR'  # Update with your actual GPIB address
voa = rm.open_resource(voa_GPIB)

time.sleep(1)

p_actual = voa.query('READ:POW?')
p_actual = float(p_actual) # convert to float so it can be rounded
p_actual = round(p_actual,3) # round 3 decimal places
print(f"Power out {p_actual}")

voa.write('INITiate2:CHANnel1:IMMediate')
voa.write('INPUT2:[CHANnel1]:WAVelength1540')
voa.write('SYST:LOC')  # Put the VOA back into local mode so the attenuation can be adjusted



