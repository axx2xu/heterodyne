from pymeasure.adapters import VISAAdapter
import time
import pyvisa

################################################################################################################################################################################
#                               ****THIS IS A TEST CODE TO TEST THE CONNECTION TO THE ECL MODULE, IT FUNCTIONS IN WRITING TO AND ENABLING LASER 3 and 4****

################################################################################################################################################################################


# Initialize the VISA resource manager
rm = pyvisa.ResourceManager()

# List all connected VISA devices (Optional: To verify connections)
print("Connected devices:", rm.list_resources())

# Create a VISA adapter for the ECL laser
ecl_adapter = rm.open_resource('GPIB0::10::INSTR')  # Update with your actual GPIB address

# Set timeout to 10 seconds (10000 milliseconds)
ecl_adapter.timeout = 10000  # in milliseconds

print("Please enter the parameters for the measurement:")
laser_3_WL = float(input("Enter the starting WL for laser 3 (nm): "))
laser_4_WL = float(input("Enter the starting WL for laser 4 (nm): "))
laser_4_step = float(input("Enter the frequency step size for laser 4 (GHz): "))
final_freq = float(input("Enter the final frequency (GHz): "))

# Set the laser 3 wavelength
print("Setting laser wavelength to 1540.00 nm...")
ecl_adapter.write("CH3:L={laser_3_WL:.2f}")

# Set the laser 4 initial wavelength to be below the laser 3 wavelength, laser 4 will then be stepeped up by 1GHz until reaching and passing laser 3 wavelength
print("Setting laser wavelength to 1539.990 nm...")
ecl_adapter.write("CH4:L={laser_4_WL:.3f}")

# Set the laser output power in dBm
print("Setting laser output power to 6.00 dBm...")
ecl_adapter.write("CH3:P=06.00") # Set laser 3 output power to 6.00 dBm
ecl_adapter.write("CH4:P=06.00") # Set laser 4 output power to 6.00 dBm

# Enable the laser
print("Enabling laser...")
ecl_adapter.write("CH3:ENABLE")
ecl_adapter.write("CH4:ENABLE")

# Wait for the laser to stabilize
time.sleep(1)

c = 299792458 # Speed of light in m/s

# Loop through amount of steps needed to reach final frequency for laser 4, 1 second delay before enabling and disabling and changing wavelength
for x in range (final_freq/laser_4_step):
    # Give laser 4 new wavelength based off increase in frequency by inputted step size

    laser_4_freq = c/laser_4_WL # Calculate current frequency of laser 4
    laser_4_new_freq = laser_4_freq - (laser_4_step*10**9) # Subtract the step size of current frequency to get new frequency (subtract frequency to increase wavelength)
    laser_4_WL = c/laser_4_new_freq # Set the new wavelength by dividing speed of light by new frequency

    ecl_adapter.write("CH4:DISABLE") # Disable laser 4

    time.sleep(1) # Wait 1 second before changing wavelength
    ecl_adapter.write("CH4:L={laser_4_WL:.3f}") # Set new wavelength for laser 4
    ecl_adapter.write("CH4:ENABLE") # Enable laser 4 with new wavelength
    time.sleep(1)
 

# Measure the laser wavelength and output power
print("Querying laser wavelength...")
wavelength = ecl_adapter.query("CH3:L?")
print(f"Laser 3 Wavelength: {wavelength}")

print("Querying laser output power...")
output_power = ecl_adapter.query("CH3:MW?")
print(f"Laser 3 Output Power: {output_power}")

time.sleep(5)

# Disable the laser
print("Disabling laser...")
ecl_adapter.write("CH3:DISABLE")

# Close the connection
ecl_adapter.close()
