from pymeasure.adapters import VISAAdapter
import time
import math
import pyvisa
import sys
import matplotlib.pyplot as plt
import os

################################################################################################################################################################################
#                         **** THIS IS A TEST CODE TO TEST THE CONNECTION TO THE ECL MODULE, WAVELENGTH METER, AND ESA, IT CURRENTLY FUNCTIONS GENERALLY AS EXPECTED
#                         **** HOWEVER AT THIS TIME IT ONLY OUTPUTS THE BEAT FREQUENCY
#                         ****         MAKE SURE TO TURN ON THE BEAT WL MODE ON THE WAVELENGTH METER BEFORE RUNNING THE PROGRAM OR IT WILL NOT WORK
#
################################################################################################################################################################################

# Disconnects from instruments and exits the program using (Ctrl + C)
def exit_program(ecl_adapter):
    """Exit program and close the connection."""
    ecl_adapter.close()
    wavelength_meter.close()
    spectrum_analyzer.close()
    keithley.close()
    print("Connection closed.")
    sys.exit(0)

# Defined function to find the peak frequency using the peak search function on the ESA
def measure_peak_frequency(spectrum_analyzer):
    """Measure the peak frequency using the spectrum analyzer."""
    try:
        """Perform peak search and return the peak frequency within the current span."""
        spectrum_analyzer.write('MKPK HI')
        time.sleep(0.5)  # Allow time for the peak search to complete
        peak_freq_1 = spectrum_analyzer.query('MKF?')
        time.sleep(0.5)
        spectrum_analyzer.write('MKPK HI')
        time.sleep(0.5)  # Allow time for the peak search to complete
        peak_freq_2 = spectrum_analyzer.query('MKF?')
        time.sleep(0.5)
        spectrum_analyzer.write('MKPK HI')
        time.sleep(0.5)  # Allow time for the peak search to complete
        peak_freq_3 = spectrum_analyzer.query('MKF?')

        peak_freq = min(peak_freq_1, peak_freq_2, peak_freq_3)

        return float(peak_freq) / 1e9 # Convert Hz to GHz
    except Exception as e:
        print("Error measuring peak frequency with spectrum analyzer:", e)
        return None

# Defined function to measure the beat frequency using the wavelength meter
def measure_wavelength_beat(wavelength_meter):
    """Measure the beat frequency using the wavelength meter."""
    try:
        wavelength_meter.write(":INIT:IMM") # Initiate a single measurement
        time.sleep(1.5)  # Increased wait time for the measurement to complete
        freq_data = wavelength_meter.query(":CALC3:DATA? FREQuency").strip().split(',') # Query to get the frequency, strip to only get the first two values (corresponding to laser 3 and 4)
        freqs = [float(freq) for freq in freq_data]
        if len(freqs) == 2:
            return abs(freqs[0] - freqs[1]) / 1e9  # Calculate the difference in GHz
    except Exception as e:
        print("Error measuring beat frequency with wavelength meter:", e)
        return None

# Defined functions to set the laser wavelength and power
def set_laser_wavelength(ecl_adapter, channel, wavelength):
    """Set the laser wavelength."""
    print(f"Setting laser {channel} wavelength to {wavelength:.3f} nm...")
    ecl_adapter.write(f"CH{channel}:L={wavelength:.3f}")

def set_laser_power(ecl_adapter, channel, power):
    """Set the laser power."""
    print(f"Setting laser {channel} output power to {power:.2f} dBm...")
    ecl_adapter.write(f"CH{channel}:P={power:.2f}")

# Initialize the VISA resource manager
rm = pyvisa.ResourceManager()

# List all connected VISA devices (Optional: To verify connections)
print("Connected devices:", rm.list_resources())

# Create a VISA adapter for the ECL laser and wavelength meter
# *** The GPIB should always be the same, so these should not need to be changed ***

ecl_adapter = rm.open_resource('GPIB0::10::INSTR')  # Update with your actual GPIB address
wavelength_meter = rm.open_resource('GPIB0::20::INSTR')  # Update with your actual GPIB address
spectrum_analyzer = rm.open_resource('GPIB0::18::INSTR')  # Update with your actual GPIB address
keithley = rm.open_resource('GPIB0::24::INSTR')  # Update with your actual GPIB address
RS_power_sensor = rm.open_resource('RSNRP::0x00a8::100940::INSTR') # Update with your actual VISA address for the RS NRP-Z58 sensor

# Set timeout to 5 seconds (5000 milliseconds)
ecl_adapter.timeout = 5000
wavelength_meter.timeout = 5000
spectrum_analyzer.timeout = 5000
keithley.timeout = 5000

# Set the Keithley to local mode
keithley.write(":SYSTem:LOCal")

print("Please enter the parameters for the measurement:")

# Value for laser 3 input wavelength check
while True:
    laser_3_WL = float(input("Enter the starting WL for laser 3 (nm): "))
    if 1540 <= laser_3_WL <= 1640:
        break
    else:
        print("Invalid input. Please enter a wavelength between 1540 and 1640 nm.")

# Value for laser 4 input wavelength check
while True:
    laser_4_WL = float(input("Enter the starting WL for laser 4 (nm) (recommend 2 nm below Laser 3): "))
    if 1540 <= laser_4_WL <= 1640:
        break
    else:
        print("Invalid input. Please enter a wavelength between 1540 and 1640 nm.")

# Value for laser 4 frequency step size check
laser_4_step = float(input("Enter the frequency step size for laser 4 (GHz): "))

# Value for the final frequency stepped to (i.e. 1 GHz step size with final frequency of 100 GHz will have 100 steps)
# final_freq = float(input("Enter the final frequency (GHz): "))
num_steps = int(input("Enter the number of steps you would like the program to take: "))

# User input for frequency threshold
user_input = input("Use default threshold of 1GHz (Enter D) or custom threshold in GHz (Enter C) (NOTE: Values under 1GHz are not guaranteed to work): ").strip().upper()
if user_input == 'D':
    freq_threshold = 1
elif user_input == 'C':
    freq_threshold = float(input("Enter frequency threshold (GHz): "))
else:
    print("Invalid input. Using default frequency threshold of 1 GHz.")
    freq_threshold = 1

# Enter a delay time for the lasers to stabilize
delay = int(input("Enter the delay time for the lasers to stabilize (seconds): "))

# Lists to store step number, beat frequency, and laser 4 wavelength
steps = []
beat_freqs = []
laser_4_wavelengths = []
beat_freq_and_power = []

try:
    # Set the laser wavelengths and power
    set_laser_wavelength(ecl_adapter, 3, laser_3_WL)
    set_laser_wavelength(ecl_adapter, 4, laser_4_WL)
    set_laser_power(ecl_adapter, 3, 1.00)
    set_laser_power(ecl_adapter, 4, 1.00)

    # Wait for the lasers to stabilize
    print("Waiting for the lasers to stabilize...")
    time.sleep(10)

    c = 299792458  # Speed of light in m/s

    # Turn on beat wavelength measurement mode
    wavelength_meter.write(":CALC3:DELTA:WAV:STAT ON")  # THIS DOES NOT SEEM TO WORK, TURN THE DELTA WL MODE ON MANUALLY BEFORE RUNNING THE PROGRAM
    time.sleep(1)  # Small delay to ensure command is processed

    last_beat_freq = 0  # Initialize the last beat frequency to 0

    # Set the reference frequency to laser 3
    laser_3_freq = c / (laser_3_WL * 1e-9)  # Convert nm to meters and calculate frequency
    wavelength_meter.write(f":CALC3:DELTA:REF:FREQ {laser_3_freq}")
    time.sleep(1)  # Small delay to ensure command is processed

    calibration_freq = 1  # Set the calibration frequency to 1 GHz

    # Loop through until the starting frequency is within the given threshold
    while calibration_freq >= freq_threshold:

        # Calculate the new wavelength for laser 4
        laser_4_freq = c / (laser_4_WL * 1e-9)  # Calculate current frequency of laser 4

        # Measure beat frequency using wavelength meter and ESA
        wl_meter_beat_freq = measure_wavelength_beat(wavelength_meter)
        esa_beat_freq = measure_peak_frequency(spectrum_analyzer)

        # If wl_meter_beat_freq is None, it means the wavelength meter failed to measure, so we use the ESA value
        if wl_meter_beat_freq is None:
            wl_meter_beat_freq = esa_beat_freq

        if wl_meter_beat_freq is None or esa_beat_freq is None:
            continue

        if wl_meter_beat_freq > 50:
            print(f"Beat Frequency (Wavelength Meter): {wl_meter_beat_freq} GHz")
            calibration_freq = wl_meter_beat_freq

            # Update the wavelength based on current frequency difference
            if wl_meter_beat_freq > 100:
                laser_4_new_freq = laser_4_freq - (50 * 1e9)
            elif 200 < wl_meter_beat_freq < 300:
                laser_4_new_freq = laser_4_freq - (150 * 1e9)
            elif 300 < wl_meter_beat_freq < 400:
                laser_4_new_freq = laser_4_freq - (250 * 1e9)
            elif wl_meter_beat_freq > 400:
                laser_4_new_freq = laser_4_freq - (350 * 1e9)
            else:
                laser_4_new_freq = laser_4_freq - (25 * 1e9)
                
            laser_4_WL = (c / laser_4_new_freq) * 1e9  # Set the new wavelength

            # Check if the new wavelength is within the bounds of the ECL laser
            if 1540 < laser_4_WL < 1660:
                set_laser_wavelength(ecl_adapter, 4, laser_4_WL)
            else:
                print(f"New wavelength for laser 4 is out of bounds: {laser_4_WL:.3f} nm")
                exit_program(ecl_adapter)
                
        elif esa_beat_freq < 50 and wl_meter_beat_freq < 50:
            print(f"Beat Frequency (ESA): {esa_beat_freq} GHz")
            calibration_freq = esa_beat_freq

            if esa_beat_freq > 40:
                laser_4_new_freq = laser_4_freq - (30 * 1e9)
            elif 30 < esa_beat_freq < 40:
                laser_4_new_freq = laser_4_freq - (20 * 1e9)
            elif 20 < esa_beat_freq < 30:
                laser_4_new_freq = laser_4_freq - (10 * 1e9)
            elif 10 < esa_beat_freq < 20:
                laser_4_new_freq = laser_4_freq - (8 * 1e9)
            elif 5 < esa_beat_freq < 10:
                laser_4_new_freq = laser_4_freq - (3 * 1e9)
            elif 1.5 < esa_beat_freq < 5:
                laser_4_new_freq = laser_4_freq - (0.5 * 1e9)
            elif 1 <= esa_beat_freq < 1.5:
                laser_4_new_freq = laser_4_freq - (.2 * 1e9)
            elif esa_beat_freq < 1:  # the value is within the threshold so break out of the loop as to not update the wavelength again unnecessarily
                break

            laser_4_WL = (c / laser_4_new_freq) * 1e9
            set_laser_wavelength(ecl_adapter, 4, laser_4_WL)

        time.sleep(1)

    laser_4_cal = round(laser_4_WL,3) # Store the calibrated wavelength for laser 4 to use in txt file header

    time.sleep(5)

    print("CALIBRATION COMPLETE, STARTING FREQUENCY WITHIN THRESHOLD. BEGINNING STEPS...")

    last_beat_freq = calibration_freq

    plt.ion()  # Turn on interactive mode for live plotting

    fig, axes = plt.subplots(2, 2, figsize=(10, 12))
    fig.delaxes(axes[0, 1])  # Remove the unnecessary subplot
    fig.suptitle(f"Center Wavelength for Laser 3: {laser_3_WL:.2f} nm, Delay: {delay} s, Threshold: {freq_threshold} GHz", fontsize=16)
    ax1, ax3, ax4 = axes[0, 0], axes[1, 0], axes[1, 1]

    ax2 = ax1.twinx()  # Create a twin y-axis for the first subplot

    # Setup first subplot (ax1)
    color = 'tab:blue'
    ax1.set_xlabel('Step Number')
    ax1.set_ylabel('Beat Frequency (GHz)', color=color)
    ax1.set_title('Beat Frequency vs Step Number')
    line1, = ax1.plot([], [], marker='o', linestyle='-', color=color)
    ax1.tick_params(axis='y', labelcolor=color)
    ax1.grid(True)

    color = 'tab:red'
    ax2.set_ylabel('Laser 4 Wavelength (nm)', color=color)
    line2, = ax2.plot([], [], marker='x', linestyle='--', color=color)
    ax2.tick_params(axis='y', labelcolor=color)

    # Setup second subplot (ax3)
    color = 'tab:blue'
    ax3.set_xlabel('Beat Frequency (GHz)')
    ax3.set_ylabel('RF Power (dBm)', color=color)
    ax3.set_title('RF Power vs Beat Frequency')
    line3, = ax3.plot([], [], marker='o', linestyle='-', color=color)
    ax3.tick_params(axis='y', labelcolor=color)
    ax3.grid(True)

    # Setup third subplot (ax4)
    color = 'tab:blue'
    ax4.set_xlabel('Beat Frequency (GHz)')
    ax4.set_ylabel('Current (A)', color=color)
    ax4.set_title('Measured Current vs Beat Frequency')
    line4, = ax4.plot([], [], marker='o', linestyle='-', color=color)
    ax4.tick_params(axis='y', labelcolor=color)
    ax4.grid(True)

    fig.tight_layout()
    # Add a centered title for the entire figure
    

    for step in range(num_steps):
        beat_freq = measure_peak_frequency(spectrum_analyzer) if last_beat_freq < 45 else measure_wavelength_beat(wavelength_meter)
        if beat_freq is None:
            continue
        print(f"Beat Frequency: {beat_freq} GHz")

        # MEASURE CURRENT FROM KEITHLEY
        response = keithley.query(":MEASure:CURRent?")
        current_values = response.split(',')

        if len(current_values) > 1:
            current = float(current_values[1])
            current = f"{current:.2e}"  # Format for display
            print("Measured Current:", current)

        # Set the Keithley back to local mode
        keithley.write(":SYSTem:LOCal")


        # WRITE TO THE R&S POWER SENSOR
        RS_power_sensor.write('INIT:CONT OFF')
        RS_power_sensor.write('SENS:FUNC "POW:AVG"')
        RS_power_sensor.write(f'SENS:FREQ {beat_freq}e9')  # Update with new beat frequency
        RS_power_sensor.write('SENS:AVER:COUN:AUTO ON')
        RS_power_sensor.write('SENS:AVER:COUN 16')
        RS_power_sensor.write('SENS:AVER:STAT ON')
        RS_power_sensor.write('SENS:AVER:TCON REP')
        RS_power_sensor.write('SENS:POW:AVG:APER 5e-3')
        RS_power_sensor.write('INIT:IMM')

        output = RS_power_sensor.query('TRIG:IMM')
        output = output.split(',')[0]  # split the output to only get first value (power in Watts)
        output_dbm = math.log10(float(output)) * 10 + 30  # convert from watts to dBm
        
        output_dbm = round(output_dbm, 2)  # round to 2 decimal places
        beat_freq = round(beat_freq, 1) # Round beat frequency to 1 decimal place
        current = float(current) # Convert current back to float

        beat_freq_and_power.append((beat_freq, output_dbm, current))

        last_beat_freq = beat_freq
        steps.append(step + 1)
        beat_freqs.append(beat_freq)
        laser_4_wavelengths.append(laser_4_WL)

        laser_4_freq = c / (laser_4_WL * 1e-9)
        laser_4_new_freq = laser_4_freq - (laser_4_step * 1e9)
        laser_4_WL = (c / laser_4_new_freq) * 1e9
        set_laser_wavelength(ecl_adapter, 4, laser_4_WL)
        print(f"(step {step + 1}/{num_steps})")

        line1.set_data(steps, beat_freqs)
        line2.set_data(steps, laser_4_wavelengths)
        line3.set_data(beat_freqs, [x[1] for x in beat_freq_and_power])
        line4.set_data(beat_freqs, [x[2] for x in beat_freq_and_power])
        ax1.relim()
        ax2.relim()
        ax3.relim()
        ax4.relim()
        ax1.autoscale_view()
        ax2.autoscale_view()
        ax3.autoscale_view()
        ax4.autoscale_view()
        plt.draw()
        plt.pause(0.1)

        time.sleep(delay)  # Delay for the lasers to stabilize based on user input

    plt.ioff()  # Turn off interactive mode

except KeyboardInterrupt:
    exit_program(ecl_adapter)

ecl_adapter.close()
wavelength_meter.close()
spectrum_analyzer.close()
keithley.close()
print("Connection closed.")
sys.exit(0)

# Static plot at the end
line1.set_data(steps, beat_freqs)
line2.set_data(steps, laser_4_wavelengths)
line3.set_data(beat_freqs, [x[1] for x in beat_freq_and_power])
ax1.relim()
ax2.relim()
ax1.autoscale_view()
ax2.autoscale_view()

beat_freq_and_power.sort(key=lambda x: x[0])
beat_freqs_pow, powers, currents = zip(*beat_freq_and_power)

line3.set_data(beat_freqs_pow, powers)
ax3.relim()
ax3.autoscale_view()
line4.set_data(beat_freqs_pow, currents)
ax4.relim()
ax4.autoscale_view()

plt.draw()
plt.show()

# Save data to a text file
# Get user input for file name
filename_input= input("Enter your desired file name: ").strip().upper()
extension = '.txt'
counter = 1

txt_filename = f'{filename_input}{extension}'
while os.path.exists(txt_filename):
    txt_filename = f'{filename_input}_{counter}{extension}'
    counter += 1


device_num = int(input("Enter the device number: "))
comment = input("Enter any trial comments: ").strip().upper()
keithley_voltage = float(input("Enter the voltage for the Keithley: "))


with open(txt_filename, 'w') as f:
    f.write("Device Number: " + str(device_num) + "\n")
    f.write("Keithley Voltage: " + str(keithley_voltage) + " V" + "\n")
    f.write("Comments: " + comment + "\n")
    f.write("Starting Wavelength for Laser 3: " + str(laser_3_WL) + " nm" + " Starting Wavelength for Laser 4: " +str(laser_4_cal) + " nm" + "\n")
    f.write("Date: " + time.strftime("%m/%d/%Y") + "\n")
    f.write("Time: " + time.strftime("%H:%M:%S") + "\n")
    f.write("\n")
    f.write("F_Beat (GHz)\tRF Pow (dBm)\tCurrent (A)\n")
    for i in range(len(steps)):
        f.write(f"{beat_freqs_pow[i]}\t\t{powers[i]}\t\t{currents[i]}\n") # Write the beat frequency, power, and current to the file in columns

print(f"Data saved to {txt_filename}")
