from pymeasure.adapters import VISAAdapter
import time
import math
import pyvisa
import sys
import matplotlib.pyplot as plt

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
    print("Connection closed.")
    sys.exit(0)

# Defined function to find the peak frequency using the peak search function on the ESA
def measure_peak_frequency(spectrum_analyzer):
    """Measure the peak frequency using the spectrum analyzer."""
    try:
        """Perform peak search and return the peak frequency within the current span."""
        spectrum_analyzer.write('MKPK HI')
        time.sleep(2)  # Allow time for the peak search to complete
        peak_freq = spectrum_analyzer.query('MKF?')
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
RS_power_sensor = rm.open_resource('RSNRP::0x00a8::100940::INSTR') # Update with your actual VISA address for the RS NRP-Z58 sensor

# Set timeout to 5 seconds (10000 milliseconds)
ecl_adapter.timeout = 5000  # in milliseconds
wavelength_meter.timeout = 5000  # in milliseconds
spectrum_analyzer.timeout = 5000  # in milliseconds

print("Clearing previous configurations...")
# Clear previous configurations
wavelength_meter.clear()
spectrum_analyzer.clear()

print("Configuring measurement instruments...")

wavelength_meter.write('*CLS')
spectrum_analyzer.write('*CLS')

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
#final_freq = float(input("Enter the final frequency (GHz): "))
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
            print(f"beat Frequency (Wavelength Meter): {wl_meter_beat_freq} GHz")
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
            print(f"beat Frequency (ESA): {esa_beat_freq} GHz")
            calibration_freq = esa_beat_freq

            if esa_beat_freq > 40:
                laser_4_new_freq = laser_4_freq - (35 * 1e9)
            elif 30 < esa_beat_freq < 40:
                laser_4_new_freq = laser_4_freq - (25 * 1e9)
            elif 20 < esa_beat_freq < 30:
                laser_4_new_freq = laser_4_freq - (15 * 1e9)
            elif 10 < esa_beat_freq < 20:
                laser_4_new_freq = laser_4_freq - (8 * 1e9)
            elif 5 < esa_beat_freq < 10:
                laser_4_new_freq = laser_4_freq - (3.5 * 1e9)
            elif 1.5 < esa_beat_freq < 5:
                laser_4_new_freq = laser_4_freq - (0.8 * 1e9)
            elif 1 <= esa_beat_freq < 1.5:
                laser_4_new_freq = laser_4_freq - (.1 * 1e9)
            elif esa_beat_freq < 1: # the value is within the threshold so break out of the loop as to not update the wavelength again unnecessarily
                break

            laser_4_WL = (c / laser_4_new_freq) * 1e9
            set_laser_wavelength(ecl_adapter, 4, laser_4_WL)

        
        time.sleep(1)

    time.sleep(5)

    print("CALIBRATION COMPLETE, STARTING FREQUENCY WITHIN THRESHOLD. BEGINNING STEPS...")

    last_beat_freq = calibration_freq

    plt.ion()  # Turn on interactive mode for live plotting
    fig, ax1 = plt.subplots(figsize=(10, 6))

    color = 'tab:blue'
    ax1.set_xlabel('Step Number')
    ax1.set_ylabel('beat Frequency (GHz)', color=color)
    line1, = ax1.plot([], [], marker='o', linestyle='-', color=color)
    ax1.tick_params(axis='y', labelcolor=color)
    ax1.grid(True)

    ax2 = ax1.twinx()  # instantiate a second axes that shares the same x-axis
    color = 'tab:red'
    ax2.set_ylabel('Laser 4 Wavelength (nm)', color=color)
    line2, = ax2.plot([], [], marker='x', linestyle='--', color=color)
    ax2.tick_params(axis='y', labelcolor=color)

    fig.tight_layout()  # otherwise the right y-label is slightly clipped
    plt.title(f"beat Frequency and Laser 4 Wavelength vs Step Number (Starting Wavelength for Laser 3: {laser_3_WL:.2f} nm)")

    for step in range(num_steps):
        if last_beat_freq < 45:
            beat_freq = measure_peak_frequency(spectrum_analyzer)
            if beat_freq is None:
                continue
            print(f"beat Frequency (ESA): {beat_freq} GHz")
        else:
            beat_freq = measure_wavelength_beat(wavelength_meter)
            if beat_freq is None:
                continue
            print(f"beat Frequency (Wavelength Meter): {beat_freq} GHz")

        # WRITE TO THE R&S POWER SENSOR
        RS_power_sensor.write('INIT:CONT OFF')
        RS_power_sensor.write('SENS:FUNC "POW:AVG"')
        RS_power_sensor.write(f'SENS:FREQ {beat_freq}e9') # Update with new beat frequency
        RS_power_sensor.write('SENS:AVER:COUN:AUTO ON')
        RS_power_sensor.write('SENS:AVER:COUN 16')
        RS_power_sensor.write('SENS:AVER:STAT ON')
        RS_power_sensor.write('SENS:AVER:TCON REP')
        RS_power_sensor.write('SENS:POW:AVG:APER 5e-3')
        RS_power_sensor.write('INIT:IMM')
    
        # Query immediate trigger
        print("Triggering and reading data...")
        output = RS_power_sensor.query('TRIG:IMM')
        output = output.split(',')[0]
        output_dbm = math.log10(float(output)) * 10 + 30
        #print("Output:", output_dbm, "dBm")

        beat_freq_and_power.append((beat_freq, output_dbm))

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
        ax1.relim()
        ax2.relim()
        ax1.autoscale_view()
        ax2.autoscale_view()
        plt.draw()
        plt.pause(0.1)

        time.sleep(1)

    time.sleep(5)
    plt.ioff()  # Turn off interactive mode

except KeyboardInterrupt:
    exit_program(ecl_adapter)


ecl_adapter.close()
print("Connection closed.")

# Static plot at the end: beat Frequency and Laser 4 Wavelength vs Step Number
line1.set_data(steps, beat_freqs)
line2.set_data(steps, laser_4_wavelengths)
ax1.relim()
ax2.relim()
ax1.autoscale_view()
ax2.autoscale_view()
plt.draw()
plt.show()

# Plot power vs beat frequency
beat_freq_and_power.sort(key=lambda x: x[0]) # sort the values by beat frequency
beat_freqs_pow, powers = zip(*beat_freq_and_power) # unpack the sorted values

fig, ax1 = plt.subplots(figsize=(10, 6))

color = 'tab:blue'
ax1.set_xlabel('Beat Frequency (GHz)')
ax1.set_ylabel('Power Measurement (dBm)', color=color)
ax1.plot(beat_freqs_pow, powers, marker='o', linestyle='-', color=color)
ax1.tick_params(axis='y', labelcolor=color)
ax1.grid(True)

fig.tight_layout()  # otherwise the right y-label is slightly clipped
plt.title(f"Delta Frequency and Laser 4 Wavelength vs Step Number (Starting Wavelength for Laser 3: {laser_3_WL:.2f} nm)")
plt.show()

