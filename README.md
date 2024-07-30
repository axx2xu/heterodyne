# Python Heterodyne Measurement Automation

**Thomas Keyes**  
*Computer Engineering, University of Virginia*

---

## Program Highlights

1. **Automatic Laser Adjustment:**
   - Adjusts the wavelength of lasers 3 and 4 to find the first measurement beat frequency as defined by the user.
   
2. **Step Size Determination:**
   - Determines the step size to loop from the user-defined start and end beat frequency based on the input number of steps.
   
3. **Automatic Equipment Selection:**
   - Selects the electrical spectrum analyzer for beat frequency measurements under 50 GHz.
   - Selects the wavelength meter for beat frequency measurements at or above 50 GHz.
   
4. **Data Integration:**
   - Reads in data from user-defined s2p files and Excel sheets to add calibrated RF power data to the raw data.

---

## Equipment Used

- **Anritsu ECL Lasers 3 and 4:** Adjusts wavelength.
- **HP ESA 8565E/Newport Commercial PD:** Reads beat frequency < 50 GHz.
- **HP Wavelength Meter 86120C:** Reads beat frequency >= 50 GHz.
- **KEOPSYS Optical Amplifier (EDFA):** Not used by the program.
- **Agilent Attenuator 81577A (VOA):** Reads actual power.
- **Keithley Source meter 2400-C (High Power 3):** Reads photocurrent.
- **Rohde & Shwarz Power Meter NRP-Z58:** Reads RF power up to 110 GHz frequency.

---

## Program Overview

### Process

1. **Calibration:**
   - Calibrates lasers 3 and 4 to a user-specified threshold of starting beat frequency.
   
2. **Measurement Loop:**
   - Iterates through a user-specified number of steps between start and stop beat frequency while measuring beat frequency, photocurrent, and RF power.
   - Updates live plots of beat frequency vs. step, photocurrent vs. beat frequency, and raw RF power vs. beat frequency during the loop.
   
3. **Plotting:**
   - Sorts plots by beat frequency and adds a calibrated RF power vs. beat frequency plot.
   - Prompts the user to save the data to a .txt file.

### Required Manual Inputs and Adjustments

- Turn on all equipment.
- Turn on Delta WL mode on Wavelength Meter.
- Zero R&S Power Meter using Power Viewer Software.
- Enable lasers and set output power.
- Enable Keithley Source meter – set output voltage and current threshold.
- Enable optical amplifier and turn on pump.
- Enable VOA and adjust attenuation.

---

## User Inputs in Program

- Starting wavelength for laser 3 and laser 4.
- Start and end beat frequency.
- Acceptable threshold within start frequency.
- Number of steps.
- Delay time (s) between updating lasers and taking new measurements.
- Output data to .txt file (yes/no).
  - File name.
  - Device number.
  - Trial comments.

---

## Program Output

### Output Plots

1. **Beat Frequency vs. Step Number:**
   - Shows if the laser and steps and beat frequency measurements are updating reasonably.
   
2. **Measured Photocurrent vs. Beat Frequency:**
   - Visualizes changes in the photocurrent reading from the Keithley.
   
3. **Raw RF Power vs. Beat Frequency:**
   - Plots the RF power vs. beat frequency live as the measurement loop runs.
   
4. **Calibrated RF Power vs. Beat Frequency:**
   - Combines the raw RF data with the calibrated RF loss and re-plots after the measurement loop finishes.

### .txt Output Data

- Contains measured data and a mix of user-defined and automatically recorded information in the header.
  - File name.
  - Device number and trial comments.
  - Keithley voltage and initial photocurrent.
  - Starting wavelength for lasers 3 and 4 after calibration.
  - Time and date.
  - Data sorted with the corresponding beat frequency.

---

## Prior to Running the Program

### File Pathing and GPIB Addresses

- **Calibrated RF Power Calculation:**
  - Update file paths within the code.
  - s2p file for RF probe loss.
  - Excel file for RF link loss (no headers, frequencies in column 1, link loss in dB in column 2).

- **GPIB Addresses:**
  - Check and label correctly at the beginning of the code.
  - View connected devices to verify addresses.

### Python Library, Driver, and Software Installation

- `pip install pymeasure pyvisa matplotlib scikit-rf numpy pandas openpyxl scipy`
- Install NRP Toolkit.
- Install VISA Library Passport for NRP.
- Install NI Max.
- Install R&S Power Viewer.

*Installation links can be found in the “requirements” file in the GitHub repository.*

---

## How to Get To and Run Program

1. Open Visual Studio Code.
2. Find the directory named “Thomas”.
3. Open the file named `heterodyne_automation.py`.
4. Run the program by selecting the triangle button in the top right of the software.
5. Enter the inputs into the command terminal.
6. Monitor the measurement loop and stop it by typing any keystroke in the terminal.
7. Save plots by clicking the disk icon in the plot window.
8. After closing the plot window, choose to output the data to a text file and provide the required details.

---

## Using the Program

### Steps of Use

1. Ensure correct setup (cables, Bias-T, chip, etc.).
2. Turn on all equipment.
3. Zero R&S Power Meter using Power Viewer software.
4. Turn on Delta WL mode on Wavelength Meter.
5. Enable lasers and set output power.
6. Turn VOA Power on but do not enable (ensure α is set to 33).
7. Turn EDFA power switch and key to ON.
8. Set Keithley voltage and current compliance, enable (ensure correct bias).
9. Turn on EDFA pump (lasers must be enabled first).
10. Enable the VOA, read the photocurrent on Keithley and adjust α until reaching the desired photocurrent.
11. Run the program and enter user inputs into the terminal.

*The program can be cancelled at any time using (ctrl + c); no data will be saved.*

---

## Suggested Settings

### Wavelength and Delay Inputs

- **Center Wavelength of 1555nm:**
  - Laser 3 set to 1555nm and laser 4 calibrated to match.
  - Wavelength meter typically reads near 1550nm.

- **Laser 4 Adjustment:**
  - Laser 4 actual wavelength is ~1nm below that of laser 3.
  - Set laser 3 wavelength 1-2nm below laser 4 for initial trials.

- **Delay Time:**
  - Delay over 2 seconds typically offers best results.

---

## Access Code and Related Files

- [GitHub Repository](https://github.com/axx2xu/heterodyne)

---

## Graphs and Results

### Beat Frequency vs. Step Graphing for Different Center Wavelengths

- **Center Wavelength 1545nm:**
  - Inconsistency for all delay periods, less inconsistency with longer delay.
  
- **Center Wavelength 1550nm:**
  - Inconsistency for all delay periods, less inconsistency with longer delay.
  
- **Center Wavelength 1555nm:**
  - Inconsistency for all delay periods, longer delay may have less consistency but were still inconsistent.

**Conclusion:**
- Of the three tested center wavelengths, 1555nm appears to be the most consistent. Results may differ significantly between periods where the instruments have been turned off. Consecutive measurements may also be inconsistent with little delay between them. The measurements seem to be more accurate as the instruments have been on for longer periods of time. Recommend to test different center wavelengths and delays to find what is most accurate at the time before taking notable measurement recordings.

---

## Single Point ESA vs. Averaged

- ESA readings are automatically averaged from the instrument.
- Little difference was found adding additional averaging from multiple query points for the ESA reading compared to single point reading.

---

# Contact
For instructions with video demonstration, access the "Python Heterodyne Automation Slides" powerpoint file. To view the slides without video demonstration, access the "Python Heterodyne Automation Slides" pdf file.

If the program is not working correctly, or if you have any questions, feel free to reach out to me by email:
axx2xu@virginia.edu
