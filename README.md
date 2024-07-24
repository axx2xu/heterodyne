# Python Heterodyne Measurement Automation
## Overview
This repository contains the code and instructions for automating heterodyne measurements using Python. The automation involves controlling various instruments, collecting data, and plotting the results.

## Code Overview
### Program Process
- Calibrates lasers 3 and 4 to a user-specified threshold of starting beat frequency.
- Iterates through a user-specified number of steps between start and stop beat frequency.
- Updates live plots of beat frequency vs. step, photocurrent vs. beat frequency, and raw RF power vs. beat frequency during the loop.
- After the loop, sorts plots by beat frequency and adds a calibrated RF power vs. beat frequency plot.
- Prompts the user to save the data to a .txt file.

### Required Equipment
- Anritsu ECL Lasers 3 and 4
- HP ESA 8565E
- Newport Commercial PD
- HP Wavelength Meter 86120C
- KEOPSYS Optical Amplifier
- Agilent Attenuator 81577A (VOA)
- Keithley Source meter 2400-C (High Power 3)
- Rohde & Schwarz Power Meter NRP-Z58


### Required Manual Inputs and Adjustments
- Turn on all equipment.
- Turn on Delta WL mode on Wavelength Meter.
- Enable lasers and set output power.
- Enable Keithley Source meter – set output voltage and current threshold.
- Enable optical amplifier and turn on the pump.
- Enable VOA and adjust attenuation.

### User Inputs in Program
- Starting wavelength for laser 3 and laser 4.
- Start and end beat frequency.
- Acceptable threshold within the start frequency to begin.
- Number of steps.
- Delay time (seconds) between updating lasers and taking new measurements.
- Output data to .txt file (yes/no).
  - File name.
  - Device number.
  - Trial comments.
  
### Program Output
The program generates live plots and saves data to a .txt file. The plots include:

- Beat frequency vs. step number.
- Photocurrent vs. beat frequency.
- Raw RF power vs. beat frequency.
- Prior to Running the Program

### File Pathing and GPIB Addresses
Ensure correct file paths for calculating calibrated RF power.
- The program reads RF probe loss data from an s2p file.
- The program reads RF link loss data from an Excel file with no headers (column 1: frequencies, column 2: link loss in dB).
- Verify all GPIB addresses for the instruments.

### Python Library, Driver, and Software Installation
- pip install pymeasure pyvisa matplotlib scikit-rf numpy pandas openpyxl scipy
- Install NRP Toolkit.
- Install VISA Library Passport for NRP.
- Install NI Max.
- Install R&S Power Viewer (optional for continuous power output reading).

## Using the Program
### Steps of Use
- Ensure setup is correct (cables, Bias-T, chip, etc.).
- Turn on all equipment.
- Turn on Delta WL mode on Wavelength Meter.
- Enable lasers and set output power.
- Turn VOA power on, do not enable (ensure α is set to 33).
- Turn EDFA power switch and key to ON.
- Set Keithley voltage and current compliance, enable.
- Ensure correct bias is applied.
- Turn on EDFA pump (lasers must be enabled first).
- Enable the VOA, read the photocurrent on Keithley, and adjust α until reaching the desired photocurrent.
- Run the program and enter user inputs into the terminal.
  - The program can be canceled at any time using (ctrl + c); no data saved.
  - To exit the loop early, type any keystroke in the terminal; the loop will exit and data will be saved.
  - At the conclusion of the program, plots will remain open and can be saved or closed.
  - Once the plots are closed, the user will be prompted to output the data or not.

  
## Suggested Settings
### Wavelength and Delay Inputs
- Center wavelength of 1555nm generally provides the best results.
- Set laser 3 wavelength 1-2nm below laser 4 for initial trials to avoid calibration loop issues.
- A delay of over 2 seconds typically offers the best results.

### Typical Issues
- Common issues when laser 4 is set between 1544.1 – 1544.5nm.
- Common issues when laser 4 is set between 1549.5 – 1549.9nm.
- Common issues when laser 4 is set between 1554.9 – 1555.2nm.

### Beat Frequency vs. Step Graphing for Different Center Wavelengths
- 1545nm: Inconsistent for all delay periods, less inconsistency with longer delay.
- 1550nm: Inconsistent for all delay periods, less inconsistency with longer delay.
- 1555nm: Inconsistent for all delay periods, longer delay may have less consistency but were still inconsistent.
Conclusion:
- Of the three tested center wavelengths, 1555nm appears to be the most consistent. Results may differ significantly between periods where the instruments have been turned off. Consecutive measurements may also be inconsistent with little delay between them. The measurements seem to be more accurate as the instruments have been on for longer periods of time. Recommend testing different center wavelengths and delays to find what is most accurate at the time before taking notable measurement recordings.

## Contact
For any questions or further information, please contact Thomas Keyes at axx2xu@virginia.edu.

All this information can also be found in the file named "Python Heterodyne Automation Slides.pdf", along with sample images from the program
