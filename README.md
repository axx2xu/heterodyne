# Python Heterodyne Measurement Automation

**Thomas Keyes**  
*Computer Engineering, University of Virginia*
*axx2xu@virginia.edu*
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

- After pressing the “START” button, a new window will pop up, prompting the user to enter the device number, comments, and to select the file save path. The program will then begin.
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
  - Select file paths in the GUI for:
    - s2p file for RF probe loss.
    - Excel file for RF link loss (no headers, frequencies in column 1, link loss in dB in column 2).

### Required Installations
- Install NRP Toolkit.
- Install VISA Library Passport for NRP.
- Install NI Max.
- Install R&S Power Viewer.

*Installation links can be found in the “drivers” file in the GitHub repository.*

---

## How to Get To and Run Program On Offline Computer

1. Open Visual Studio Code.
2. Find the directory named “Thomas”.
3. Open the file named `heterodyne_automation.py`.
4. Run the program by selecting the triangle button in the top right of the software.
5. Enter the inputs into the command terminal.
6. Monitor the measurement loop and stop it by typing any keystroke in the terminal.
7. Save plots by clicking the disk icon in the plot window.
8. After closing the plot window, choose to output the data to a text file and provide the required details.

---
## Opening the Program (Ideal Method)
- The designed and ideal way to run the program is to simply run the distributed .exe file named “heterodyne_automation.exe”
- If this runs successfully, the GUI should pop up as shown in the figure
- If this does not work, follow the steps usinga virtual environment to run the program through python

## Setting Up a Virtual Environment to Run the Program

1. Ensure that Python is installed on your computer.
2. Download the “heterodyne_automation.py” and “requirements.txt” files from the GitHub repository into a new project directory.
3. Open a terminal and navigate to the project directory.
4. Create a virtual environment:
   ```sh
   python -m venv venv
   ```
5. Activate the virtual environment:
   - On Windows:
     ```sh
     .\venv\Scripts\activate
     ```
   - On macOS/Linux:
     ```sh
     source venv/bin/activate
     ```
6. Install the required packages:
   ```sh
   pip install -r requirements.txt
   ```
7. Run the program:
   ```sh
   python heterodyne_automation.py
   ```
---

## Using the Program

### Steps of Use
#### Pre-Program Setup
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

#### Setting Up User Inputs
1. Starting WL for Laser 3
- This will be a stable wavelength and will not be changed by the program; it should be set to your desired wavelength.
2. Starting WL for laser 4
- This output will be altered by the program to drive the beat frequency. The initial setting should be 2-3nm below that of Laser 3 unless you have already calibrated your starting wavelengths.
3. If you have already manually calibrated the starting beat frequency, enter the exact wavelengths from the ECL into the inputs of the program, and de-select the “automatic initial beat frequency search” check.
4. Starting and ending beat frequency set the range of beat frequencies that the program will parse and take measurements of.
5. Number of steps
- This will determine the number of steps the program will take between the start and end beat frequency.
6. Delay between steps
- This will set the delay the program takes between updating the lasers and taking the next measurements. I recommend at least a 3 second delay to ensure accurate measurements.
7. RF link loss file
- If applicable include existing RF link loss file, it must have .s2p file formatting.
8. RF probe loss file
- If applicable, include existing RF probe loss file, it must have .xlsx file formatting.
9. Press the “START” button
- A new window will pop up prompting the user to enter the device number, user comments, and save file path
10. Device number
- Enter the device number or another identification parameter (e.g., serial number).
11. Comments
- Enter any additional user comments about the device or trial run
12. Choose file save path
- Search the file directory for the desired save location, then select the save button.

#### While the Program is Running
- Once the measurement loop begins, current measurements will appear in the output window, such as photocurrent, RF power, etc.
- To stop the loop during the measurement, press the “STOP” button. Any measurements that have been taken will be automatically saved.
- To reset the program, press the “RESET” button. No data will be saved and the program should restart. If the program is not running correctly, it may need to be closed and re-opened.
- The plots and data will automatically save to the path file selected during the initial program setup. If you would like to save the plots again, press the save button to and select the desired file path.

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
