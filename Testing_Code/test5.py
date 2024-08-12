import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from openpyxl import Workbook
import matplotlib.pyplot as plt
import threading
import time
import math
import numpy as np
import mplcursors

# Initialize the main Tkinter window
root = tk.Tk()
root.title("Measurement and Plotting GUI")
root.geometry("1200x800")  # Adjust the size as needed

# Create a frame for user inputs on the left side
input_frame = ttk.Frame(root, width=200)
input_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)

# Create a frame for the plots on the right side
plot_frame = ttk.Frame(root)
plot_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)

# Create input labels and fields
tk.Label(input_frame, text="Number of Steps:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
num_steps_var = tk.IntVar(value=10)
num_steps_entry = ttk.Entry(input_frame, textvariable=num_steps_var)
num_steps_entry.grid(row=0, column=1, padx=5, pady=5)

tk.Label(input_frame, text="Delay Between Steps (s):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
delay_var = tk.DoubleVar(value=3.0)
delay_entry = ttk.Entry(input_frame, textvariable=delay_var)
delay_entry.grid(row=1, column=1, padx=5, pady=5)

tk.Label(input_frame, text="Starting Beat Frequency (GHz):").grid(row=2, column=0, padx=5, pady=5, sticky="e")
start_freq_var = tk.DoubleVar(value=10.0)
start_freq_entry = ttk.Entry(input_frame, textvariable=start_freq_var)
start_freq_entry.grid(row=2, column=1, padx=5, pady=5)

tk.Label(input_frame, text="Ending Beat Frequency (GHz):").grid(row=3, column=0, padx=5, pady=5, sticky="e")
end_freq_var = tk.DoubleVar(value=20.0)
end_freq_entry = ttk.Entry(input_frame, textvariable=end_freq_var)
end_freq_entry.grid(row=3, column=1, padx=5, pady=5)

# Create control buttons below the inputs
start_button = ttk.Button(input_frame, text="Start", command=lambda: threading.Thread(target=data_collection).start())
start_button.grid(row=4, column=0, columnspan=2, pady=10)

stop_button = ttk.Button(input_frame, text="Stop", command=lambda: on_stop())
stop_button.grid(row=5, column=0, columnspan=2, pady=10)

save_button = ttk.Button(input_frame, text="Save", command=lambda: on_save())
save_button.grid(row=6, column=0, columnspan=2, pady=10)

# Initialize Matplotlib Figure and Axes
fig, axes = plt.subplots(2, 2, figsize=(8, 6))
fig.suptitle("Real-time Measurement Plots", fontsize=16)
ax1, ax3, ax4, ax5 = axes[0, 0], axes[0, 1], axes[1, 0], axes[1, 1]
ax2 = ax1.twinx()  # Create a twin y-axis for the first subplot

# Embed the Matplotlib figure into Tkinter
canvas = FigureCanvasTkAgg(fig, master=plot_frame)
canvas.draw()
canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

# Optional: Add Matplotlib Navigation Toolbar
toolbar = NavigationToolbar2Tk(canvas, plot_frame)
toolbar.update()
canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

# Initialize data lists for plotting
steps = []
beat_freqs = []
laser_4_wavelengths = []
beat_freq_and_power = []  # List of tuples: (beat_freq, output_dbm, current, p_actual)
calibrated_rf = []
photo_currents = []
rf_loss = []
powers = []
p_actuals = []

# Define threading events
stop_event = threading.Event()  # Event to stop the data collection
data_ready_event = threading.Event()  # Event to signal that new data is ready for plotting

# Setup first subplot (ax1)
color1 = 'tab:blue'
ax1.set_xlabel('Step Number')
ax1.set_ylabel('Beat Frequency (GHz)', color=color1)
ax1.set_title('Beat Frequency vs Step Number')
line1, = ax1.plot([], [], marker='o', linestyle='-', color=color1)
ax1.tick_params(axis='y', labelcolor=color1)
ax1.grid(True)

color2 = 'tab:red'
ax2.set_ylabel('Laser 4 Wavelength (nm)', color=color2)
line2, = ax2.plot([], [], marker='x', linestyle='--', color=color2)
ax2.tick_params(axis='y', labelcolor=color2)
ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x:.3f}'.rstrip('0').rstrip('.')))

# Setup second subplot (ax3)
color3 = 'tab:blue'
ax3.set_xlabel('Beat Frequency (GHz)')
ax3.set_ylabel('Raw RF Power (dBm)', color=color3)
ax3.set_title('Raw RF Power vs Beat Frequency')
line3, = ax3.plot([], [], marker='o', linestyle='-', color=color3)
ax3.tick_params(axis='y', labelcolor=color3)
ax3.grid(True)

# Setup third subplot (ax4)
color4 = 'tab:blue'
ax4.set_xlabel('Beat Frequency (GHz)')
ax4.set_ylabel('Photocurrent (mA)', color=color4)
ax4.set_title('Measured Photocurrent vs Beat Frequency')
line4, = ax4.plot([], [], marker='o', linestyle='-', color=color4)
ax4.tick_params(axis='y', labelcolor=color4)
ax4.grid(True)

# Setup fourth subplot (ax5)
color5 = 'tab:blue'
ax5.set_xlabel('Beat Frequency (GHz)')
ax5.set_ylabel('Calibrated RF Power (dBm)', color=color5)
ax5.set_title('Calibrated RF Power vs Beat Frequency')
line5, = ax5.plot([], [], marker='o', linestyle='-', color=color5)
ax5.tick_params(axis='y', labelcolor=color5)
ax5.grid(True)

fig.tight_layout()

# Add hover functionality using mplcursors, but only annotate markers (not the line itself)
mplcursors.cursor(line1, hover=True)
mplcursors.cursor(line2, hover=True)
mplcursors.cursor(line3, hover=True)
mplcursors.cursor(line4, hover=True)
mplcursors.cursor(line5, hover=True)

# Define the data collection function
def data_collection():
    global steps, beat_freqs, laser_4_wavelengths, beat_freq_and_power
    try:
        # Get user inputs for the measurement
        num_steps = num_steps_var.get()
        delay = delay_var.get()
        start_freq = start_freq_var.get()
        end_freq = end_freq_var.get()

        # Calculate step size
        laser_4_step = (end_freq - start_freq) / num_steps  # Calculate the step size for laser 4 frequency

        # Initialize laser wavelengths, frequencies, and other parameters
        # For example:
        c = 299792458  # Speed of light in m/s
        laser_3_WL = 1550  # Starting wavelength for laser 3 in nm
        laser_4_WL = 1550  # Starting wavelength for laser 4 in nm

        # Begin data collection loop
        for step in range(num_steps):
            if stop_event.is_set():
                print("Data collection stopped by user.")
                break

            print(f"Step {step + 1} of {num_steps}")

            # Simulate data acquisition
            # Replace these with actual instrument measurements
            beat_freq = start_freq + step * ((end_freq - start_freq) / num_steps)
            output_dbm = -30 + step  # Dummy RF power value
            current = 5 + 0.1 * step  # Dummy photocurrent value
            p_actual = -10 + 0.2 * step  # Dummy VOA P actual value

            # Append data to lists
            steps.append(step + 1)
            beat_freqs.append(beat_freq)
            laser_4_wavelengths.append(laser_4_WL)
            beat_freq_and_power.append((beat_freq, output_dbm, current, p_actual))
            powers.append(output_dbm)
            photo_currents.append(current)
            p_actuals.append(p_actual)

            # Update laser wavelength for next step
            laser_4_freq = c / (laser_4_WL * 1e-9)
            laser_4_new_freq = laser_4_freq - (laser_4_step * 1e9)
            laser_4_WL = (c / laser_4_new_freq) * 1e9
            # set_laser_wavelength(ecl_adapter, 4, laser_4_WL)  # Uncomment when using actual instruments

            # Simulate calibration (Replace with actual calibration computation)
            rf_loss_value = 2  # Dummy RF loss value
            rf_loss.append(rf_loss_value)
            calibrated_power = output_dbm - rf_loss_value
            calibrated_rf.append(calibrated_power)

            # Signal that new data is ready
            data_ready_event.set()

            # Wait for the specified delay
            time.sleep(delay)

        print("Data collection completed.")

    except Exception as e:
        print(f"Error in data collection: {e}")
        messagebox.showerror("Data Collection Error", str(e))
        stop_event.set()

# Define the function to update plots
def update_plots():
    if data_ready_event.is_set():
        # Update line data
        line1.set_data(steps, beat_freqs)
        line2.set_data(steps, laser_4_wavelengths)
        line3.set_data(beat_freqs, powers)
        line4.set_data(beat_freqs, photo_currents)
        line5.set_data(beat_freqs, calibrated_rf)

        # Adjust axes
        ax1.relim()
        ax2.relim()
        ax3.relim()
        ax4.relim()
        ax5.relim()

        ax1.autoscale_view()
        ax2.autoscale_view()
        ax3.autoscale_view()
        ax4.autoscale_view()
        ax5.autoscale_view()

        canvas.draw()

        # Clear the event
        data_ready_event.clear()

    # Schedule the next check
    if not stop_event.is_set():
        root.after(100, update_plots)

# Define the function to handle the "Stop" button
def on_stop():
    if messagebox.askyesno("Confirm Stop", "Are you sure you want to stop the data collection?"):
        stop_event.set()
        print("Data collection will be stopped.")

# Define the function to handle the "Save" button with a pop-up input window
def on_save():
    # Prompt user for file path
    file_path = filedialog.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
    )
    if not file_path:
        return  # User cancelled

    # Create a pop-up window for additional inputs
    input_window = tk.Toplevel(root)
    input_window.title("Save Data Inputs")
    input_window.geometry("300x200")

    # Create input fields for device number and comments
    tk.Label(input_window, text="Device Number:").grid(row=0, column=0, padx=10, pady=10)
    device_num_var = tk.StringVar()
    device_num_entry = ttk.Entry(input_window, textvariable=device_num_var)
    device_num_entry.grid(row=0, column=1, padx=10, pady=10)

    tk.Label(input_window, text="Comments:").grid(row=1, column=0, padx=10, pady=10)
    comment_var = tk.StringVar()
    comment_entry = ttk.Entry(input_window, textvariable=comment_var)
    comment_entry.grid(row=1, column=1, padx=10, pady=10)

    def save_data():
        device_num = device_num_var.get().strip()
        comment = comment_var.get().strip().upper()
        keithley_voltage = "0.0"  # Replace with actual voltage measurement

        laser_3_WL = 1550  # Starting wavelength for laser 3 in nm

        with open(file_path, 'w') as f:
            f.write("DEVICE NUMBER: " + str(device_num) + "\n")
            f.write("COMMENTS: " + comment + "\n")
            f.write("KEITHLEY VOLTAGE: " + str(keithley_voltage) + " V" + "\n")
            f.write("INITIAL PHOTOCURRENT: " + str(photo_currents[0]) + " (mA)" + "\n")
            f.write("STARTING WAVELENGTH FOR LASER 3: " + str(laser_3_WL) + " (nm) :" + " STARTING WAVELENGTH FOR LASER 4: " + str(laser_4_wavelengths[0]) + " (nm) :" + " DELAY: " + str(delay_var.get()) + " (s) " + "\n")
            f.write("DATE: " + time.strftime("%m/%d/%Y") + "\n")
            f.write("TIME: " + time.strftime("%H:%M:%S") + "\n")
            f.write("\n")
            f.write("F_BEAT(GHz)\tPHOTOCURRENT (mA)\tRaw RF POW (dBm)\tRF Loss (dB)\t\tCal RF POW (dBm)\tVOA P Actual (dBm)\n")
            for i in range(len(steps)):
                f.write(f"{beat_freqs[i]:<10.2f}\t{float(photo_currents[i]):<10.4e}\t\t{powers[i]:<10.2f}\t\t{rf_loss[i]:<10.2f}\t\t{calibrated_rf[i]:<10.2f}\t\t{p_actuals[i]:<10.3f}\n")

        # Save the plot as an image
        plot_file_path = file_path.rsplit('.', 1)[0] + '.png'
        fig.savefig(plot_file_path, bbox_inches='tight')
        print(f"Data and plot saved to {file_path} and {plot_file_path}")

        # Create Excel Workbook and Sheet
        wb = Workbook()
        ws = wb.active
        ws.title = "Experiment Data"

        # Write the header information
        ws.append(["DEVICE NUMBER", device_num])
        ws.append(["COMMENTS", comment])
        ws.append(["KEITHLEY VOLTAGE", f"{keithley_voltage} V"])
        ws.append(["INITIAL PHOTOCURRENT", f"{photo_currents[0]} (mA)"])
        ws.append(["STARTING WAVELENGTH FOR LASER 3", f"{laser_3_WL} (nm)"])
        ws.append(["STARTING WAVELENGTH FOR LASER 4", f"{laser_4_wavelengths[0]} (nm)"])
        ws.append(["DELAY", f"{delay_var.get()} (s)"])
        ws.append(["DATE", time.strftime("%m/%d/%Y")])
        ws.append(["TIME", time.strftime("%H:%M:%S")])
        ws.append([])  # Add an empty row for spacing

        # Write the table header
        ws.append(["F_BEAT (GHz)", "PHOTOCURRENT (mA)", "Raw RF POW (dBm)", "RF Loss (dB)", "Cal RF POW (dBm)", "VOA P Actual (dBm)"])

        # Write the data rows
        for i in range(len(steps)):
            ws.append([
                f"{beat_freqs[i]:.2f}",
                f"{float(photo_currents[i]):.4e}",
                f"{powers[i]:.2f}",
                f"{rf_loss[i]:.2f}",
                f"{calibrated_rf[i]:.2f}",
                f"{p_actuals[i]:.3f}"
            ])

        # Adjust column widths
        count = 0
        for column in ws.columns:
            if(count == 1):
                continue # Don't reformat the second column as long comments will make the column too wide
            max_length = 0
            column = list(column)  # Convert the column to a list
            for cell in column:
                try:
                    # Compute the length of the cell value
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)  # Add a little extra space
            ws.column_dimensions[column[0].column_letter].width = adjusted_width
            count+=1 # Increment column count

        # Save the workbook
        excel_file_path = file_path.replace(".txt", ".xlsx")
        wb.save(excel_file_path)
        print(f"Excel data saved to {excel_file_path}")

        # Close the pop-up window
        input_window.destroy()

    # Add a button to save the data and close the pop-up
    save_button = ttk.Button(input_window, text="Save", command=save_data)
    save_button.grid(row=2, column=0, columnspan=2, pady=20)

# Define the function to handle window closing
def on_closing():
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        stop_event.set()
        root.destroy()

root.protocol("WM_DELETE_WINDOW", on_closing)

# Start the plot updating loop
root.after(100, update_plots)

# Start the Tkinter main loop
root.mainloop()

print("Program terminated gracefully.")
