import time
import pyvisa
import sys
from openpyxl import Workbook
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import mplcursors
from matplotlib.figure import Figure
from matplotlib.ticker import FuncFormatter


# Initialize the VISA resource manager
rm = pyvisa.ResourceManager()

# List all connected VISA devices (Optional: To verify connections)
print("Connected devices:", rm.list_resources())

# Create a VISA adapter for the ECL laser and wavelength meter
# *** The GPIB should always be the same, so these should not need to be changed ***

ecl_adapter_GPIB = 'GPIB0::10::INSTR' # Update with your actual GPIB address
wavelength_meter_GPIB = 'GPIB0::20::INSTR'  # Update with your actual GPIB address
keithley_GPIB = 'GPIB0::24::INSTR'  # Update with your actual GPIB address

# Function definitions for various measurements
def exit_program():
    """Exit program and close the connection."""
    ecl_adapter.close()
    wavelength_meter.close()
    keithley.close()
    update_message_feed("Connection closed.")
    sys.exit(0)

def measure_wavelength(wavelength_meter):
    """Measure the beat frequency using the wavelength meter."""
    try:
        wavelength_meter.write(":INIT:IMM") # Initiate a single measurement
        time.sleep(0.5)  # Increased wait time for the measurement to complete
        wavelength_data = wavelength_meter.query(":FETCH:SCALar:POWer:WAVelength?").strip().split(',')
        wavelengths = [float(wavelength) for wavelength in wavelength_data]
        if len(wavelengths) == 1:
            return abs(wavelengths[0])
    except Exception as e:
        update_message_feed(f"Error measuring wavelength with wavelength meter: {e}")
        return None

def set_laser_wavelength(ecl_adapter, channel, wavelength):
    """Set the laser wavelength."""
    update_message_feed(f"Setting laser {channel} wavelength to {wavelength:.3f} nm...")
    ecl_adapter.write(f"CH{channel}:L={wavelength:.3f}")


def update_message_feed(message):
    """Update the message feed with a new message."""
    message_feed.insert(tk.END, message + "\n")
    message_feed.see(tk.END)  # Scroll to the latest message
    root.update_idletasks()  # Update the GUI to reflect the new message

# Start threading events
data_ready_event = threading.Event()
stop_event = threading.Event()

# Initialize the main Tkinter window
root = tk.Tk()
root.title("Measurement and Plotting GUI")
root.geometry("1200x800")  # Adjust the size as needed
root.state('zoomed')  # Make the window fullscreen

# Create a frame for user inputs on the left side
input_frame = ttk.Frame(root, width=200)
input_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)

# Create a frame for the plots on the right side
plot_frame = ttk.Frame(root)
plot_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)

# Create input labels and fields
tk.Label(input_frame, text="ECL Laser Output Number:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
laser_var = tk.DoubleVar(value=4)
laser_entry = ttk.Entry(input_frame, textvariable=laser_var)
laser_entry.grid(row=0, column=1, padx=5, pady=5)

tk.Label(input_frame, text="Starting Laser WL (nm):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
laser_val_var = tk.DoubleVar(value=1550)
laser_val_entry = ttk.Entry(input_frame, textvariable=laser_val_var)
laser_val_entry.grid(row=1, column=1, padx=5, pady=5)

tk.Label(input_frame, text="Ending Laser WL (nm):").grid(row=2, column=0, padx=5, pady=5, sticky="e")
laser_end_var = tk.DoubleVar(value=0)
laser_end_entry = ttk.Entry(input_frame, textvariable=laser_end_var)
laser_end_entry.grid(row=2, column=1, padx=5, pady=5)

tk.Label(input_frame, text="Number of Steps:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
num_steps_var = tk.IntVar(value=0)
num_steps_entry = ttk.Entry(input_frame, textvariable=num_steps_var)
num_steps_entry.grid(row=3, column=1, padx=5, pady=5)

tk.Label(input_frame, text="Delay Between Steps (s):").grid(row=4, column=0, padx=5, pady=5, sticky="e")
delay_var = tk.DoubleVar(value=3.0)
delay_entry = ttk.Entry(input_frame, textvariable=delay_var)
delay_entry.grid(row=4, column=1, padx=5, pady=5)

tk.Label(input_frame, text="Enable Automatic Start Wavelength Search:").grid(row=5, column=0, padx=5, pady=5, sticky="e")
enable_search_var = tk.BooleanVar(value=True)  # Default is True (enabled)
enable_search_checkbox = ttk.Checkbutton(input_frame, variable=enable_search_var)
enable_search_checkbox.grid(row=5, column=1, padx=5, pady=5, sticky="w")

# Create control buttons below the inputs
start_button = ttk.Button(input_frame, text="START", command=lambda: threading.Thread(target=data_collection).start())
start_button.grid(row=6, column=0, columnspan=2, pady=10)

stop_button = ttk.Button(input_frame, text="STOP", command=lambda: on_stop())
stop_button.grid(row=7, column=0, columnspan=2, pady=10)
tk.Label(input_frame, text="NOTE: This will only stop data collection during the frequency sweep").grid(row=8, column=0, columnspan=2, pady=2)


# Add a Text widget to display live messages
message_feed = tk.Text(input_frame, wrap="word", height=10, width=50)
message_feed.grid(row=8, column=0, columnspan=3, padx=5, pady=5)

cancel_button = ttk.Button(input_frame, text="RESET", command=lambda: on_cancel())
cancel_button.grid(row=9, column=0, columnspan=2, pady=10)
tk.Label(input_frame, text="NOTE: This will reset the program and no data will be saved").grid(row=9, column=0, columnspan=2, pady=2)

# Initialize Matplotlib Figure and Axes
fig = Figure(figsize=(8, 6))
axes = fig.subplots(1, 2)
fig.suptitle("Measurement Plots", fontsize=16)
ax1, ax3 = axes[0], axes[1]
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
wavelengths = []
laser_wavelengths = []
photo_currents = []

# Setup first subplot (ax1)
color1 = 'tab:blue'
ax1.set_xlabel('Step Number')
ax1.set_ylabel('Measured Wavelength (nm)', color=color1)
ax1.set_title('Wavelength vs Step Number')
line1, = ax1.plot([], [], linestyle='-', color=color1)  # Line without markers
markers1, = ax1.plot([], [], 'o', color=color1)  # Markers only
ax1.tick_params(axis='y', labelcolor=color1)
ax1.grid(True)

color2 = 'tab:red'
ax2.set_ylabel('ECL Set Laser Wavelength (nm)', color=color2)
line2, = ax2.plot([], [], linestyle='--', color=color2)  # Line without markers
markers2, = ax2.plot([], [], 'x', color=color2)  # Markers only
ax2.tick_params(axis='y', labelcolor=color2)
ax2.yaxis.set_major_formatter(FuncFormatter(lambda x, _: f'{x:.3f}'.rstrip('0').rstrip('.')))

# Setup second subplot (ax3)
color3 = 'tab:blue'
ax3.set_xlabel('Measured Wavelength (nm)')
ax3.set_ylabel('Photocurrent (mA)', color=color3)
ax3.set_title('Photocurrent vs. Wavelength')
line3, = ax3.plot([], [], linestyle='-', color=color3)  # Line without markers
markers3, = ax3.plot([], [], 'o', color=color3)  # Markers only
ax3.tick_params(axis='y', labelcolor=color3)
ax3.grid(True)

fig.tight_layout()

# Add hover functionality using mplcursors, but only annotate the markers (actual data points)
mplcursors.cursor(markers1, hover=mplcursors.HoverMode.Transient)
mplcursors.cursor(markers3, hover=mplcursors.HoverMode.Transient)


# Create VISA adapters with the connected equipment - send error message to message feed in GUI if connection fails
def open_instruments():
    global ecl_adapter, wavelength_meter, spectrum_analyzer, keithley, RS_power_sensor, voa
    try:
        ecl_adapter = rm.open_resource(ecl_adapter_GPIB)  # Update with your actual GPIB address
        wavelength_meter = rm.open_resource(wavelength_meter_GPIB)  # Update with your actual GPIB address
        keithley = rm.open_resource(keithley_GPIB)  # Update with your actual GPIB address
        keithley.write(":SYSTem:LOCal")
    except:
        update_message_feed("Error connecting to VISA devices. Check GPIB addresses and connections.")

# Define the data collection function
def data_collection():
    stop_event.clear()
    data_ready_event.clear()

    global steps, wavelengths, laser_wavelengths, photo_currents, wavelength_diff
    global start_time, start_time_sweep, sweep_run_time, total_run_time

    
    """Get user inputs and start data collection process"""

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

      # Function to close the pop-up window and prompt for file path
    def choose_file_and_close():
        global file_path, excel_file_path, plot_file_path

        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if not file_path:
            return  # User cancelled, do not close the window
        excel_file_path = file_path.replace(".txt", ".xlsx")
        plot_file_path = file_path.rsplit('.', 1)[0] + '.png'


        # Close the input window
        input_window.destroy()
        # You can now use `file_path`, `excel_file_path`, and `plot_file_path` in your program

    # Add a button to save the data and close the pop-up
    save_button = ttk.Button(input_window, text="CHOOSE FILE SAVE PATH", command=choose_file_and_close)
    save_button.grid(row=2, column=0, columnspan=2, pady=20)

    try:
        open_instruments()
        time.sleep(5)

        start_time = time.time()

        # Get user inputs for the measurement
        laser_num = laser_var.get()
        laser_start_WL = laser_val_var.get()
        laser_end_WL = laser_end_var.get()
        num_steps = num_steps_var.get()
        delay = delay_var.get()
        enable_search = enable_search_var.get()

        new_laser_wl = laser_start_WL

        if enable_search:
            ##############################################################
            ### Search for the desired wavelength using value from WLM ###
            ##############################################################

            set_laser_wavelength(ecl_adapter, laser_num, laser_val_var-5)

            #Wait for the lasers to stabilize
            update_message_feed("Waiting for the lasers to stabilize...")
            time.sleep(10)
            update_message_feed("RUNNING AUTOMATIC WAVELENGTH SEARCH LOOP...")

            # Measure beat frequency using wavelength meter and ESA
            try:
                current_wavelength = measure_wavelength(wavelength_meter)
            except:
                update_message_feed("Failed to read wavelength from the wavelength meter")

            wavelength_diff = laser_val_var - current_wavelength

            while wavelength_diff >= 0.005:
                if stop_event.is_set():
                    update_message_feed("Data collection stopped by user.")
                    return

                if wavelength_diff >= 0.05:
                    update_message_feed(f"Current Wavelength: {current_wavelength} nm")
                    new_laser_wl = new_laser_wl + 0.9 * (wavelength_diff)

                    # Check if the new wavelength is within the bounds of the ECL laser and different from the current wavelength
                    if 1540 < new_laser_wl < 1660:
                        set_laser_wavelength(ecl_adapter, laser_num, new_laser_wl)
                    else:
                        update_message_feed(f"New wavelength for laser is out of bounds: {new_laser_wl:.3f} nm")
                        exit_program()

                else: 
                    update_message_feed(f"Current Wavelength: {current_wavelength} nm")
                    new_laser_wl = new_laser_wl + 0.8 * (wavelength_diff)

                    # Check if the new wavelength is within the bounds of the ECL laser and different from the current wavelength
                    if 1540 < new_laser_wl < 1660:
                        set_laser_wavelength(ecl_adapter, laser_num, new_laser_wl)
                    else:
                        update_message_feed(f"New wavelength for laser is out of bounds: {new_laser_wl:.3f} nm")
                        exit_program()

                time.sleep(3)
            if stop_event.is_set():
                update_message_feed("Data collection stopped by user.")
                return

        laser_step = (laser_end_WL - laser_start_WL) / num_steps  # Calculate the step size for laser 4 frequency

        update_message_feed("BEGINNING MEASUREMENT LOOP...")
        start_time_sweep = time.time() # Start time for the measurement loop

         # Get initial photocurrent for text output
        # MEASURE CURRENT FROM KEITHLEY
        response = keithley.query(":MEASure:CURRent?")
        initial_current_values = response.split(',')

        if len(initial_current_values) > 1:
            initial_current = float(initial_current_values[1]) * 1000  # Convert to mA
            initial_current = round(initial_current,3)  # Format for display

        # Set the Keithley back to local mode
        keithley.write(":SYSTem:LOCal")

        global looping
        looping = True
        # Begin data collection loop
        for step in range(num_steps):
            if stop_event.is_set():
                update_message_feed("Data collection stopped by user.")
                looping = False
                time_end = time.time() # End time for the measurement loop
                sweep_run_time = time_end - start_time_sweep # Calculate the total run time
                total_run_time = time_end - start_time # Calculate the total run time

                # Update the plots with the final calibrated data
                data_ready_event.set()
                break


            update_message_feed(f"Step {step + 1} of {num_steps}")

             # Measure beat frequency using wavelength meter and ESA
            try:
                current_wavelength = measure_wavelength(wavelength_meter)
            except:
                update_message_feed("Failed to read wavelength from the wavelength meter")


            # MEASURE CURRENT FROM KEITHLEY
            response = keithley.query(":MEASure:CURRent?")
            current_values = response.split(',')

            if len(current_values) > 1:
                current = float(current_values[1]) * 1000  # Convert to mA
                current = round(current,3)  # Format for display

            # Set the Keithley back to local mode
            keithley.write(":SYSTem:LOCal")

            update_message_feed(f"Measured Wavelength: {current_wavelength:.3e} (nm)")
            update_message_feed(f"Measured Photocurrent: {current} (mA)")

            # Append the data to the lists
            steps.append(step + 1)
            laser_wavelengths.append(new_laser_wl)
            wavelengths.append(current_wavelength)
            photo_currents.append(current)
        
            new_laser_wl = new_laser_wl + laser_step #Update wavelength with step size
            set_laser_wavelength(ecl_adapter, laser_num, new_laser_wl)

            data_ready_event.set()

            # Wait for the specified delay
            time.sleep(delay)

        looping = False
        update_message_feed("Data collection completed.")
        time_end = time.time() # End time for the measurement loop
        sweep_run_time = time_end - start_time_sweep # Calculate the total run time
        total_run_time = time_end - start_time # Calculate the total run time

        # Update the plots with the final calibrated data
        data_ready_event.set()

        device_num = device_num_var.get().strip()
        user_comment = comment_var.get().strip().upper()
        keithley_voltage = keithley.query(":SOUR:VOLT:LEV:IMM:AMPL?").strip()  # Get the keithley voltage from the keithley
        keithley_voltage = f"{float(keithley_voltage):.3e}"  # Format the keithley voltage for display
        keithley.write(":SYSTem:LOCal")  # Set the keithley back to local mode

        # Adjust subplot parameters to add space for comments
        fig.subplots_adjust(top=0.8)

        # Add comments to the plot
        comments = [
            f"Device Number: {device_num}",
            f"Comments: {user_comment}",
            f"Date: {time.strftime('%m/%d/%Y')}",
            f"Time: {time.strftime('%H:%M:%S')}",
            f"Frequency Sweep Run Time: {sweep_run_time:.2f} s",
            f"Total Run Time: {total_run_time:.2f} s",
            f"Keithley Voltage: {keithley_voltage} V",
        ]

        # Position for comments (adjust as needed)
        x_comment = 0.5  # X position for the comments
        y_comment_start = 0.94  # Starting Y position for the comments
        y_comment_step = 0.02  # Step size for Y position

        for i, comment in enumerate(comments):
            fig.text(x_comment, y_comment_start - i * y_comment_step, comment, wrap=True, horizontalalignment='center', fontsize=10)

        # Maximize the window for different backends using Tkinter's window management
        try:
            root.state('zoomed')  # Maximizes the window for Windows
        except AttributeError:
            try:
                root.attributes('-fullscreen', True)  # Alternative approach to maximize the window cross-platform
            except AttributeError:
                pass  # If none of these work, just continue


        # Adjust title and axis label font properties
        title_font = {'size': '14', 'weight': 'bold'}
        label_font = {'size': '12', 'weight': 'bold'}

        ax1.set_title('Wavelength vs Step Number', fontdict=title_font)
        ax1.set_xlabel('Step Number', fontdict=label_font)
        ax1.set_ylabel('Measured Wavelength (nm)', fontdict=label_font)
        ax2.set_ylabel('ECL Set Laser Wavelength (nm)', fontdict=label_font)

        ax3.set_title('Raw RF Power vs Beat Frequency', fontdict=title_font)
        ax3.set_xlabel('Measured Wavelength (nm)', fontdict=label_font)
        ax3.set_ylabel('Photocurrent (mA)', fontdict=label_font)

        # Adjust layout
        fig.tight_layout(rect=[0, 0, 1, 0.85])
        canvas.draw()

        with open(file_path, 'w') as f:
            f.write("DEVICE NUMBER: " + str(device_num) + "\n")
            f.write("COMMENTS: " + user_comment + "\n")
            f.write("KEITHLEY VOLTAGE: " + str(keithley_voltage) + " V" + "\n")
            f.write("FREQUENCY SWEEP RUN TIME: " + str(f"{sweep_run_time:.2f}") + " s" + "\n")
            f.write("TOTAL RUN TIME: " + str(f"{total_run_time:.2f}") + " s" + "\n")
            f.write("INITIAL PHOTOCURRENT: " + str(photo_currents[0]) + " (mA)" + "\n")
            f.write("STARTING WAVELENGTH: " + str(laser_start_WL) + " (nm) :"  + " DELAY: " + str(delay_var.get()) + " (s) " + "\n")
            f.write("DATE: " + time.strftime("%m/%d/%Y") + "\n")
            f.write("TIME: " + time.strftime("%H:%M:%S") + "\n")
            f.write("\n")
            f.write("Wavelength (nm)\tI_PD (mA)\n")
            for i in range(len(steps)):
                f.write(f"{wavelengths[i]:<10.3f}{photo_currents[i]:<10.4f}\n")

        # Save the plot as an image
        
        fig.savefig(plot_file_path, bbox_inches='tight')
        update_message_feed(f"Data and plot saved to {file_path} and {plot_file_path}")

        # Create Excel Workbook and Sheet
        wb = Workbook()
        ws = wb.active
        ws.title = "Experiment Data"

        # Write the header information
        ws.append(["DEVICE NUMBER", device_num])
        ws.append(["COMMENTS", user_comment])
        ws.append(["KEITHLEY VOLTAGE", f"{keithley_voltage} V"])
        ws.append(["INITIAL PHOTOCURRENT", f"{photo_currents[0]} (mA)"])
        ws.append(["STARTING WAVELENGTH FOR LASER", f"{laser_start_WL} (nm)"]) 
        ws.append(["DELAY", f"{delay_var.get()} (s)"])
        ws.append(["FREQUENCY SWEEP RUN TIME", f"{sweep_run_time:.2f} s"])
        ws.append(["TOTAL RUN TIME", f"{total_run_time:.2f} s"])
        ws.append(["DATE", time.strftime("%m/%d/%Y")])
        ws.append(["TIME", time.strftime("%H:%M:%S")])
        ws.append([])  # Add an empty row for spacing

        # Write the table header
        ws.append(["Wavelength (nm)", "I_PD (mA)"])

        # Write the data rows
        for i in range(len(steps)):
            ws.append([
                f"{wavelengths[i]:.3f}",
                f"{photo_currents[i]}",
            ])
        # Save the workbook
        
        wb.save(excel_file_path)
        update_message_feed(f"Excel data saved to {excel_file_path}")

    except Exception as e:
        update_message_feed(f"Error in data collection: {e}")
        reset_program()

# Define the function to update plots
def update_plots():
    if data_ready_event.is_set():
        # Update line and marker data
        line1.set_data(steps, wavelengths)
        markers1.set_data(steps, wavelengths)
        line2.set_data(steps, laser_wavelengths)
        markers2.set_data(steps, laser_wavelengths)
        line3.set_data(wavelengths, photo_currents)
        markers3.set_data(wavelengths, photo_currents)


        # Adjust axes
        ax1.relim()
        ax2.relim()
        ax3.relim()

        ax1.autoscale_view()
        ax2.autoscale_view()
        ax3.autoscale_view()

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
        update_message_feed("Data collection will be stopped.")

def reset_program():
    global steps, laser_wavelengths, photo_currents,  stop_event, data_ready_event

    # Reset all the lists and variables
    steps = []
    laser_wavelengths = []
    photo_currents = []

    # Clear the message feed
    message_feed.delete(1.0, tk.END)
    
    # Remove all text items from the figure except the suptitle
    texts_to_remove = [txt for txt in fig.texts if txt != fig._suptitle]

    for txt in texts_to_remove:
        txt.remove()

    # Clear the data from the plots without removing formatting
    line1.set_data([], [])
    markers1.set_data([], [])
    line2.set_data([], [])
    markers2.set_data([], [])
    line3.set_data([], [])
    markers3.set_data([], [])

    # Rescale the axes to reflect the cleared data
    for ax in [ax1, ax2, ax3]:
        ax.relim()
        ax.autoscale_view()

    # Redraw the canvas to update the plots
    canvas.draw()

    # Restart the plot updating loop
    root.after(100, update_plots)

    # Reset the event flags
    stop_event.clear()
    data_ready_event.clear()

    update_message_feed("Program reset and ready to start again.")


# Define the function to handle window closing
def on_closing():
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        stop_event.set()
        root.destroy()
        sys.exit(0)

def on_cancel():
    if messagebox.askyesno("Confirm Exit", "Are you sure you want to reset the program?"):
        stop_event.set()  # Ensure any ongoing measurements are stopped
        time.sleep(4) # Wait for the measurements to stop
        reset_program()  # Reset the program for a new start
        


root.protocol("WM_DELETE_WINDOW", on_closing)

# Start the plot updating loop
root.after(100, update_plots)

# Start the Tkinter main loop
root.mainloop()
