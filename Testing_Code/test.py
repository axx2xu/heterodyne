import matplotlib.pyplot as plt
import numpy as np
import math
import time
import threading

# Sample constants for testing
c = 3e8  # Speed of light in m/s
end_freq = 30  # GHz
start_freq = 5  # GHz
num_steps = 10  # Number of steps

# Sample data lists
steps = []
beat_freqs = []
laser_4_wavelengths = []
beat_freq_and_power = []

# Mock functions for testing (replace with actual hardware interface functions)
def measure_peak_frequency(spectrum_analyzer):
    return np.random.uniform(10, 40)

def measure_wavelength_beat(wavelength_meter):
    return np.random.uniform(10, 40)

def set_laser_wavelength(adapter, channel, wavelength):
    pass

def exit_loop():
    pass

# Setup the plot
fig, axes = plt.subplots(2, 2, figsize=(10, 8))
fig.suptitle(f"Center Wavelength for Laser 3: {start_freq:.2f} nm, Delay: {num_steps} s", fontsize=16)

# Unpack axes for individual plots
ax1, ax3, ax4, ax5 = axes[0, 0], axes[0, 1], axes[1, 0], axes[1, 1]

# Plot 1: Beat Frequency vs Step Number
ax2 = ax1.twinx()  # Create a twin y-axis for the first subplot
ax1.set_xlabel('Step Number')
ax1.set_ylabel('Beat Frequency (GHz)', color='tab:blue')
ax1.set_title('Beat Frequency vs Step Number')
line1, = ax1.plot([], [], marker='o', linestyle='-', color='tab:blue', label='Beat Frequency')
ax1.tick_params(axis='y', labelcolor='tab:blue')
ax1.grid(True)

ax2.set_ylabel('Laser 4 Wavelength (nm)', color='tab:red')
line2, = ax2.plot([], [], marker='x', linestyle='--', color='tab:red', label='Laser 4 Wavelength')
ax2.tick_params(axis='y', labelcolor='tab:red')
ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{x:.3f}'.rstrip('0').rstrip('.')))

# Plot 2: Raw RF Power vs Beat Frequency
ax3.set_xlabel('Beat Frequency (GHz)')
ax3.set_ylabel('Raw RF Power (dBm)', color='tab:blue')
ax3.set_title('Raw RF Power vs Beat Frequency')
line3, = ax3.plot([], [], marker='o', linestyle='-', color='tab:blue', label='Raw RF Power')
ax3.tick_params(axis='y', labelcolor='tab:blue')
ax3.grid(True)

# Plot 3: Measured Photocurrent vs Beat Frequency
ax4.set_xlabel('Beat Frequency (GHz)')
ax4.set_ylabel('Photocurrent (mA)', color='tab:blue')
ax4.set_title('Measured Photocurrent vs Beat Frequency')
line4, = ax4.plot([], [], marker='o', linestyle='-', color='tab:blue', label='Photocurrent')
ax4.tick_params(axis='y', labelcolor='tab:blue')
ax4.grid(True)

# Plot 4: Calibrated RF Power vs Beat Frequency
ax5.set_xlabel('Beat Frequency (GHz)')
ax5.set_ylabel('Calibrated RF Power (dBm)', color='tab:blue')
ax5.set_title('Calibrated RF Power vs Beat Frequency')
line5, = ax5.plot([], [], marker='o', linestyle='-', color='tab:blue', label='Calibrated RF Power')
ax5.tick_params(axis='y', labelcolor='tab:blue')
ax5.grid(True)

# Adjust layout
fig.tight_layout(rect=[0, 0.03, 1, 0.95])  # Add space for suptitle

# Maximize the window for different backends
manager = plt.get_current_fig_manager()
try:
    # TkAgg backend
    manager.window.state('zoomed')
except AttributeError:
    try:
        # Qt5Agg backend
        manager.window.showMaximized()
    except AttributeError:
        try:
            # WxAgg backend
            manager.window.Maximize(True)
        except AttributeError:
            pass  # If none of these work, just continue

# Calculate step size
laser_4_step = (end_freq - start_freq) / num_steps  # Calculate the step size for laser 4 wavelength

# Simulate measurements
for step in range(num_steps):
    # Simulated measurement values
    beat_freq = np.random.uniform(10, 40)
    current = np.random.uniform(0.5, 1.5)  # Simulated current
    p_actual = np.random.uniform(-30, -20)  # Simulated power

    beat_freq_and_power.append((beat_freq, np.random.uniform(-30, -20), current, p_actual))

    # Update the lists
    steps.append(step + 1)
    beat_freqs.append(beat_freq)
    laser_4_wavelengths.append(np.random.uniform(1500, 1600))

    # Update the plots with the new data
    line1.set_data(steps, beat_freqs)
    line2.set_data(steps, laser_4_wavelengths)
    line3.set_data(beat_freqs, [x[1] for x in beat_freq_and_power])
    line4.set_data(beat_freqs, [x[2] for x in beat_freq_and_power])
    line5.set_data(beat_freqs, [x[3] for x in beat_freq_and_power])

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

    plt.draw()
    plt.pause(0.1)

# Hover function for each plot
def hover(event, ax, line, data_x, data_y, annot):
    if ax == event.inaxes:  # Check if the event is within the current axis
        cont, ind = line.contains(event)
        if cont:  # If the hover event is over a data point
            update_annot(ind, line, annot, data_x, data_y)
            annot.set_visible(True)
            fig.canvas.draw_idle()
        else:
            annot.set_visible(False)
            fig.canvas.draw_idle()

def update_annot(ind, line, annot, data_x, data_y):
    # Get the position of the hovered point
    x, y = data_x[ind["ind"][0]], data_y[ind["ind"][0]]
    annot.xy = (x, y)
    text = f"{line.get_label()}:\nX: {x:.2f}\nY: {y:.2f}"
    annot.set_text(text)
    annot.get_bbox_patch().set_facecolor('yellow')
    annot.get_bbox_patch().set_alpha(0.7)

# Initialize annotations for each axis
annot1 = ax1.annotate("", xy=(0, 0), xytext=(20, 20),
                      textcoords="offset points", bbox=dict(boxstyle="round", fc="w"),
                      arrowprops=dict(arrowstyle="->"))
annot1.set_visible(False)

annot2 = ax3.annotate("", xy=(0, 0), xytext=(20, 20),
                      textcoords="offset points", bbox=dict(boxstyle="round", fc="w"),
                      arrowprops=dict(arrowstyle="->"))
annot2.set_visible(False)

annot3 = ax4.annotate("", xy=(0, 0), xytext=(20, 20),
                      textcoords="offset points", bbox=dict(boxstyle="round", fc="w"),
                      arrowprops=dict(arrowstyle="->"))
annot3.set_visible(False)

annot4 = ax5.annotate("", xy=(0, 0), xytext=(20, 20),
                      textcoords="offset points", bbox=dict(boxstyle="round", fc="w"),
                      arrowprops=dict(arrowstyle="->"))
annot4.set_visible(False)

# Connect hover events for each plot
fig.canvas.mpl_connect("motion_notify_event", lambda event: hover(event, ax1, line1, steps, beat_freqs, annot1))
fig.canvas.mpl_connect("motion_notify_event", lambda event: hover(event, ax3, line3, beat_freqs, [x[1] for x in beat_freq_and_power], annot2))
fig.canvas.mpl_connect("motion_notify_event", lambda event: hover(event, ax4, line4, beat_freqs, [x[2] for x in beat_freq_and_power], annot3))
fig.canvas.mpl_connect("motion_notify_event", lambda event: hover(event, ax5, line5, beat_freqs, [x[3] for x in beat_freq_and_power], annot4))

# Show the plot
plt.show()
