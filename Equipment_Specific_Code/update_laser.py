import pyvisa
import time

rm = pyvisa.ResourceManager()

# List all connected VISA devices (Optional: To verify connections)
print("Connected devices:", rm.list_resources())

ecl_adapter_GPIB = 'GPIB0::10::INSTR' # Update with your actual GPIB address
ecl_adapter = rm.open_resource(ecl_adapter_GPIB)  # Update with your actual GPIB address
voa_GPIB = 'GPIB0::26::INSTR'  # Update with your actual GPIB address
voa = rm.open_resource(voa_GPIB)

#ecl_adapter.timeout = 1000  # in milliseconds

def set_laser_wavelength(ecl_adapter, channel, wavelength):
    """Set the laser wavelength."""
    ecl_adapter.write(f"CH{channel}:L={wavelength:.3f}")

time.sleep(3)

laser_channel = 3
new_wavelength = 1555
try:
    set_laser_wavelength(ecl_adapter, laser_channel, new_wavelength)
except:
    print("Updating laser failed")

time.sleep(10)

