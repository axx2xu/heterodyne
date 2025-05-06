import pyvisa
import time

def main():
    rm = pyvisa.ResourceManager()
    print("Connected instruments:", rm.list_resources())
    
    # Change this to your actual VISA address
    addr = 'GPIB0::13::INSTR'
    meter = rm.open_resource(addr)
    meter.timeout = 5000  # in ms
    
    # 1) Zero sensor A
    # Syntax: ZERO <s>      s = A or B
    # Zeros the selected sensor, compensating for internal noise/EMF. :contentReference[oaicite:0]{index=0}:contentReference[oaicite:1]{index=1}
    print("Zeroing sensor A…")
    meter.write('ZERO A')
    time.sleep(8)  # allow zero to complete

    # 2) Set the calibration‐factor lookup frequency
    # Syntax: CFFRQ <s>,<value>[units]
    #    e.g. CFFRQ A,25E9 sets 25 GHz for sensor A. :contentReference[oaicite:2]{index=2}:contentReference[oaicite:3]{index=3}
    print("Setting frequency to 2 GHz…")
    meter.write('CFFRQ A,2E9')
    
    #  — and tell it to use that frequency as cal‐factor source
    # Syntax: CFSRC <s>,FREQ
    #    FREQ = use the value from CFFRQ for calibration interpolation. :contentReference[oaicite:4]{index=4}:contentReference[oaicite:5]{index=5}
    meter.write('CFSRC A,FREQ')

    # 3) Trigger a new reading
    # Syntax: STA <c>
    #    Restarts averaging/measurement on channel c (1 or 2). :contentReference[oaicite:6]{index=6}:contentReference[oaicite:7]{index=7}
    meter.write('STA 1')
    time.sleep(0.1)
    
    # 4) Query the next available power reading on channel 1
    # Syntax: O <c>
    #    Returns one raw reading (dBm or Watt) from channel c. :contentReference[oaicite:8]{index=8}:contentReference[oaicite:9]{index=9}
    print("Reading power…")
    reading = meter.query('O 1')
    
    # reading comes back as ASCII, e.g. "-12.34"
    pwr_dbm = float(reading.strip())
    print(f"Measured power: {pwr_dbm:.2f} dBm")

    meter.close()

if __name__ == '__main__':
    main()
