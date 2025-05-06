import pyvisa
import time

def main():
    # 1. Initialize VISA and list connected instruments
    rm = pyvisa.ResourceManager()
    instruments = rm.list_resources()
    print("Connected instruments:", instruments)
    
    # # 2. Replace with your analyzer's VISA address (GPIB, USB, TCP/IP, etc.)
    sa_addr = "GPIB0::20::INSTR"
    sa = rm.open_resource(sa_addr)
    sa.timeout = 10000  # 10 s timeout
    
    # # 3. Identify the instrument (optional, but useful for verifying connection)
    print("IDN:", sa.query("*IDN?"))
    
    # # 4. Switch to frequency-domain sweep mode
    sa.write("FREQ:MODE SWE")  # conform to analyzer sweep mode :contentReference[oaicite:0]{index=0}:contentReference[oaicite:1]{index=1}:contentReference[oaicite:2]{index=2}:contentReference[oaicite:3]{index=3}

    # # 5. Set resolution bandwidth to 50 kHz
    sa.write("BAND:RES 50kHz")  # SENSe:BANDwidth[:RESolution] :contentReference[oaicite:4]{index=4}:contentReference[oaicite:5]{index=5}

    # # 6. Set video bandwidth to 200 kHz
    sa.write("BAND:VID 200kHz")  # SENSe:BANDwidth:VIDeo :contentReference[oaicite:6]{index=6}:contentReference[oaicite:7]{index=7}

    # # 7. Fix sweep time at 500 ms
    sa.write("SWE:TIME 500ms")  # SWEep:TIME 2.5 ms–16000 s :contentReference[oaicite:8]{index=8}:contentReference[oaicite:9]{index=9}

    # # 8. Set RF input attenuation to 0 dB
    sa.write("INP:ATT 0dB")  # INPut:ATTenuation 0 dB–75 dB :contentReference[oaicite:10]{index=10}:contentReference[oaicite:11]{index=11}

    # # 9. Set reference level (top of screen) to –50 dBm
    sa.write("DISP:TRAC:Y:RLEV -45dBm")  # DISPlay:TRACe:Y:RLEVel –130 dBm–30 dBm :contentReference[oaicite:12]{index=12}:contentReference[oaicite:13]{index=13}

    # # 10. Define span and center frequency
    sa.write("FREQ:SPAN 200MHz")   # SENSe:FREQuency:SPAN 0–fmax :contentReference[oaicite:14]{index=14}:contentReference[oaicite:15]{index=15}
    sa.write("FREQ:CENT 2.2GHz")   # SENSe:FREQuency:CENTer 0–fmax :contentReference[oaicite:16]{index=16}:contentReference[oaicite:17]{index=17}

    # # 11. Trigger an immediate sweep and wait for it to finish
    sa.write("INIT:IMM")
    sa.query("*OPC?")  # blocks until sweep complete

    print("Sweep complete")
    sa.close()

if __name__ == "__main__":
    main()
