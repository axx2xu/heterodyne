import pyvisa
import time

def main():
    # 1) Initialize VISA and list resources
    rm = pyvisa.ResourceManager()
    resources = rm.list_resources()
    print("Available instruments:", resources)

    # 2) Choose your instrument GPIB address here
    instrument_address = 'GPIB0::19::INSTR'  # ← adjust as needed
    
    # 3) Open the instrument
    sig = rm.open_resource(instrument_address)
    
    # 4) Identify the instrument
    idn = sig.query('*IDN?')
    print("Instrument IDN:", idn.strip())
    
    # 5) Enable RF output
    #    SCPI: OUTPut[:STATe] ON|OFF  :contentReference[oaicite:0]{index=0}:contentReference[oaicite:1]{index=1}
    print("Enabling RF output...")
    sig.write(':OUTP:STATe ON')
    time.sleep(0.1)
    
    # 6) Set output frequency (in CW mode)
    #    SCPI: FREQuency[:CW] <val><unit>  
    #    Example shorthand: :FREQ 2GHZ  :contentReference[oaicite:2]{index=2}:contentReference[oaicite:3]{index=3}
    freq = '2GHZ'
    print(f"Setting frequency to {freq}...")
    sig.write(f':FREQ {freq}')
    time.sleep(0.1)
    
    # 7) Set output power level (CW)
    #    SCPI: [:SOURce]:POWer:LEVel:IMMediate:AMPLitude <val><unit>  
    #    (alias “POW <val>DBM”)  :contentReference[oaicite:4]{index=4}:contentReference[oaicite:5]{index=5}
    power = '-20DBM'
    print(f"Setting power to {power}...")
    sig.write(f':SOUR:POW:LEV:IMM:AMPL {power}')
    time.sleep(0.1)
    
    # 8) Query back each setting
    freq_q = sig.query(':FREQ?').strip()
    pow_q  = sig.query(':POW?').strip()
    out_q  = sig.query(':OUTP?').strip()
    
    print(f"Frequency reading: {freq_q}")
    print(f"Power reading:     {pow_q}")
    print(f"RF Output state:   {('ON' if out_q in ['1','ON'] else 'OFF')}")


    time.sleep(1)
    # 9) Disable RF output if you like
    # sig.write(':OUTP:STATe OFF')

    # 10) Clean up
    sig.close()
    print("Done.")

if __name__ == '__main__':
    main()
