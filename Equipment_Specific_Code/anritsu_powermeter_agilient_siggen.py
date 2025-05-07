import pyvisa
import time
from openpyxl import Workbook
from datetime import datetime

def main():
    rm = pyvisa.ResourceManager()
    print("Instruments:", rm.list_resources())

    siggen = rm.open_resource('GPIB0::19::INSTR')
    meter = rm.open_resource('GPIB0::13::INSTR')
    siggen.timeout = 5000  # ms
    meter.timeout = 5000  # ms

    ############# INPUT PARAMETERS #############
    start_pow = -10  # dBm
    end_pow = 6
    step_pow = 1  # dBm

    start_freq = 0.1  # GHz
    stop_freq = 10  # GHz
    step_freq = 0.1  # GHz
    ############################################

    # ── Static Power Meter setup ───────────────────────────────────────────
    meter.write('CFFRQ A,2E9')
    meter.write('CFSRC A,FREQ')
    
    # ── SigGen amplitude & RF ON ──────────────────────────────────
    siggen.write(f":SOUR:POW:LEV:IMM:AMPL {start_pow}DBM")
    siggen.write(":OUTP:STATe ON")
    time.sleep(0.2)

    # ── Prepare Excel ─────────────────────────────────────────────
    wb = Workbook()
    ws = wb.active
    ws.title = "Sweep Data"
    ws.append([
        "SigGen Freq (GHz)",
        "SigGen Power (dBm)",
        "Power Meter Reading (dBm)",
        "Unlevel Error"
    ])

    # ── Sweep 0.1 → 5.0 GHz in 0.1 GHz steps ──────────────────────
    pow = start_pow
    while pow <= end_pow:
        f = start_freq
        print(f"Power: {pow} dBm")
        print(f"")

        while f <= stop_freq:
            freq_str = f"{f:.1f}GHZ"

            # 1) set siggen freq
            siggen.write(f":FREQ {freq_str}")

            # 2) update power meter cal freq to match
            meter.write(f"CFFRQ A,{f}E9")
            meter.write('CFSRC A,FREQ')

            time.sleep(0.1)

            # 3) read back siggen settings
            pwr_rb  = siggen.query(":POW?").strip()    # e.g. "0.0"

            meter.write('STA 1')
            reading = meter.query('O 1')
            time.sleep(0.1)

            # 6) fetch peak freq & level
            pwr_dbm = float(reading.strip())

            # … after reading your power meter …
            pow_cond   = int(siggen.query('STAT:QUES:POW:COND?'))
            unlevel_now = bool(pow_cond & 2)

            ws.append([
                f,                # GHz
                pwr_rb,           # SigGen-reported power
                pwr_dbm,          # meter reading
                unlevel_now,      # current unleveled state (True/False)
            ])

            print(f"{f:.1f} GHz → Power Meter Reading: {pwr_dbm:.2f} dBm")

            f += step_freq
        pow += step_pow
        siggen.write(f":SOUR:POW:LEV:IMM:AMPL {pow}DBM")
        time.sleep(1)

    # ── Tidy up & save ────────────────────────────────────────────
    for col in ws.columns:
        maxlen = max(len(str(c.value)) for c in col)
        ws.column_dimensions[col[0].column_letter].width = maxlen + 2

    out_file = "measurement_data.xlsx"
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_file = f"measurement_data_{ts}.xlsx"
    wb.save(out_file)
    print("Saved:", out_file)

    siggen.write(":OUTP:STATe OFF")
    siggen.close()
    meter.close()

if __name__ == "__main__":
    main()
