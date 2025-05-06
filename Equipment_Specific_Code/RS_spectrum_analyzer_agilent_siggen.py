import pyvisa
import time
from openpyxl import Workbook

def main():
    rm = pyvisa.ResourceManager()
    print("Instruments:", rm.list_resources())

    siggen = rm.open_resource('GPIB0::19::INSTR')
    sa     = rm.open_resource('GPIB0::20::INSTR')
    siggen.timeout = sa.timeout = 10000  # ms

    # ── Static SA setup ───────────────────────────────────────────
    sa.write("FREQ:MODE SWE")       # sweep mode
    sa.write("BAND:RES 50kHz")      # RBW
    sa.write("BAND:VID 200kHz")     # VBW
    sa.write("SWE:TIME 500ms")      # sweep time
    sa.write("INP:ATT 0dB")         # input attenuation
    sa.write("DISP:TRAC:Y:RLEV -5dBm")  # reference level
    sa.write("FREQ:SPAN 200MHz")    # span
    sa.write("INIT:CONT OFF")       # single‐sweep

    # ── SigGen amplitude & RF ON ──────────────────────────────────
    siggen.write(":SOUR:POW:LEV:IMM:AMPL -5DBM")
    siggen.write(":OUTP:STATe ON")
    time.sleep(0.2)

    # ── Prepare Excel ─────────────────────────────────────────────
    wb = Workbook()
    ws = wb.active
    ws.title = "Sweep Data"
    ws.append([
        "SigGen Freq (GHz)",
        "SigGen Power (dBm)",
        "SA Peak Freq (GHz)",
        "SA Peak Level (dBm)"
    ])

    # ── Sweep 0.1 → 5.0 GHz in 0.1 GHz steps ──────────────────────
    f = 0.1
    while f <= 5.0:
        freq_str = f"{f:.1f}GHZ"

        # 1) set siggen freq
        siggen.write(f":FREQ {freq_str}")

        # 2) update SA center to match
        sa.write(f"FREQ:CENT {freq_str}")

        time.sleep(0.05)

        # 3) read back siggen settings
        freq_rb = siggen.query(":FREQ?").strip()   # e.g. "0.1GHZ"
        pwr_rb  = siggen.query(":POW?").strip()    # e.g. "0.0"

        # 4) trigger & wait single sweep
        sa.write("INIT:IMM")
        sa.query("*OPC?")

        # 5) 1-point peak search
        sa.write("CALC:MARK:FUNC:FPE 1")
        sa.query("*OPC?")

        # 6) fetch peak freq & level
        pf = float(sa.query("CALC:MARK:FUNC:FPE:X?").strip()) / 1e9
        pl = float(sa.query("CALC:MARK:FUNC:FPE:Y?").strip())

        # 7) log to Excel
        ws.append([f, pwr_rb, round(pf, 6), pl])
        print(f"{f:.1f} GHz → SA peak {pf:.3f} GHz @ {pl:.2f} dBm")

        f += 0.1

    # ── Tidy up & save ────────────────────────────────────────────
    for col in ws.columns:
        maxlen = max(len(str(c.value)) for c in col)
        ws.column_dimensions[col[0].column_letter].width = maxlen + 2

    out_file = "measurement_data.xlsx"
    wb.save(out_file)
    print("Saved:", out_file)

    siggen.write(":OUTP:STATe OFF")
    siggen.close()
    sa.close()

if __name__ == "__main__":
    main()
