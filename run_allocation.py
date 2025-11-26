#!/usr/bin/env python3
"""
run_allocation.py - wrapper to call process_data.run_pipeline
"""
import argparse, sys, os
from process_data import run_pipeline

def main():
    parser = argparse.ArgumentParser(description="Run allocation pipeline")
    parser.add_argument("--input", "-i", required=True, help="Path to raw ECW CSV/XLSX")
    parser.add_argument("--wb", "-w", required=False, help="Path to workbook (Audentes_Verification1.xlsm) containing Help sheet")
    parser.add_argument("--outdir", "-o", default="outputs", help="Output directory")
    args = parser.parse_args()

    inp = args.input
    wb = args.wb
    outdir = args.outdir
    os.makedirs(outdir, exist_ok=True)

    try:
        res = run_pipeline(inp, wb, outdir)
        print("Processing finished. Outputs:")
        print("HX CSV:", res["hx_csv"])
        print("Warnings:", res["warnings"])
        print("Allocation debug:", res["debug"])
    except Exception as e:
        print("Error:", e)
        raise

if __name__ == "__main__":
    main()
