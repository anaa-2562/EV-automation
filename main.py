import os
import json
import shutil
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime

# Import processing and upload modules
try:
    import pandas as pd
    from macro import audentes_verification_cleaned
    from process_data import run_pipeline
    from upload_hx import hx_upload
except ImportError as e:
    print(f"Warning: Could not import modules: {e}")


def _ensure_dirs(cfg: dict) -> tuple[str, str, str]:
    """Create input, output, and log directories if they don't exist."""
    base = os.getcwd()
    in_dir = os.path.join(base, cfg.get("input_folder", "inputs"))
    out_dir = os.path.join(base, cfg.get("output_folder", "outputs"))
    log_dir = os.path.join(base, cfg.get("log_folder", "logs"))
    for d in (in_dir, out_dir, log_dir):
        os.makedirs(d, exist_ok=True)
    return in_dir, out_dir, log_dir


def _load_config() -> dict:
    """Load configuration from config.json."""
    cfg_path = os.path.join(os.getcwd(), "config.json")
    try:
        with open(cfg_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


_current_log_path = None

def _log(message: str, log_path: str = None) -> None:
    """Write message to log file with timestamp."""
    global _current_log_path
    if log_path:
        _current_log_path = log_path
    elif _current_log_path is None:
        cfg = _load_config()
        _, _, log_dir = _ensure_dirs(cfg)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        _current_log_path = os.path.join(log_dir, f"run_{ts}.txt")
    with open(_current_log_path, "a", encoding="utf-8") as f:
        log_ts = datetime.now().strftime("%H:%M:%S")
        f.write(f"[{log_ts}] {message}\n")


def run_process_async(ecw_path: str, template_path: str, escalation_path: str, status_label: tk.Label, root: tk.Tk) -> None:
    """Run the complete automation workflow in a separate thread."""
    try:
        cfg = _load_config()
        in_dir, out_dir, log_dir = _ensure_dirs(cfg)

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        run_log_path = os.path.join(log_dir, f"run_{ts}.txt")
        _log("=" * 60, log_path=run_log_path)
        _log("Audentes Verification Automation Tool - Run Started", log_path=run_log_path)
        _log("=" * 60, log_path=run_log_path)

        # Step 1: Save files to inputs/
        status_label.config(text="Saving files to inputs/ folder...")
        root.update()
        ecw_copy = os.path.join(in_dir, f"eCW_{ts}{os.path.splitext(ecw_path)[1]}")
        tpl_copy = os.path.join(in_dir, f"Template_{ts}{os.path.splitext(template_path)[1]}")
        shutil.copy2(ecw_path, ecw_copy)
        shutil.copy2(template_path, tpl_copy)
        _log(f"eCW file loaded: {os.path.basename(ecw_copy)}", log_path=run_log_path)
        _log(f"Template file loaded: {os.path.basename(tpl_copy)}", log_path=run_log_path)

        escalation_copy = None
        if escalation_path and os.path.isfile(escalation_path):
            try:
                esc_ext = os.path.splitext(escalation_path)[1] or ".csv"
                escalation_copy = os.path.join(in_dir, f"Escalation_{ts}{esc_ext}")
                shutil.copy2(escalation_path, escalation_copy)
                _log(f"Escalation tracker loaded: {os.path.basename(escalation_copy)}", log_path=run_log_path)
            except Exception as e:
                escalation_copy = None
                _log(f"Warning: Could not copy escalation tracker: {e}", log_path=run_log_path)

        # === NEW STEP 2: Run Macro Cleaning ===
        try:
            status_label.config(text="Running VBA macro cleanup step...")
            root.update()
            cleaned_file = os.path.join(out_dir, f"Audentes_Verification_Cleaned_{ts}.xlsx")
            audentes_verification_cleaned(ecw_copy, tpl_copy, cleaned_file)
            _log(f"Macro cleanup complete: {os.path.basename(cleaned_file)}", log_path=run_log_path)
        except Exception as e:
            error_msg = f"Macro cleanup failed: {str(e)}"
            _log(f"ERROR: {error_msg}", log_path=run_log_path)
            messagebox.showerror("Macro Step Error", f"❌ {error_msg}\n\nCheck logs for details.")
            status_label.config(text="Macro cleanup failed. See logs.")
            return

        # === STEP 2.5: Process Data (All filters now handled in run_pipeline) ===
        status_label.config(text="Processing cleaned data (Filters + Mapping + Allocation)...")
        root.update()
        try:
            # All filtering (Visit Status, WC, Escalation) is now handled inside run_pipeline
            escalation_input = escalation_copy or (escalation_path if escalation_path and os.path.isfile(escalation_path) else None)
            if escalation_input:
                _log(f"Using escalation tracker: {escalation_input}", log_path=run_log_path)
            else:
                _log("No escalation tracker provided; skipping escalation filter", log_path=run_log_path)
            
            pipeline_result = run_pipeline(cleaned_file, tpl_copy, out_dir, escalation_file_path=escalation_input, log_path=run_log_path)
            output_path = pipeline_result["hx_csv"]
            processed_count = pipeline_result.get("processed_count", 0)
            _log(f"Data processed successfully. {processed_count} records remaining.", log_path=run_log_path)
            _log(f"HX output saved: {os.path.basename(output_path)}", log_path=run_log_path)
        except Exception as e:
            error_msg = f"Data processing failed: {str(e)}"
            _log(f"ERROR: {error_msg}", log_path=run_log_path)
            messagebox.showerror("Processing Error", f"❌ {error_msg}\n\nCheck logs for details.")
            status_label.config(text="Processing failed. See logs.")
            return

        # === STEP 4: Upload to HX ===
        status_label.config(text="Uploading to HealthX portal...")
        root.update()
        try:
            success, upload_message = hx_upload(output_path, log_path=run_log_path)
            if success:
                _log("Upload to HealthX successful.", log_path=run_log_path)
                status_label.config(text="✅ Process Complete! Upload successful.")
                messagebox.showinfo(
                    "Audentes Automation Tool",
                    f"✅ Process Complete!\n\n"
                    f"• Processed {processed_count} records\n"
                    f"• HX File: {os.path.basename(output_path)}\n"
                    f"• Upload: Successful\n\n"
                    f"Check logs folder for details."
                )
            else:
                _log(f"Upload failed: {upload_message}", log_path=run_log_path)
                status_label.config(text="⚠️ Processed successfully, but upload failed.")
                messagebox.showwarning(
                    "Upload Warning",
                    f"⚠️ Processed successfully, but upload failed:\n\n{upload_message}\n\n"
                    f"File saved at:\n{output_path}\n\nCheck logs."
                )
        except Exception as e:
            error_msg = f"Upload failed: {str(e)}"
            _log(f"ERROR: {error_msg}", log_path=run_log_path)
            status_label.config(text="⚠️ Processed successfully, but upload failed.")
            messagebox.showwarning(
                "Upload Error",
                f"⚠️ Processed successfully, but upload failed:\n\n{error_msg}\n\n"
                f"File saved at:\n{output_path}\n\nCheck logs."
            )

    except Exception as exc:
        error_msg = f"Unexpected error: {str(exc)}"
        _log(f"FATAL ERROR: {error_msg}", log_path=run_log_path if 'run_log_path' in locals() else None)
        messagebox.showerror("Audentes Automation Tool", f"❌ {error_msg}\n\nSee logs folder for details.")
        status_label.config(text="Error occurred. See logs.")


def on_run_click(ecw_var: tk.StringVar, tpl_var: tk.StringVar, esc_var: tk.StringVar, status_label: tk.Label, root: tk.Tk) -> None:
    """Validate inputs and start automation."""
    ecw_path = ecw_var.get().strip()
    template_path = tpl_var.get().strip()
    if not ecw_path or not os.path.isfile(ecw_path):
        messagebox.showwarning("Input Required", "Please select a valid eCW report (.xlsx/.xlsm/.csv).")
        return
    if not template_path or not os.path.isfile(template_path):
        messagebox.showwarning("Input Required", "Please select a valid template (.xlsx/.xlsm).")
        return
    escalation_path = esc_var.get().strip()
    if escalation_path and not os.path.isfile(escalation_path):
        messagebox.showwarning("Input Warning", "Escalation tracker path is invalid. Please re-select or leave blank.")
        return

    status_label.config(text="Starting process...")
    t = threading.Thread(target=run_process_async, args=(ecw_path, template_path, escalation_path, status_label, root), daemon=True)
    t.start()


def build_gui() -> None:
    """Build and display the main GUI."""
    base_dir = os.path.dirname(os.path.abspath(getattr(sys, 'frozen', False) and sys.executable or __file__))
    os.chdir(base_dir)
    cfg = _load_config()
    _ensure_dirs(cfg)

    root = tk.Tk()
    root.title("Audentes Verification Automation")
    root.geometry("700x300")

    header = tk.Label(root, text="Audentes Verification Automation Tool", font=("Arial", 14, "bold"))
    header.grid(row=0, column=0, columnspan=3, padx=10, pady=15)

    ecw_var = tk.StringVar()
    tpl_var = tk.StringVar()
    esc_var = tk.StringVar()

    def browse_ecw():
        path = filedialog.askopenfilename(title="Select eCW Report", filetypes=[("Data files", "*.xlsx *.xls *.xlsm *.csv")])
        if path:
            ecw_var.set(path)
            status_label.config(text=f"Selected: {os.path.basename(path)}")

    def browse_tpl():
        path = filedialog.askopenfilename(title="Select Template/Macro File", filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")])
        if path:
            tpl_var.set(path)
            status_label.config(text=f"Selected: {os.path.basename(path)}")

    def browse_esc():
        path = filedialog.askopenfilename(title="Select Escalation Tracker", filetypes=[("Data files", "*.csv *.xlsx *.xlsm"), ("CSV files", "*.csv"), ("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")])
        if path:
            esc_var.set(path)
            status_label.config(text=f"Selected escalation file: {os.path.basename(path)}")

    tk.Label(root, text="eCW Report (.xlsx/.xlsm/.csv):").grid(row=1, column=0, padx=10, pady=10, sticky="w")
    tk.Entry(root, textvariable=ecw_var, width=60).grid(row=1, column=1, padx=10, pady=10, sticky="ew")
    tk.Button(root, text="Select eCW Report", command=browse_ecw).grid(row=1, column=2, padx=10, pady=10)

    tk.Label(root, text="Template/Macro File (.xlsx/.xlsm):").grid(row=2, column=0, padx=10, pady=10, sticky="w")
    tk.Entry(root, textvariable=tpl_var, width=60).grid(row=2, column=1, padx=10, pady=10, sticky="ew")
    tk.Button(root, text="Select Template/Macro File", command=browse_tpl).grid(row=2, column=2, padx=10, pady=10)

    tk.Label(root, text="Escalation Tracker (.csv/.xlsx/.xlsm, optional):").grid(row=3, column=0, padx=10, pady=10, sticky="w")
    tk.Entry(root, textvariable=esc_var, width=60).grid(row=3, column=1, padx=10, pady=10, sticky="ew")
    tk.Button(root, text="Select Escalation Tracker", command=browse_esc).grid(row=3, column=2, padx=10, pady=10)

    root.columnconfigure(1, weight=1)
    status_label = tk.Label(root, text="Ready - Select files and click 'Run Process'", wraplength=600, justify="left")
    status_label.grid(row=4, column=0, columnspan=3, padx=10, pady=10, sticky="w")

    tk.Button(
        root,
        text="Run Process",
        command=lambda: on_run_click(ecw_var, tpl_var, esc_var, status_label, root),
        font=("Arial", 11, "bold"),
        bg="#4CAF50",
        fg="white",
        padx=20,
        pady=5
    ).grid(row=5, column=0, columnspan=3, padx=10, pady=15)

    root.mainloop()


if __name__ == "__main__":
    build_gui()
