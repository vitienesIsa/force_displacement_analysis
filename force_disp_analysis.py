
##############################################################################
#
# Force-displacement curve analysis
#
# Developed by Isabela Vitienes as part of the publication 
#   Vitienes I.*, Ross E.* et al., Bone, 2026
#
# Iterates through per-measurement CSV data, plots each force-displacement
# curve, and prompts user to define (by clicking in plot) the following 
# points:
#   1. x_0 - start of loading
#   2. x_yield - yield point, i.e. transition from elastic to plastic 
#      deformation
#   3. x_ult - ultimate load, i.e. peak load
#   4. x_fail - fracture/failure, i.e. when load drops to/near zero
#
# X-coordinates are recorded and correspoinding y-values are extracted from 
# csv.
#
# Results are updated and saved to an excel file after each plot.
# Therefore, progress is not lost if script is interrupted. 
# Already-processed measurements are skipped. 
#
#
# INPUT: root path to folder containing measurements
# OUTPUT: .xlsx sheet stored in root path with x- and y-coordinates of
#         selected points, and derived outcome measures:
#            - Yield force (N)
#            - Ultimate force (N)
#            - Failure force (N)
#            - Stiffness (N/mm) 
#            - Post-yield displacement (mm)
#            - Work-to-failure (N*mm)
#            - Sampling rate (Hz)
#            - Measurement duration (s)
#
# **Note: 
#   Script assumes that data is stored in root path as follows:
#   
#   Root path 
#       |
#       ---Folder, date of measurement
#               |
#               ---Folder, specimen-specific measurements
#                       |
#                       --- .csv file
#
##############################################################################




import os
import re
import glob
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("TkAgg")   # change to "Qt5Agg" if TkAgg is not available in Spyder
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from scipy.integrate import trapezoid
import tkinter as tk
from tkinter import filedialog



# Define inputs and outputs

ROOT_DIR   = None
OUTPUT_XLS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "manual_annotations.xlsx")

COL_TIME      = "Time"
COL_EXTENSION = "Extension"
COL_LOAD      = "Load"

POINT_LABELS  = ["x_0", "x_yield", "x_ult", "x_fail"]
POINT_COLORS  = ["green", "orange", "blue", "red"]


COLUMNS = [
    "date", "specimen", "file_number", "file_path",
    "x_0_mm",     "F_at_x0_N",
    "x_yield_mm", "F_yield_N",
    "x_ult_mm",   "F_ult_N",
    "x_fail_mm",  "F_fail_N",
]

OUTCOME_COLUMNS = [
    "date", "specimen", "file_number", "file_path",
    "F ult adjusted",
    "Sampling rate (Hz)",
    "Total time (s)",
    "Number of data points",
    "Estimated number of points",
    "Stiffness (N/mm)",
    "Post-yield displacement (mm)",
    "Work-to-failure (N*mm)",
    "Processing status",
]


# Functions

def init_workbook(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Annotations"
    header_fill = PatternFill("solid", fgColor="4F81BD")
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    for col_idx, col_name in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        ws.column_dimensions[cell.column_letter].width = max(14, len(col_name) + 2)
    outcome_ws = wb.create_sheet("Outcomes")
    for col_idx, col_name in enumerate(OUTCOME_COLUMNS, start=1):
        cell = outcome_ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        outcome_ws.column_dimensions[cell.column_letter].width = max(14, len(col_name) + 2)
    wb.save(path)
    return wb


def load_existing(path):
    if not os.path.exists(path):
        return pd.DataFrame(columns=COLUMNS)
    try:
        df = pd.read_excel(path, sheet_name="Annotations")
        return df
    except Exception:
        return pd.DataFrame(columns=COLUMNS)


def append_row_to_excel(path, row_dict):
    if not os.path.exists(path):
        init_workbook(path)
    wb = openpyxl.load_workbook(path)
    ws = wb["Annotations"]
    next_row = ws.max_row + 1
    data_font = Font(name="Arial", size=11)
    for col_idx, col_name in enumerate(COLUMNS, start=1):
        val = row_dict.get(col_name, "")
        cell = ws.cell(row=next_row, column=col_idx, value=val)
        cell.font = data_font
        # Alternate row shading
        if next_row % 2 == 0:
            cell.fill = PatternFill("solid", fgColor="DCE6F1")
    wb.save(path)


def append_outcome_to_excel(path, row_dict):
    if not os.path.exists(path):
        init_workbook(path)
    wb = openpyxl.load_workbook(path)
    if "Outcomes" not in wb.sheetnames:
        ws = wb.create_sheet("Outcomes")
        header_fill = PatternFill("solid", fgColor="4F81BD")
        header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        for col_idx, col_name in enumerate(OUTCOME_COLUMNS, start=1):
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")
            ws.column_dimensions[cell.column_letter].width = max(14, len(col_name) + 2)
    ws = wb["Outcomes"]
    next_row = ws.max_row + 1
    data_font = Font(name="Arial", size=11)
    for col_idx, col_name in enumerate(OUTCOME_COLUMNS, start=1):
        val = row_dict.get(col_name, "")
        cell = ws.cell(row=next_row, column=col_idx, value=val)
        cell.font = data_font
        if next_row % 2 == 0:
            cell.fill = PatternFill("solid", fgColor="DCE6F1")
    wb.save(path)


def find_csvs(root):
    date_folders = sorted(
        f for f in glob.glob(os.path.join(root, "*")) if os.path.isdir(f)
    )
    for date_folder in date_folders:
        date_name = os.path.basename(date_folder)
        specimen_folders = sorted(
            f for f in glob.glob(os.path.join(date_folder, "*")) if os.path.isdir(f)
        )
        for spec_folder in specimen_folders:
            spec_name = os.path.basename(spec_folder)
            for csv_path in sorted(glob.glob(os.path.join(spec_folder, "*.csv"))):
                stem  = os.path.splitext(os.path.basename(csv_path))[0]
                match = re.search(r"(\d)$", stem)
                file_num = match.group(1) if match else "?"
                yield date_name, spec_name, file_num, csv_path


def load_csv(path):
    df = pd.read_csv(path, header=0, skiprows=[1])
    df = df.rename(columns=lambda c: c.strip())
    for col in (COL_TIME, COL_EXTENSION, COL_LOAD):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    df = df[[COL_TIME, COL_EXTENSION, COL_LOAD]].dropna()
    df = df.sort_values(COL_TIME).reset_index(drop=True)
    return df


def nearest_point(ext, load, x_clicked):
    idx = int(np.argmin(np.abs(ext - x_clicked)))
    return float(ext[idx]), float(load[idx])


def calculate_sampling_rate(curve):

    time = curve["Time"].to_numpy(
        dtype=float
    )

    dt = np.diff(time)

    # Ignore zero or negative intervals
    dt = dt[dt > 0]

    if len(dt) == 0:
        return np.nan

    median_dt = np.median(dt)

    if median_dt <= 0:
        return np.nan

    return 1.0 / median_dt


def calculate_total_time(curve):

    time = curve["Time"].to_numpy(
        dtype=float
    )

    return time[-1] - time[0]


def calculate_estimated_number_of_points(
    sampling_rate,
    total_time
):

    if (
        np.isnan(sampling_rate)
        or
        np.isnan(total_time)
    ):
        return np.nan

    return (
        sampling_rate
        *
        total_time
        +
        1
    )


def calculate_work_to_failure(
    curve,
    x0,
    x_fail
):

    x = curve["Displacement"].to_numpy(
        dtype=float
    )

    F = curve["Force"].to_numpy(
        dtype=float
    )

    # --------------------------------------------------------
    # Check displacement range
    # --------------------------------------------------------

    x_min = np.min(x)
    x_max = np.max(x)

    if x0 < x_min or x0 > x_max:

        raise ValueError(
            f"x_0 = {x0} mm is outside the measured "
            f"displacement range "
            f"[{x_min}, {x_max}] mm."
        )

    if x_fail < x_min or x_fail > x_max:

        raise ValueError(
            f"x_fail = {x_fail} mm is outside the measured "
            f"displacement range "
            f"[{x_min}, {x_max}] mm."
        )

    if x_fail <= x0:

        raise ValueError(
            f"x_fail ({x_fail}) must be greater than "
            f"x_0 ({x0})."
        )

    # --------------------------------------------------------
    # Sort by displacement
    # --------------------------------------------------------

    sort_index = np.argsort(x)

    x_sorted = x[sort_index]

    F_sorted = F[sort_index]

    # --------------------------------------------------------
    # Remove duplicate displacement values
    # --------------------------------------------------------

    unique_mask = np.concatenate(
        (
            [True],
            np.diff(x_sorted) != 0
        )
    )

    x_sorted = x_sorted[unique_mask]

    F_sorted = F_sorted[unique_mask]

    # --------------------------------------------------------
    # Points between x0 and x_fail
    # --------------------------------------------------------

    mask = (
        (x_sorted > x0)
        &
        (x_sorted < x_fail)
    )

    x_middle = x_sorted[mask]

    F_middle = F_sorted[mask]

    # --------------------------------------------------------
    # Force at x0
    # --------------------------------------------------------

    F_at_x0 = np.interp(
        x0,
        x_sorted,
        F_sorted
    )

    # --------------------------------------------------------
    # Force at x_fail
    # --------------------------------------------------------

    F_at_xfail = np.interp(
        x_fail,
        x_sorted,
        F_sorted
    )

    # --------------------------------------------------------
    # Integration arrays
    # --------------------------------------------------------

    x_integrate = np.concatenate(
        (
            [x0],
            x_middle,
            [x_fail]
        )
    )

    F_integrate = np.concatenate(
        (
            [F_at_x0],
            F_middle,
            [F_at_xfail]
        )
    )

    # --------------------------------------------------------
    # Numerical integration
    #
    # N × mm
    # --------------------------------------------------------

    return trapezoid(
        F_integrate,
        x_integrate
    )


def calculate_outcomes_from_curve(curve, x0, f0, x_yield, f_yield, f_ult, f_fail, x_fail):
    curve = curve.rename(columns={
        COL_EXTENSION: "Displacement",
        COL_LOAD: "Force",
    })

    F_ult_adjusted = max(
        f_ult,
        f_fail
    )

    if x_yield == x0:
        raise ValueError(
            "x_yield_mm equals x_0_mm."
        )

    stiffness = (
        f_yield - f0
    ) / (
        x_yield - x0
    )

    post_yield = (
        x_fail
        -
        x_yield
    )

    actual_points = len(curve)
    sampling_rate = calculate_sampling_rate(curve)
    total_time = calculate_total_time(curve)
    estimated_points = calculate_estimated_number_of_points(
        sampling_rate,
        total_time
    )
    work = calculate_work_to_failure(
        curve,
        x0,
        x_fail
    )

    return {
        "F ult adjusted": F_ult_adjusted,
        "Sampling rate (Hz)": sampling_rate,
        "Total time (s)": total_time,
        "Number of data points": actual_points,
        "Estimated number of points": estimated_points,
        "Stiffness (N/mm)": stiffness,
        "Post-yield displacement (mm)": post_yield,
        "Work-to-failure (N*mm)": work,
        "Processing status": "OK",
    }


def annotate_specimen(date_name, spec_name, file_num, csv_path):
    df  = load_csv(csv_path)
    ext  = df[COL_EXTENSION].to_numpy()
    load = df[COL_LOAD].to_numpy()

    fig, ax = plt.subplots(figsize=(9, 5))
    ax.plot(ext, load, color="gray", lw=1)
    ax.set_xlabel("Extension (mm)")
    ax.set_ylabel("Load (N)")
    ax.set_title(
        f"{date_name} / {spec_name} (file {file_num})\n"
        f"Click 4 points in order: {', '.join(POINT_LABELS)}"
    )
    fig.tight_layout()

    clicks    = []
    markers   = []
    label_txt = []

    def on_click(event):
        if event.inaxes != ax or event.button != 1:
            return
        if len(clicks) >= 4:
            return
        x_c, _ = event.xdata, event.ydata
        x_data, y_data = nearest_point(ext, load, x_c)
        clicks.append((x_data, y_data))

        label = POINT_LABELS[len(clicks) - 1]
        color = POINT_COLORS[len(clicks) - 1]
        m, = ax.plot(x_data, y_data, "o", color=color, ms=10, zorder=5)
        t  = ax.text(x_data, y_data, f"  {label}\n  ({x_data:.2f}, {y_data:.2f})",
                     fontsize=8, color=color, va="bottom")
        markers.append(m)
        label_txt.append(t)
        fig.canvas.draw()

        if len(clicks) == 4:
            ax.set_title(
                f"{date_name} / {spec_name} (file {file_num})\n"
                "All 4 points selected — close window to confirm",
                color="darkgreen"
            )
            fig.canvas.draw()

    fig.canvas.mpl_connect("button_press_event", on_click)
    plt.show(block=True)   # blocks until window is closed

    if len(clicks) < 4:
        print(f"  WARNING: only {len(clicks)} points selected for "
              f"{date_name}/{spec_name} file {file_num} — skipping.")
        return None

    (x0, f0), (xy, fy), (xu, fu), (xf, ff) = clicks

    try:
        outcomes = calculate_outcomes_from_curve(
            df,
            x0, f0,
            xy, fy,
            xu, fu,
            ff, xf
        )
    except Exception as error:
        print(f"  ERROR calculating outcomes: {error}")
        outcomes = {
            column: ""
            for column in OUTCOME_COLUMNS
            if column not in ("date", "specimen", "file_number", "file_path", "Processing status")
        }
        outcomes["Processing status"] = f"ERROR: {error}"

    result = {
        "date":        date_name,
        "specimen":    spec_name,
        "file_number": file_num,
        "file_path":   csv_path,
        "x_0_mm":      x0,   "F_at_x0_N":  f0,
        "x_yield_mm":  xy,   "F_yield_N":   fy,
        "x_ult_mm":    xu,   "F_ult_N":     fu,
        "x_fail_mm":   xf,   "F_fail_N":    ff,
    }

    outcome_row = {
        "date": date_name,
        "specimen": spec_name,
        "file_number": file_num,
        "file_path": csv_path,
    }
    outcome_row.update(outcomes)
    result["outcomes"] = outcome_row

    return result






# Main

def main():
    root = tk.Tk()
    root.withdraw()
    selected_root = filedialog.askdirectory(
        title="Select root folder containing measurements"
    )
    root.destroy()

    if not selected_root:
        print("No input directory selected.")
        return False

    global ROOT_DIR
    ROOT_DIR = selected_root

    # Initialise workbook if it doesn't exist yet
    if not os.path.exists(OUTPUT_XLS):
        init_workbook(OUTPUT_XLS)
        print(f"Created new output file: {OUTPUT_XLS}")
    else:
        print(f"Resuming — appending to: {OUTPUT_XLS}")
        wb = openpyxl.load_workbook(OUTPUT_XLS)
        if "Outcomes" not in wb.sheetnames:
            ws = wb.create_sheet("Outcomes")
            header_fill = PatternFill("solid", fgColor="4F81BD")
            header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
            for col_idx, col_name in enumerate(OUTCOME_COLUMNS, start=1):
                cell = ws.cell(row=1, column=col_idx, value=col_name)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center")
                ws.column_dimensions[cell.column_letter].width = max(14, len(col_name) + 2)
            wb.save(OUTPUT_XLS)

    # Load already-processed specimens so we can skip them
    existing = load_existing(OUTPUT_XLS)
    done_keys = set(
        zip(
            existing["date"].astype(str),
            existing["specimen"].astype(str),
            existing["file_number"].astype(str),
        )
    ) if len(existing) else set()

    all_csvs = list(find_csvs(ROOT_DIR))
    total    = len(all_csvs)

    for i, (date_name, spec_name, file_num, csv_path) in enumerate(all_csvs, start=1):
        key = (str(date_name), str(spec_name), str(file_num))
        if key in done_keys:
            continue

        result = annotate_specimen(date_name, spec_name, file_num, csv_path)

        if result is not None:
            append_row_to_excel(OUTPUT_XLS, result)
            append_outcome_to_excel(OUTPUT_XLS, result["outcomes"])
            
    print(f"\nAll done. Results saved to:\n  {OUTPUT_XLS}")




if __name__ == "__main__":
    main()
