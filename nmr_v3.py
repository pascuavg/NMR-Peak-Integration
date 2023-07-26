###   This is a working copy
###   Multiple spectra for integrating for area and calculating concentration
###   Zipped files + directory folders
###   Single excel output with concentration, area, and peak limits
###   Multiple spectra for binning
###   Spectra overlay with peak limits and identities

import nmrglue as ng
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import zipfile
import shutil
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

def find_pdata_directories(root_dir):
    pdata_dirs = []
    for dirpath, _, filenames in os.walk(root_dir):
        if "/pdata/1" in dirpath:
            pdata_dirs.append(dirpath)
    return pdata_dirs

def process_selected_dirs_concentration():
    global selected_pdata_dirs, tsp_concentration

    if not selected_pdata_dirs:
        messagebox.showerror("Error", "Please select a data directory.")
        return

    try:
        tsp_concentration = float(concentration_entry.get())
    except ValueError:
        messagebox.showerror("Error", "Invalid concentration value. Please enter a valid number.")
        return

    results_concentration = []
    results_area = []
    peak_identities = set()

    for root_dir in selected_pdata_dirs:
        pdata_dirs = find_pdata_directories(root_dir)
        if not pdata_dirs:
            messagebox.showwarning("No pdata/1 Directories", f"No pdata/1 directories found in '{root_dir}'. Skipping.")
            continue

        for pdata_dir in pdata_dirs:
            try:
                dic, data = ng.bruker.read_pdata(pdata_dir, scale_data=True)
            except OSError:
                continue

            udic = ng.bruker.guess_udic(dic, data)
            uc = ng.fileiobase.uc_from_udic(udic)
            ppm_scale = uc.ppm_scale()

            peak_data = pd.read_excel("peak_limits.xlsx")

            ref_name = peak_data.at[0, 'Peak identity']
            ref_start = peak_data.at[0, 'ppm start']
            ref_end = peak_data.at[0, 'ppm end']
            ref_num_protons = peak_data.at[0, '# protons']

            ref_min_index = np.abs(ppm_scale - ref_start).argmin()
            ref_max_index = np.abs(ppm_scale - ref_end).argmin()
            if ref_min_index > ref_max_index:
                ref_min_index, ref_max_index = ref_max_index, ref_min_index

            ref_peak = data[ref_min_index:ref_max_index + 1]
            ref_area = ref_peak.sum()

            for index, row in peak_data.iloc[1:].iterrows():
                name = row['Peak identity']
                peak_identities.add(name)
                start = row['ppm start']
                end = row['ppm end']
                num_protons_peak = row['# protons']

                min_index = np.abs(ppm_scale - start).argmin()
                max_index = np.abs(ppm_scale - end).argmin()
                if min_index > max_index:
                    min_index, max_index = max_index, min_index

                peak = data[min_index:max_index + 1]
                peak_area = peak.sum()

                concentration = (peak_area / ref_area) * tsp_concentration * ref_num_protons / num_protons_peak

                results_concentration.append({'Name': name, 'Concentration': concentration, 'Parent File Path': pdata_dir})
                results_area.append({'Name': name, 'Area': peak_area, 'Parent File Path': pdata_dir})

    if not results_concentration:
        messagebox.showerror("No Data", "No pdata/1 directories were processed.")
        return

    concentrations_df = pd.DataFrame(results_concentration)
    concentrations_df = concentrations_df.pivot(index='Parent File Path', columns='Name', values='Concentration')

    areas_df = pd.DataFrame(results_area)
    areas_df = areas_df.pivot(index='Parent File Path', columns='Name', values='Area')

    with pd.ExcelWriter('nmr_analysis_results.xlsx') as writer:
        concentrations_df.to_excel(writer, sheet_name='Concentrations')
        areas_df.to_excel(writer, sheet_name='Areas')

    messagebox.showinfo("Processing Complete", "Check the results in 'nmr_analysis_results.xlsx'")

    # Clean up temporary directories
    for dir in selected_pdata_dirs:
        if 'temp' in dir:
            shutil.rmtree(dir)

def process_selected_dirs_binning():
    global selected_pdata_dirs

    if not selected_pdata_dirs:
        messagebox.showerror("Error", "Please select a data directory.")
        return

    try:
        binning_step_size = float(binning_entry.get())
    except ValueError:
        messagebox.showerror("Error", "Invalid binning step size. Please enter a valid number.")
        return

    all_bins = []
    all_ppm_scales = []

    for root_dir in selected_pdata_dirs:
        pdata_dirs = find_pdata_directories(root_dir)
        if not pdata_dirs:
            messagebox.showwarning("No pdata/1 Directories", f"No pdata/1 directories found in '{root_dir}'. Skipping.")
            continue

        for pdata_dir in pdata_dirs:
            try:
                dic, data = ng.bruker.read_pdata(pdata_dir, scale_data=True)
            except OSError:
                continue

            udic = ng.bruker.guess_udic(dic, data)
            uc = ng.fileiobase.uc_from_udic(udic)
            ppm_scale = uc.ppm_scale()
            all_ppm_scales.append(ppm_scale)
            
    min_ppm = min([ppm.min() for ppm in all_ppm_scales])
    max_ppm = max([ppm.max() for ppm in all_ppm_scales])

    bin_edges = np.arange(min_ppm, max_ppm, binning_step_size)

    for ppm_scale in all_ppm_scales:
        bin_indices = np.digitize(ppm_scale, bin_edges)
        binned_spectrum = [data[bin_indices == i].sum() for i in range(1, len(bin_edges))]

        all_bins.append(binned_spectrum)

    #binning_df = pd.DataFrame(all_bins, columns=bin_edges[:-1], index=selected_pdata_dirs)
    binning_df = pd.DataFrame(all_bins, columns=bin_edges[:-1], index=[os.path.dirname(pdata_dir) for pdata_dir in pdata_dirs])

    binning_df = binning_df.T

    with pd.ExcelWriter('nmr_binning_results.xlsx') as writer:
        binning_df.to_excel(writer, sheet_name='Binning')

    messagebox.showinfo("Binning Complete", "Check the results in 'nmr_binning_results.xlsx'")

    # Clean up temporary directories
    for dir in selected_pdata_dirs:
        if 'temp' in dir:
            shutil.rmtree(dir)

def browse_directory():
    global selected_pdata_dirs
    file_or_dir = messagebox.askquestion("Zipped File or Directory", "Are you selecting a zipped file?", icon='question')
    if file_or_dir == 'yes': # If the user chooses 'yes', we'll open the file selection dialog.
        root_dir = filedialog.askopenfilename(title="Select the zipped file containing processed Bruker NMR data")
        if root_dir.endswith('.zip'):
            with zipfile.ZipFile(root_dir, 'r') as zip_ref:
                zip_ref.extractall('Zipped file')
                root_dir = 'Zipped file'
    else: # If the user chooses 'no', we'll open the directory selection dialog.
        root_dir = filedialog.askdirectory(title="Select the root directory containing processed Bruker NMR data")
    selected_pdata_dirs = [root_dir]

    peak_limits = pd.read_excel("peak_limits.xlsx")

    spectra = []
    for root_dir in selected_pdata_dirs:
        pdata_dirs = find_pdata_directories(root_dir)
        for pdata_dir in pdata_dirs:
            try:
                dic, data = ng.bruker.read_pdata(pdata_dir, scale_data=True)
                udic = ng.bruker.guess_udic(dic, data)
                uc = ng.fileiobase.uc_from_udic(udic)
                ppm_scale = uc.ppm_scale()
                spectra.append((ppm_scale, data))
            except OSError:
                continue
    selected_dirs_label.config(text=f"Selected Directory: {root_dir}")

    plot_spectra(spectra,peak_limits)

def plot_spectra(spectra, peak_data):
    for ax in fig.axes:
        ax.cla()  # clear the plot
    ax = fig.add_subplot(111)
    for ppm_scale, data in spectra:
        ax.plot(ppm_scale, data)
        for index, row in peak_data.iterrows():
            name = row['Peak identity']
            start = row['ppm start']
            end = row['ppm end']
            ax.plot([start, end], [0, 0], color='black', linewidth=3)  # Draw line for integration region
            mid_point = (start + end) / 2  # Calculate midpoint for label placement
            ax.text(mid_point, -5.0, name, ha='center', va='top', color='black', fontsize=10, rotation=315)
    ax.set_xlabel("ppm")
    ax.set_ylabel("Intensity")
    ax.invert_xaxis()
    canvas.draw()



selected_pdata_dirs = []
tsp_concentration = 0
binning_step = 0

root = tk.Tk()
root.title("Bruker NMR Plasma Metabolism Analysis")

selected_dirs_label = tk.Label(root, text="Selected Directory: None")
selected_dirs_label.pack(pady=10)

dirs_button = tk.Button(root, text="Browse Data Directory or Zipped File", command=browse_directory)
dirs_button.pack()

concentration_label = tk.Label(root, text="Enter Reference TSP Concentration:")
concentration_label.pack(pady=5)

concentration_entry = tk.Entry(root)
concentration_entry.pack()

process_concentration_button = tk.Button(root, text="Process Concentration and Area", command=process_selected_dirs_concentration)
process_concentration_button.pack(pady=10)

binning_label = tk.Label(root, text="Enter Binning Step Size in ppm (Optional):")
binning_label.pack(pady=5)

binning_entry = tk.Entry(root)
binning_entry.pack()

process_binning_button = tk.Button(root, text="Process Binning", command=process_selected_dirs_binning)
process_binning_button.pack(pady=10)

fig = plt.Figure(figsize=(5, 5), dpi=100)
canvas = FigureCanvasTkAgg(fig, master=root)
canvas.get_tk_widget().pack()

toolbar = NavigationToolbar2Tk(canvas, root)
toolbar.update()
canvas.get_tk_widget().pack()

root.mainloop()