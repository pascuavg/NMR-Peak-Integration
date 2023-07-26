###   This is a working copy
###   Multiple spectra for integrating for area and calculating concentration
###   Zipped files + directory folders
###   Single excel output with concentration, area, and peak limits

import nmrglue as ng
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import zipfile
import shutil


def find_pdata_directories(root_dir):
    pdata_dirs = []
    for dirpath, _, filenames in os.walk(root_dir):
        if "/pdata/1" in dirpath:
            pdata_dirs.append(dirpath)
    return pdata_dirs

def process_selected_dirs():
    global selected_pdata_dirs, tsp_concentration

    if not selected_pdata_dirs:
        messagebox.showerror("Error", "Please select a data directory.")
        return

    try:
        tsp_concentration = float(concentration_entry.get())
    except ValueError:
        messagebox.showerror("Error", "Invalid concentration value. Please enter a valid number.")
        return

    results = []

    peak_identities = set()

    for root_dir in selected_pdata_dirs:
        pdata_dirs = find_pdata_directories(root_dir)
        if not pdata_dirs:
            messagebox.showwarning("No pdata/1 Directories", f"No pdata/1 directories found in '{root_dir}'. Skipping.")
            continue

        pdata_results = {}
        for pdata_dir in pdata_dirs:
            try:
                dic, data = ng.bruker.read_pdata(pdata_dir, scale_data=True)
            except OSError:
                continue

            udic = ng.bruker.guess_udic(dic, data)
            uc = ng.fileiobase.uc_from_udic(udic)
            ppm_scale = uc.ppm_scale()

            peak_data = pd.read_excel("peak_limits.xlsx")

            concentration_data = []
            area_data = []

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

                concentration_data.append({'Name': name, 'Concentration': concentration, 'Parent File Path': pdata_dir})
                area_data.append({'Name': name, 'Area': peak_area, 'Parent File Path': pdata_dir})

            pdata_results[pdata_dir] = {
                'Concentration': pd.DataFrame(concentration_data),
                'Area': pd.DataFrame(area_data),
                'Peak Limits': peak_data
            }

        for pdata_df in pdata_results.values():
            results.extend(pdata_df['Concentration'].to_dict(orient='records'))

    if not results:
        messagebox.showerror("No Data", "No pdata/1 directories were processed.")
        return

    concentrations_df = pd.DataFrame(results)
    concentrations_df = concentrations_df.pivot(index='Parent File Path', columns='Name', values='Concentration')

    peak_areas_df = pd.concat([pdata_df['Area'] for pdata_df in pdata_results.values()], ignore_index=True)
    peak_areas_df = peak_areas_df.pivot(index='Parent File Path', columns='Name', values='Area')

    peak_limits_df = pd.concat([pdata_df['Peak Limits'] for pdata_df in pdata_results.values()], ignore_index=True)

    with pd.ExcelWriter('nmr_integration_results.xlsx') as writer:
        concentrations_df.to_excel(writer, sheet_name='Concentrations')
        peak_areas_df.to_excel(writer, sheet_name='Peak Areas')
        peak_limits_df.to_excel(writer, sheet_name='Peak Limits', index=False)

    messagebox.showinfo("Analysis Complete", "Check the results in 'nmr_integration_results.xlsx'")

    # Clean up temporary directories
    if 'temp' in selected_pdata_dirs:
        shutil.rmtree('temp')

def browse_directory():
    global selected_pdata_dirs
    file_or_dir = messagebox.askquestion("Zipped File or Directory", "Are you selecting a zipped file?", icon='question')
    if file_or_dir == 'yes': # If the user chooses 'yes', we'll open the file selection dialog.
        root_dir = filedialog.askopenfilename(title="Select the zipped file containing processed Bruker NMR data")
        if root_dir.endswith('.zip'):
            with zipfile.ZipFile(root_dir, 'r') as zip_ref:
                zip_ref.extractall('temp')
                root_dir = 'temp'
    else: # If the user chooses 'no', we'll open the directory selection dialog.
        root_dir = filedialog.askdirectory(title="Select the root directory containing processed Bruker NMR data")
    selected_pdata_dirs = [root_dir]
    #selected_dirs_label.config(text=f"Selected Directory: {root_dir}") #Minor issue saying temp when zipped files are selected

selected_pdata_dirs = []
tsp_concentration = 0

root = tk.Tk()
root.title("Bruker NMR Plasma Metabolism Analysis")

selected_dirs_label = tk.Label(root, text="Selected Directory: None")
selected_dirs_label.pack(pady=10)

dirs_button = tk.Button(root, text="Browse Data Directory or Zipped File", command=browse_directory)
dirs_button.pack()

concentration_label = tk.Label(root, text="Enter Reference Concentration in Micromolar:")
concentration_label.pack(pady=5)

concentration_entry = tk.Entry(root)
concentration_entry.pack()

submit_button = tk.Button(root, text="Submit", command=process_selected_dirs)
submit_button.pack(pady=10)

#progress_label = tk.Label(root, text="Select a file or directory. Enter reference concentration. Submit.")
#progress_label.pack()

root.mainloop()