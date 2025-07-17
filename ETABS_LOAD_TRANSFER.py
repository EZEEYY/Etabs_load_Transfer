import pandas as pd
import numpy as np
import comtypes.client
import matplotlib.pyplot as plt
from comtypes.client import CreateObject
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import comtypes

    
def select_load_cases(all_cases):
    selected = []

    def on_ok():
        for i, var in enumerate(variables):
            if var.get():
                selected.append(all_cases[i])
        root.destroy()

    root = tk.Tk()
    root.title("Select Load Cases")

    tk.Label(root, text="Choose Load Cases to Apply:").pack(padx=10, pady=10)

    frame = ttk.Frame(root)
    frame.pack(fill="both", expand=True, padx=10, pady=(0,10))

    canvas = tk.Canvas(frame)
    scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    variables = []
    for case in all_cases:
        var = tk.BooleanVar()
        chk = ttk.Checkbutton(scrollable_frame, text=case, variable=var)
        chk.pack(anchor="w")
        variables.append(var)

    tk.Button(root, text="OK", command=on_ok).pack(pady=5)

    root.geometry("300x400") 
    root.mainloop()

    return selected

def select_files_window():
    src_path = ""
    dest_path = ""

    def browse_source():
        nonlocal src_path
        file = filedialog.askopenfilename(
            title="Select SOURCE ETABS File",
            filetypes=[("ETABS Files", "*.EDB"), ("All Files", "*.*")]
        )
        if file:
            src_path = file
            src_entry_var.set(src_path)

    def browse_destination():
        nonlocal dest_path
        file = filedialog.asksaveasfilename(
            title="Select Destination File",
            defaultextension=".EDB",
            filetypes=[("ETABS Files", "*.EDB"), ("All Files", "*.*")]
        )
        if file:
            dest_path = file
            dest_entry_var.set(dest_path)

    def on_ok():
        nonlocal src_path, dest_path
        if not src_path or not dest_path:
            messagebox.showerror("Error", "Please select both source and destination files.")
            return
        root.destroy()

    root = tk.Tk()
    root.title("Select source and Destination ETABS Files")

    ttk.Label(root, text="SOURCE ETABS File:").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 2))
    src_entry_var = tk.StringVar()
    src_entry = ttk.Entry(root, textvariable=src_entry_var, width=60)
    src_entry.grid(row=1, column=0, padx=10, sticky="we")
    ttk.Button(root, text="Browse...", command=browse_source).grid(row=1, column=1, padx=10, pady=2)

    ttk.Label(root, text="Destination ETABS File:").grid(row=2, column=0, sticky="w", padx=10, pady=(10, 2))
    dest_entry_var = tk.StringVar()
    dest_entry = ttk.Entry(root, textvariable=dest_entry_var, width=60)
    dest_entry.grid(row=3, column=0, padx=10, sticky="we")
    ttk.Button(root, text="Browse...", command=browse_destination).grid(row=3, column=1, padx=10, pady=2)

    ttk.Button(root, text="OK", command=on_ok).grid(row=4, column=0, columnspan=2, pady=15)

    root.columnconfigure(0, weight=1)
    root.geometry("700x170")
    root.mainloop()

    return src_path, dest_path

if __name__ == "__main__":
    source_file, destination_file = select_files_window()
    print("Source file:", source_file)
    print("Destination file:", destination_file)



# Create API helper object
helper = comtypes.client.CreateObject('ETABSv1.Helper')
helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)
EtabsObject = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)
EtabsObject = helper.CreateObjectProgID("CSI.ETABS.API.ETABSObject")

# Start ETABS application
EtabsObject.ApplicationStart()

# Create SapModel object
SapModel = EtabsObject.SapModel


# Open the model
SapModel.File.OpenFile(source_file)
# Run analysis
ret = SapModel.Analyze.RunAnalysis()

# Retrieve the list of all response combinations
nos, combinations, nill= SapModel.RespCombo.GetNameList()

# Update units to kN-m
SapModel.SetPresentUnits(6)



# Get all joint (point) names
ret, joint_names, zero = SapModel.PointObj.GetNameList()

supported_joints = []


for joint in joint_names:
    restraints, zero = SapModel.PointObj.GetRestraint(joint)
    print(restraints)
    if restraints and any(r == 1 for r in restraints):
        x, y, z, a = SapModel.PointObj.GetCoordCartesian(joint)
        print(x)
        print(y)

        supported_joints.append({
            "Name": joint,
            "Restraints": restraints,
            "X": x,
            "Y": y,
            "Z": z
        })

for joint in supported_joints:
    print(f"{joint['Name']} → Restraints: {joint['Restraints']}, Coordinates: X={joint['X']}, Y={joint['Y']}, Z={joint['Z']}")

ret, load_cases, lame = SapModel.LoadCases.GetNameList()
selected_cases = select_load_cases(load_cases)
print(selected_cases)

if not selected_cases:
    raise Exception("❌ No load cases selected. Aborting.")


# Get all joints
ret, joint_names, zero = SapModel.PointObj.GetNameList()



results_data = []
def max_abs(val_tuple):
    return max(val_tuple, key=abs) 

for load_case in selected_cases:
    SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
    SapModel.Results.Setup.SetCaseSelectedForOutput(load_case)

    print(f"\n=== Load Case: {load_case} ===")
    print("\nMax Reaction Forces at Support Joints:")

    for joint in [j['Name'] for j in supported_joints]:
        try:
            ret, NumberResults, Obj, Elm, LoadCase, StepType, F1, F2, F3, M1, M2, M3, zero = SapModel.Results.JointReact(joint, 0)
            print(ret, NumberResults, Obj, Elm, LoadCase, StepType, F1, F2, F3, M1, M2, M3, zero)
            if NumberResults == 0:
                continue

            fx = max_abs(F1)
            fy = max_abs(F2)
            fz = max_abs(F3)
            mx = max_abs(M1)
            my = max_abs(M2)
            mz = max_abs(M3)

            x, y, z, a = SapModel.PointObj.GetCoordCartesian(joint)

            print(f"{joint} → LoadCase: {load_case} | FX={fx:.2f}, FY={fy:.2f}, FZ={fz:.2f}, MX={mx:.2f}, MY={my:.2f}, MZ={mz:.2f}")

            results_data.append({
                "Joint": joint,
                "X": x,
                "Y": y,
                "Z": z,
                "LoadCase": load_case,
                "FX": fx,
                "FY": fy,
                "FZ": fz,
                "MX": mx,
                "MY": my,
                "MZ": mz
            })

        except Exception as e:
            print(f"Error processing joint {joint}: {e}")

print(results_data)
df = pd.DataFrame(results_data)


# GUI for saving Excel file
def save_excel_file(results_data):
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="Save reaction results as Excel"
    )
    if file_path:
        df = pd.DataFrame(results_data)
        df.to_excel(file_path, index=False)
        print(f"Results saved to: {file_path}")
    else:
        print("Excel save cancelled.")

save_excel_file(results_data)








# === Open existing model ===
SapModel.File.OpenFile(destination_file)


# === Set units to MKS (N, m, C) ===
SapModel.SetPresentUnits(6)  # 6 = N, m, C




#print(f"Joint loads successfully applied to: {destination_file}")


# Get all existing joints in model
ret, existing_joints, zero = SapModel.PointObj.GetNameList()

# Map of coordinates of existing joints (to avoid duplicates)
existing_coords = {}
for joint in existing_joints:
    x, y, z,a = SapModel.PointObj.GetCoordCartesian(joint)
    existing_coords[(round(x, 6), round(y, 6), round(z, 6))] = joint 

for data in results_data:
    x, y, z = data["X"], data["Y"], data["Z"]
    coord_key = (round(x, 6), round(y, 6), round(z, 6))

    if coord_key in existing_coords:
        ret = existing_coords[coord_key]
    else:
        ret, _ = SapModel.PointObj.AddCartesian(x, y, z, "")
        existing_coords[coord_key] = ret  # Add to map to avoid duplicates if needed

    load_case = data["LoadCase"]

    def safe_first(val):
        if isinstance(val, (tuple, list, np.ndarray)):
            return val[0] if len(val) > 0 else 0.0
        return val

    fx = safe_first(data["FX"])
    fy = safe_first(data["FY"])
    fz = safe_first(data["FZ"])
    mx = safe_first(data["MX"])
    my = safe_first(data["MY"])
    mz = safe_first(data["MZ"])


    fx = -fx
    fy = -fy
    fz = -fz
    mx = -mx
    my = -my
    mz = -mz

    SapModel.PointObj.SetLoadForce(ret, load_case, [fx, fy, fz, mx, my, mz])


# Save model
SapModel.File.Save()
print("All loads applied and model saved.")



print("Model saved successfully.")
EtabsObject.ApplicationExit(False)


