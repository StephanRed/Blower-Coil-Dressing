import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os
import re

# Load initial Excel structure (empty until make is selected)
file_path = 'Blower coil dressing LABELS HI.xlsx'
parts_file = 'Parts.xlsx'
eazy_file = 'Eazy_Data.xlsx'
csv_file = 'CSV.csv'
make_sheets = {
    "CSB": "CSB Data",
    "HI": "HI Data",
    "HSe": "HSe Data"
}

#resest quantity values in Parts.xls
def reset_quantities():
    df = pd.read_excel(parts_file)
    df["Quantity"] = 0
    df.to_excel(parts_file, index=False)

reset_quantities()

# Normalize pipe sizes for consistent comparison
def normalize_size(size):
    return str(size).strip().lstrip("'").replace(' \"', '').replace('"', '')

# Data holders
evap_specs = {}
valve_specs = {}
valve_type = ""
sizes = ['1/2', '5/8', '3/4', '7/8', '1-1/8', '1-3/8', '1-5/8', '2-1/8', '2-5/8']
orifice_sizes = ['1', '2', '3', '4', '5', '6', '7', '8', '9']

collected_parts = {}  # Global dictionary to store parts and quantities

# GUI setup
root = tk.Tk()
root.title("Blower Coil Dressing")
root.geometry("510x365")

make_var = tk.StringVar()
model_var = tk.StringVar()
suction_var = tk.StringVar()
liquid_var = tk.StringVar()
valve_var = tk.StringVar()
orifice_var = tk.StringVar()
parts_collection = {}  # Stores parts and quantities
evaporator_count = tk.IntVar(value=0)  # Tracks number of evaporators added

form_frame = ttk.Frame(root)
form_frame.pack()

# Row 0
ttk.Label(form_frame, text="Select Make:").grid(row=0, column=0, sticky="e", padx=5, pady=2)
make_menu = ttk.Combobox(form_frame, textvariable=make_var, values=list(make_sheets.keys()))
make_menu.grid(row=0, column=1)

# Row 1
ttk.Label(form_frame, text="Select Model:").grid(row=1, column=0, sticky="e", padx=5, pady=2)
model_menu = ttk.Combobox(form_frame, textvariable=model_var)
model_menu.grid(row=1, column=1, padx=5, pady=2)

# Row 2
ttk.Label(form_frame, text="Main Suction Size:").grid(row=0, column=2, sticky="e", padx=5, pady=2)
suction_menu = ttk.Combobox(form_frame, textvariable=suction_var, values=sizes)
suction_menu.grid(row=0, column=3, padx=5, pady=2)

# Row 3
ttk.Label(form_frame, text="Main Liquid Size:").grid(row=1, column=2, sticky="e", padx=5, pady=2)
liquid_menu = ttk.Combobox(form_frame, textvariable=liquid_var, values=sizes)
liquid_menu.grid(row=1, column=3, padx=5, pady=2)

# Row 4
ttk.Label(form_frame, text="Valve Type:").grid(row=4, column=0, sticky="e", padx=5, pady=2)
valve_menu = ttk.Combobox(form_frame, textvariable=valve_var)
valve_menu.grid(row=4, column=1, padx=5, pady=2)

ttk.Label(form_frame, text="Orifice Size:").grid(row=4, column=2, sticky="e", padx=5, pady=2)
orifice_menu = ttk.Combobox(form_frame, textvariable=orifice_var, values=orifice_sizes)
orifice_menu.grid(row=4, column=3, padx=5, pady=2)

evaporator_label = ttk.Label(form_frame, text="Evaporators added: 0")
evaporator_label.grid(row=6, column=0, columnspan=2, pady=5)



# Load valve data once
valve_df = pd.read_excel(file_path, sheet_name="Valve Data", header=0)
valve_df.columns = valve_df.columns.str.strip()
for _, row in valve_df.iterrows():
    valve = str(row.get('Model', '')).strip()
    inlet = normalize_size(row.get('Inlet Size', ''))
    outlet = normalize_size(row.get('Outlet Size', ''))
    if valve:
        valve_specs[valve] = {'inlet': inlet, 'outlet': outlet}
valve_menu['values'] = list(valve_specs.keys())

# === Reload data when Make is changed ===
def on_make_change(event):
    global evap_specs
    selected_make = make_var.get()
    sheet = make_sheets.get(selected_make)
    if not sheet:
        return

    try:
        df = pd.read_excel(file_path, sheet_name=sheet, header=2)
        df.columns = df.columns.str.strip()

        evap_specs = {}
        for _, row in df.iterrows():
            model = str(row['Model']).strip()
            suction = normalize_size(str(row.get('Suction size', '')).strip())
            liquid = normalize_size(str(row.get('Liquid size', '')).strip())
            fans = row.get('Number of Fans', None)
            if model:
                evap_specs[model] = {'suction': suction, 'liquid': liquid, 'fans': fans}

        # Update model dropdown
        model_menu['values'] = list(evap_specs.keys())
        model_var.set("")  # Clear current selection
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load sheet '{sheet}': {e}")

make_menu.bind("<<ComboboxSelected>>", on_make_change)

def compare_sizes(main, spec):
    if not main or not spec:
        return "INVALID SIZE"
    return "NONE" if main == spec else f"{main} to {spec}"

    # Update model dropdown
    model_menu['values'] = list(evap_specs.keys())
    model_var.set("")  # Clear current selection

def generate_parts():
    output_text.delete("1.0", tk.END)

    model = model_var.get()
    main_suction = normalize_size(suction_var.get())
    main_liquid = normalize_size(liquid_var.get())
    valve_type = valve_var.get()
    orifice_size = orifice_var.get()

    if not all([model, main_suction, main_liquid, valve_type]):
        messagebox.showerror("Error", "Please select all inputs.")
        return

    spec = evap_specs.get(model)
    if not spec:
        messagebox.showerror("Error", "Model not found in data.")
        return

    evap_suction = normalize_size(spec.get('suction', ''))
    evap_liquid = normalize_size(spec.get('liquid', ''))
    fans = spec.get('fans', 0)

    suction_reducer = compare_sizes(main_suction, evap_suction)

    output = [
        f"Suction reducer: {suction_reducer}",
        f"Suction Elbow size: {evap_suction}",
        f"Suction P-Trap size: {main_suction}",
    ]

    valve_data = valve_specs.get(valve_type)
    if valve_data:
        inlet = valve_data.get('inlet', '')
        outlet = valve_data.get('outlet', '')

        reducer_1 = compare_sizes(inlet, main_liquid)
        reducer_2 = compare_sizes(outlet, evap_liquid)

        output.append(f"Liquid reducer 1: {reducer_1}")
        output.append(f"Liquid reducer 2: {reducer_2}")
    else:
        output.append("Valve data not found.")

    try:
        freezer = model[-2:].upper() == "4L" or model[-1:].upper() == "E" or model[6].upper() == "W"
        armaflex = f"Armaflex {'1' if freezer else '3/4'}\" thick: {main_suction}"
    except:
        armaflex = f"Armaflex 3/4\" thick: {main_suction}"

    if valve_type[:3].upper() != "EEV":
        output.append(f"Solenoid: {main_liquid} Castel")
    output.append(f"Sweat on drier: {main_liquid}")
    output.append(armaflex)
    output.append(f"Expansion Valve: {valve_type}")
    output.append(f"Orifice {valve_type}:{orifice_size}")

    if fans:
        try:
            fans_int = (int(fans) + 1) * 2
            output.append(f"M8 Dome Nuts: {fans_int}")
            output.append(f"80x80x3mm Plate: {fans_int}")
        except:
            output.append(f"Mountings: (Invalid fan count: {fans})")

    output_text.insert(tk.END, "\n".join(output))

def categorize_part(desc):
    desc = desc.lower()
    if "reducer" in desc:
        return "Reducers"
    elif "elbow" in desc:
        return "Elbows"
    elif "p-trap" in desc:
        return "P-Traps"
    elif "solenoid" in desc:
        return "Solenoid"
    elif "drier" in desc:
        return "Driers"
    elif "armaflex 1\"" in desc:
        return "Armaflex 1\""
    elif "armaflex 3/4\"" in desc:
        return "Armaflex 3/4\""
    elif "orifice tes2" in desc:
        return "Orifice TES2"
    elif "orifice tes 5" in desc:
        return "Orifice TES 5"
    elif "orifice eev2028/l" in desc:
        return "Orifice EEV2028/L"
    elif "orifice eev2028/s" in desc:
        return "Orifice EEV2028/S"
    elif "nuts" in desc:
        return "Nuts"
    elif "plate" in desc:
        return "Plates"
    elif "expansion valve" in desc:
        return "Expansion Valve"
    else:
        return "Accessories"  # fallback

def save_part_to_excel(desc, part):
    category = categorize_part(desc)
    df = pd.read_excel(parts_file)

    match = (df['Category'].str.strip().str.lower() == category.lower()) & \
            (df['Part Description'].str.strip().str.lower() == part.lower())

    if category in ["Nuts", "Plates"]:
        match = (df['Category'] == category)
        if not df[match].empty:
            idx = df[match].index[0]
            df.at[idx, 'Quantity'] += int(part)
    else:

        if not df[match].empty:
            idx = df[match].index[0]
            df.at[idx, 'Quantity'] += 1
        else:
            if part != "NONE":
                new_row  =pd.DataFrame([{
                    "Category": category,
                    "Part Description": part,
                    "Quantity": 1
                }])
                df = pd.concat([df, new_row], ignore_index=True)

    df.to_excel(parts_file, index=False)


def add_parts():
    global collected_parts
    text_content = output_text.get("1.0", tk.END).strip()
    if not text_content:
        return
    lines = text_content.split("\n")
    for line in lines:
        if ":" in line:
            desc, part = line.split(":", 1)
            save_part_to_excel(desc.strip(), part.strip())
        # Clear output and update evaporator count label only if not empty

    output_text.delete("1.0", tk.END)
    count = evaporator_count.get() + 1
    evaporator_count.set(count)
    evaporator_label.config(text=f"Evaporators added: {count}")


def show_totals():
    df = pd.read_excel(parts_file)

    # Clear the output field (adjust to your widget type)
    output_text.delete('1.0', 'end')  # if using Listbox
    # output_text.delete('1.0', 'end')  # if using Text widget

    # Get unique categories in order
    categories = df['Category'].dropna().unique()

    for category in categories:
        # Filter parts with this category and quantity > 0
        category_parts = df[
            (df['Category'].str.strip().str.lower() == category.strip().lower()) &
            (df['Quantity'] > 0)
            ]

        if not category_parts.empty:
            # Show category header
            output_text.insert('end', f"{category.strip()}\n")
            # Show each part
            for _, row in category_parts.iterrows():
                part_desc = row['Part Description']
                quantity = int(row['Quantity'])
                output_text.insert('end', f"{part_desc.strip()} : {quantity}\n")
            output_text.insert('end', "\n")  # Add spacing between categories

#Export parts to csv
def export_to_csv():
    global valve_type
    solenoid_bundle = [{
        "StockItem": "142861",
        "Quantity": 0,
        "MiscDescription": "Solenoid Coil only 230V"
    },{
        "StockItem": "142863",
        "Quantity": 0,
        "MiscDescription": "Solenoid Plug only"
    }
    ]
    EEV_bundle = [{
        "StockItem": "142205",
        "Quantity": 0,
        "MiscDescription": "EEV Coil 230V"
    }, {
        "StockItem": "142201",
        "Quantity": 0,
        "MiscDescription": "EEV Coil Plug + 6.5m Cable"
    }
    ]
    try:
        # Load Excel files
        parts_df = pd.read_excel(parts_file )
        parts_df = parts_df[parts_df["Quantity"] > 0]

        eazy_df = pd.read_excel(eazy_file)

        with open("CSV.csv", "w", encoding="utf-8") as f:
            f.write("Group,Description,Comments\n")  # First blank row
            f.write("\n")  # Second blank row
        # Clean column names if needed
        parts_df.columns = parts_df.columns.str.strip()
        eazy_df.columns = eazy_df.columns.str.strip()

        def extract_sizes(text):
            pattern = r'\b\d+-\d+/\d+|\b\d+/\d+\b'
            matches = re.findall(pattern, text)
            return matches[:2]  # Return only first two matches

        # Rename to make comparing easier
        parts_df.rename(columns={"Part Description": "Part_Desc"}, inplace=True)
        eazy_df.rename(columns={"Description": "Category", "Part": "Eazy_Part"}, inplace=True)

        results = []
        solenoid_count = 0
        solenoid_bundle_working = [
            {"StockItem": item["StockItem"], "Quantity": 0, "MiscDescription": item["MiscDescription"]}
            for item in solenoid_bundle
        ]
        EEV_count = 0
        EEV_bundle_working = [
            {"StockItem": item["StockItem"], "Quantity": 0, "MiscDescription": item["MiscDescription"]}
            for item in EEV_bundle
        ]
        # Iterate through each row in parts_df to find matching category
        for _, part_row in parts_df.iterrows():
            category = part_row['Category']
            part_desc = str(part_row['Part_Desc']).strip()
            quantity = part_row['Quantity']

            # Filter eazy_df by the same category
            eazy_subset = eazy_df[eazy_df['Category'] == category]

            match = None

            # Apply category-specific logic
            if category == "Reducers":
                # Extract sizes from part description
                sizes = part_desc.replace(" ", "").split("to")
                sizes = [s.strip() for s in sizes]
                sizes_set = set(sizes)

                for _, eazy_row in eazy_subset.iterrows():
                    eazy_part = str(eazy_row['Eazy_Part'])
                    eazy_sizes = extract_sizes(eazy_part)
                    if set(eazy_sizes) == sizes_set:
                        match = {
                            "StockItem": eazy_row['ID'],
                            "Quantity": quantity,
                            "MiscDescription": category + " " + eazy_part
                        }
                        break
            elif category in ["Elbows", "P-Traps"]:
                for _, eazy_row in eazy_subset.iterrows():
                    eazy_part = str(eazy_row['Eazy_Part'])
                    # Extract sizes safely from Eazy_Part
                    sizes = extract_sizes(eazy_part)
                    # Match part_desc exactly
                    if any(size.strip() == part_desc.strip() for size in sizes):
                        match = {
                            "StockItem": eazy_row['ID'],
                            "Quantity": quantity,
                            "MiscDescription": category + " " + eazy_part
                        }
                        break
            elif category in "Solenoid":
                eazy_subset = eazy_df[eazy_df['Category'].str.contains(category, case=False, na=False)]
                for _, eazy_row in eazy_subset.iterrows():
                    eazy_part = str(eazy_row["Eazy_Part"])  # 🔐 convert to string safely
                    size = str(part_desc.split(" ")[0])

                    if (size in eazy_part) and (category in eazy_part):
                        match = {
                            "StockItem": eazy_row['ID'],
                            "Quantity": quantity,
                            "MiscDescription": eazy_part
                        }
                        break
            elif category in "Driers":
                eazy_subset = eazy_df[eazy_df['Category'].str.contains("Drier", case=False, na=False)]
                for _, eazy_row in eazy_subset.iterrows():
                    size = str(part_desc)
                    if size in eazy_row["Category"]:
                        match = {
                            "StockItem": eazy_row['ID'],
                            "Quantity": quantity,
                            "MiscDescription": eazy_row["Category"]
                        }
                        break
            elif "Armaflex" in category:
                eazy_subset = eazy_df[eazy_df['Category'].str.contains("Armaflex", case=False, na=False)]
                for _, eazy_row in eazy_subset.iterrows():
                    eazy_part = str(eazy_row["Eazy_Part"])
                    eazy_size = extract_sizes(eazy_part)
                    if (category == eazy_row["Category"]) and (part_desc == eazy_size[0]):
                        match = {
                            "StockItem": eazy_row['ID'],
                            "Quantity": quantity,
                            "MiscDescription": f"Armaflex {eazy_part}"
                        }
                        break
            elif ("Expansion Valve" in category) and ("TES" in part_desc):
                valve_type = part_desc
                eazy_subset = eazy_df[eazy_df['Category'].str.contains("Expansion Valve", case=False, na=False)]
                for _, eazy_row in eazy_subset.iterrows():
                    eazy_part = str(eazy_row['Eazy_Part'])
                    if valve_type in eazy_part:
                        match = {
                            "StockItem": eazy_row['ID'],
                            "Quantity": quantity,
                            "MiscDescription": category + " " + eazy_part
                        }
                        break
            elif "Orifice" in category:
                if "EEV2028" in category:
                    EEV_s = category[-1:].upper() == "S"
                    eazy_subset = eazy_df[eazy_df['Eazy_Part'].str.contains(f"EEV2028/{'S' if EEV_s else 'L'}", case=False, na=False)]
                    for _, eazy_row in eazy_subset.iterrows():
                        orifice_size = part_desc
                        EEV_valve = str(eazy_row["Eazy_Part"])
                        if EEV_valve.endswith(orifice_size):
                            match = {
                                "StockItem": eazy_row['ID'],
                                "Quantity": quantity,
                                "MiscDescription": "Electronic Expansion Valve " + EEV_valve
                            }
                            break
                else:
                    eazy_subset = eazy_df[eazy_df['Category'].str.contains("Orifice", case=False, na=False)]
                    valve_type = category.replace("Orifice ", "")
                    for _, eazy_row in eazy_subset.iterrows():
                        eazy_part = str(eazy_row['Eazy_Part']).replace(" -", "")
                        check =  f"{valve_type} NO.{part_desc}"
                        if check in eazy_part:
                            match = {
                                "StockItem": eazy_row['ID'],
                                "Quantity": quantity,
                                "MiscDescription": "Orifice " + eazy_part
                            }
                            break
            elif category in "Plates":
                for _, eazy_row in eazy_df.iterrows():
                    if category in eazy_row["Category"]:
                        match = {
                            "StockItem": eazy_row['ID'],
                            "Quantity": quantity,
                            "MiscDescription": eazy_row["Category"]
                        }
                        break
            elif category in "Nuts":
                for _, eazy_row in eazy_df.iterrows():
                    if category in eazy_row["Category"]:
                        match = {
                            "StockItem": eazy_row['ID'],
                            "Quantity": quantity,
                            "MiscDescription": eazy_row["Category"]
                        }
                        break

            # Add more `elif category == ...:` blocks here for Accessories, Fittings, etc.

            # If a match was found, add to results

            if match:
                results.append(match)
                if "Solenoid" in match["MiscDescription"]:
                    solenoid_count += match["Quantity"]
                if "EEV" in match["MiscDescription"]:
                    EEV_count += match["Quantity"]
        # Convert to DataFrame and export to CSV
        # After the loop is finished
        extras = [{
            "StockItem": "142807",
            "Quantity": evaporator_count.get(),
            "MiscDescription": "Gland Compression NO.0 PVC"
        },{
            "StockItem": "142527",
            "Quantity": evaporator_count.get(),
            "MiscDescription": "Solder Flare Adapter"
        },{
            "StockItem": "140143",
            "Quantity": 1,
            "MiscDescription": "Cable Ties (100/pack) T30R Black 150 L x 3.5 W"
        },{
            "StockItem": "143985",
            "Quantity":1,
            "MiscDescription": "Contact Adhesive (12/box bulk) 250ml"
        }]

        if solenoid_count > 0:
            for item in solenoid_bundle_working:
                item["Quantity"] = solenoid_count
            results.extend(solenoid_bundle_working)
        if EEV_count > 0:
            for item in EEV_bundle_working:
                item["Quantity"] = EEV_count
            results.extend(EEV_bundle_working)

        results.extend(extras)

        results_df = pd.DataFrame(results)
        results_df.to_csv(csv_file, mode="a", sep=",", index=False, header=True)  # Now header appears on row 3
        messagebox.showinfo("Complete", "Exported successfully.")

    except Exception as e:
        messagebox.showerror("Export Failed", str(e))

#Buttons
submit_btn = ttk.Button(form_frame, text="Generate Parts", command=generate_parts)
submit_btn.grid(row=5, column=1, padx=5, pady=2, sticky="e")

add_btn = ttk.Button(form_frame, text="Add", command=add_parts)
add_btn.grid(row=5, column=2, padx=5, pady=2, sticky="w")

show_btn = ttk.Button(form_frame, text="Show Total", command=show_totals)
show_btn.grid(row=5, column=3, padx=5, pady=2, sticky="w")

export_button = ttk.Button(form_frame, text="Export", command=export_to_csv)
export_button.grid(row=6, column=2, padx=5, pady=2, sticky="w")

# Output Box
output_text = tk.Text(root, height=15, width=60)
output_text.pack(pady=5)

root.mainloop()
