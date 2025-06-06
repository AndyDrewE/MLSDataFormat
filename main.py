import csv
import os
from tkinter import Tk, Label, filedialog
from tkinterdnd2 import DND_FILES, TkinterDnD

def process_csv(file_path):
    desktop = os.path.join(os.path.expanduser("~"), "Desktop", "Exports")
    os.makedirs(desktop, exist_ok=True)
    output_file = os.path.join(desktop, "Export_Formatted.csv")

    with open(file_path, newline='', encoding='utf-8') as infile:
        reader = list(csv.reader(infile))

    headers = reader[0]
    rows = reader[1:]

    def ucase(val):
        return val.strip().upper() if val else ""

    for row in rows:
        # Ensure row is long enough
        row += [""] * (30 - len(row))

        # Replace OKLAHOMA
        if ucase(row[2]) == "OKLAHOMA":
            row[2] = "OK"

        # MLSNumber
        if row[5].isdigit():
            row[5] = "OKCMLS#" + row[5]

        # Concessions
        if not row[7] and row[8]:
            row[8] = ""
        elif row[7] and not row[8]:
            row[8] = "0"

        # One-story default
        if not row[13]:
            row[13] = "1"

        # Design
        if ucase(row[14]) == "TRADITIONAL":
            row[14] = "Trad"

        # Heating
        heat = ucase(row[19])
        if heat in ["CENTRAL GAS", "CENTRAL ELECTRIC"]:
            row[19] = "FWA"
        elif heat == "FLOOR FURNACE":
            row[19] = "RAD"
        else:
            row[19] = "   "

        # Cooling
        cool = ucase(row[20])
        if cool in ["CENTRAL ELEC", "CENTRAL GAS"]:
            row[20] = "CAC"
        else:
            row[20] = "   "

        # Garage formatting
        cap = row[21]
        desc = ucase(row[12])
        if cap == "0":
            row[21] = "None"
        elif desc == "ATTACHED":
            row[21] = f"{cap}ga{cap}dw"
        elif desc == "DETACHED":
            row[21] = f"{cap}gd{cap}dw"
        elif desc == "CARPORT":
            row[21] = f"{cap}cp"

        # Sales status
        status = ucase(row[22])
        if status == "ACTIVE":
            row[25] = "Listing"
        elif status == "PENDING":
            row[25] = "Listing"
            row[22] = "Contract"
        elif status == "SOLD":
            row[22] = "Settled sale"

        # Property type formatting
        stories = row[13]
        design = row[14]
        prop = ucase(row[24])
        if prop in ["CONDOMINIUM", "HALF DUPLEX", "TOWNHOUSE"]:
            row[24] = f"AT{stories};{design}"
        elif prop == "PATIO HOME":
            row[24] = "AT"
        else:
            row[24] = f"DT{stories};{design}"

        # Fee simple + source
        row[26] = "Fee Simple"
        row[27] = "Ext Insp PR B/P "

    with open(output_file, "w", newline='', encoding='utf-8') as outfile:
        writer = csv.writer(outfile)
        writer.writerow(headers)
        writer.writerows(rows)

    print("Processing complete.")
    return output_file

# GUI with drag-and-drop
def create_gui():
    def drop(event):
        file_path = event.data.strip("{}")
        if file_path.lower().endswith(".csv"):
            try:
                result = process_csv(file_path)
                label.config(text=f"✅ Done!\nSaved to:\n{result}\n\nDrop another file to process.")
            except Exception as e:
                label.config(text=f"❌ Error:\n{str(e)}")
        else:
            label.config(text="⚠️ Please drop a valid .csv file.")

    app = TkinterDnD.Tk()
    app.title("CSV Formatter")
    app.geometry("400x200")

    label = Label(app, text="Drop your CSV file here", width=40, height=10, bg="lightgray")
    label.pack(padx=10, pady=10)

    label.drop_target_register(DND_FILES)
    label.dnd_bind("<<Drop>>", drop)

    app.mainloop()

if __name__ == "__main__":
    create_gui()
