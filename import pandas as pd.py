import pandas as pd
import json
from tkinter import filedialog

def clean_hex(hex_str):
    if pd.isna(hex_str):
        return None

    hex_str = str(hex_str).replace(" ", "").upper()

    if len(hex_str) != 4:
        return None

    # convert to little endian (like your save uses)
    return hex_str[2:] + hex_str[:2]


def excel_to_json(input_file, output_file):
    df = pd.read_excel(input_file)

    # Normalize column names
    df.columns = df.columns.str.strip().str.upper()

    result = []

    for _, row in df.iterrows():
        effect = str(row["SPECIAL EFFECT"]).strip()
        hex_id = clean_hex(row["HEX ID"])
        cap = str(row["EFFECT MAX"]).strip()
        typ = str(row["TYPE"]).strip()

        if not hex_id:
            continue

        result.append({
            "Effect": effect,
            "Effect Max": cap,
            "type": typ,
            "id": hex_id
        })

    with open(output_file, "w") as f:
        json.dump(result, f, indent=4)

    print(f"Saved to {output_file}")


# === FILE PICKER ===
path = filedialog.askopenfilename(title="Select Excel file")

excel_to_json(path, "effects.json")