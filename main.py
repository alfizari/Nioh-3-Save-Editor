import subprocess
from tkinter import Tk, filedialog, ttk, messagebox
import os
from pathlib import Path
import json
import tkinter as tk
from checksum import patch_checksum
from openpyxl import Workbook
import shutil

AMRITA_OFFSET  = 0x3ADE49
GOLD_OFFSET    = 0x3ADE59
CONSTITUTION   = 0x3ADF63
HEART          = 0x3ADF67
STAMINA        = 0x3ADF6B
STRENGTH       = 0x3ADF6F
SKILL          = 0x3ADF73
INTELLECT      = 0x3ADF77
MAGIC          = 0x3ADF7B

ITEM_START    = 0x240355
ITEM_SIZE     = 0xF0
ITEM_SLOTS    = 0x9C4

USABLE_START  = 0x2D2B21
USABLE_SIZE   = 0xE8
USABLE_SLOTS  = 0x5DC

STORAGE_START = 0x327A8D
STORAGE_SIZE  = 0xE8
STORAGE_SLOTS = 0x189


data                   = None
MODE                   = None
IMPORT_MODE            = None
import_data            = None
decrypted_path_import  = None
decrypted_path         = None
items                  = []
usables                = []
storage                = []

base_dir = Path(__file__).parent


def load_json(file_name):
    with open(base_dir / file_name, "r") as f:
        return json.load(f)

items_json           = load_json("items_little_endian.json")
effects_json         = load_json("effects_big_endian.json")


effect_dropdown_list = ["{} - {}".format(e["id"], e["Effect"]) for e in effects_json]


_effects_by_type = {}
for _e in effects_json:
    _t = _e.get("type", "")
    if _t not in _effects_by_type:
        _effects_by_type[_t] = []
    _effects_by_type[_t].append("{} - {}".format(_e["id"], _e["Effect"]))


_effect_types_available = frozenset(_effects_by_type.keys())


_auto_map = {
    itype: itype
    for itype in set(v.get("type", "") for v in items_json.values())
    if itype in _effect_types_available
}

_MANUAL_OVERRIDES = {
    "Sword":        "Weapon",
    "Dual Swords":  "Weapon",
    "Axe":          "Weapon",
    "Spear":        "Weapon",
    "Odachi":       "Weapon",
    "Kusarigama":   "Weapon",
    "Tonfa":        "Weapon",
    "Hatchet":      "Weapon",
    "Rifle":        "Weapon",
    "Cannon":       "Weapon",
    "Bow":          "Weapon",
    "Helmet":       "Armor",
    "Chest":        "Armor",
    "Gloves":       "Armor",
    "Leggings":     "Armor",
}

_ITEM_TYPE_TO_EFFECT_TYPE = {**_auto_map, **_MANUAL_OVERRIDES}


def _get_effect_list_for_item(item):
    """Return the filtered effect list for this item, or the full list as fallback."""
    hex_id   = swap_endian_hex(item.get("item_id", 0))
    _, itype = lookup_item(hex_id)
    etype    = _ITEM_TYPE_TO_EFFECT_TYPE.get(itype)
    if etype and etype in _effects_by_type:
        return _effects_by_type[etype]
    if itype and itype != "?":
        print("[effect filter] No mapping for item type '{}' (itype={}, etype={})".format(
            itype, itype, etype))
    return effect_dropdown_list


# ---------------------------------------------------------------------------
# Category Effect Icon constants (0x41 per effect, relative to effect base)
# ---------------------------------------------------------------------------
CEI_LABELS = {
    0x00: "Normal (Swappable)",
    0x01: "Locked",
    0x02: "Requires Umbracite",
}

# Effect Extra constants (0x42 per effect, relative to effect base)
EE_LABELS = {
    0x00: "None",
    0x04: "Star Trait",
    0x08: "Smithing Material",
    0x84: "Star Trait + Unknown",  # 132 decimal
}

EQUIPPED_FLAG_OPTIONS = {
    0x00: "Equipped",
    0x11: "Not Equipped",
}

EQUIPMENT_FLAGS1_OPTIONS = {
    0x00: "Normal",
    0x02: "New",
    0x04: "Locked",
    0x06: "New + Locked",
    0x01: "Familiar (check)",
    0x86: "All Flags",
}

EQUIPMENT_FLAGS2_OPTIONS = {
    0x00: "Normal",
    0x02: "Favourite",
    0x10: "Crucible Weapon",
}

RARITY_OPTIONS = {
    0x00: "Common",
    0x01: "Uncommon",
    0x02: "Rare",
    0x03: "Exotic",
    0x04: "Divine",
    0x05: "Ethereal",
}


def find_value_at_offset(section_data, offset, byte_size):
    try:
        value_bytes = section_data[offset:offset + byte_size]
        if len(value_bytes) == byte_size:
            return int.from_bytes(value_bytes, "little")
    except IndexError:
        pass
    return None

def write_le(value, length):
    if isinstance(value, int):
        return value.to_bytes(length, "little")
    elif isinstance(value, (bytes, bytearray)):
        if len(value) != length:
            raise ValueError("Expected {} bytes, got {}".format(length, len(value)))
        return value
    raise TypeError("Cannot convert type {} to bytes".format(type(value)))

def swap_endian_hex(val):
    return "{:04X}".format(((val & 0xFF) << 8) | (val >> 8))

def swap_u16(val):
    return ((val & 0xFF) << 8) | (val >> 8)

def read_u8(buf, offset):
    return buf[offset]

def read_u16(buf, offset):
    return int.from_bytes(buf[offset:offset + 2], "little")

def read_u32(buf, offset):
    return int.from_bytes(buf[offset:offset + 4], "little")

def lookup_item(item_id_hex):
    if item_id_hex in items_json:
        return items_json[item_id_hex]["name"], items_json[item_id_hex]["type"]
    return "Unknown", "?"

# ==================== FILE OPERATIONS ====================
def open_file():
    global data, MODE, decrypted_path
    file_path = filedialog.askopenfilename(
        title="Select Save File",
        filetypes=[("Save Files", "*.BIN"), ("All Files", "*.*")]
    )
    if not file_path:
        return False

    if os.path.basename(file_path) == "SAVEDATA.BIN":
        backup_path = Path(file_path).with_name("SAVEDATA.BIN.BACKUP")
        counter = 1
        while backup_path.exists():
            backup_path = Path(file_path).with_name(f"SAVEDATA.BIN.BACKUP_{counter}")
            counter += 1
        shutil.copy(file_path, backup_path)
        print(f"Backup created: {backup_path}")

        MODE = "PC"
        exe_path = base_dir / "pc" / "pc.exe"
        subprocess.run([str(exe_path), file_path], cwd=exe_path.parent,
                       input="\n", text=True, capture_output=True)
        decrypted_path = exe_path.parent / "decr_SAVEDATA.BIN"
        with open(decrypted_path, "rb") as f:
            data = bytearray(f.read())
        return True

    messagebox.showerror("Error", "Unknown file format. Please select SAVEDATA.BIN (PC)")
    return False


def open_file_import():
    global import_data, IMPORT_MODE, decrypted_path_import
    file_path = filedialog.askopenfilename(
        title="Select Save File to Import",
        filetypes=[("Save Files", "*.BIN"), ("All Files", "*.*")]
    )
    if not file_path:
        return False
    if os.path.basename(file_path) == "SAVEDATA.BIN":
        IMPORT_MODE = "PC"
        exe_path = base_dir / "pc_import" / "pc.exe"
        subprocess.run([str(exe_path), file_path], cwd=exe_path.parent,
                       input="\n", text=True, capture_output=True)
        decrypted_path_import = exe_path.parent / "decr_SAVEDATA.BIN"
        with open(decrypted_path_import, "rb") as f:
            import_data = bytearray(f.read())
        return True
    messagebox.showerror("Error", "Unknown file format. Please select SAVEDATA.BIN (PC)")
    return False


def import_save():
    global data
    if not open_file_import():
        return
    if data is None:
        messagebox.showerror("Error", "Please load your current save file first, then click Import.")
        return
    if messagebox.askyesno("Confirm", "This will replace your current character. Continue?"):
        data = data[:0x15F] + import_data[0x15F:]
        if len(data) != 0x9001B0:
            messagebox.showerror("Import Error", "Size mismatch after import.")
            return
        player_items()
        player_usables()
        player_storage()
        messagebox.showinfo("Success", "File imported. Load in-game to apply the character.")


def save_file():
    global data, decrypted_path
    if data is None or decrypted_path is None:
        messagebox.showwarning("Warning", "No data to save. Load a file first.")
        return
    write_items_to_data()
    write_usables_to_data()
    write_storage_to_data()
    data, checksum = patch_checksum(data)
    print("Final checksum: {:08X}".format(checksum))
    if MODE == "PC":
        with open(decrypted_path, "wb") as f:
            f.write(data)
        exe_path = base_dir / "pc" / "pc.exe"
        subprocess.run([str(exe_path), decrypted_path], cwd=exe_path.parent,
                       input="\n", text=True, capture_output=True)
        last_path = base_dir / "pc" / "decr_decr_SAVEDATA.BIN"
        with open(last_path, "rb") as f:
            final_data = f.read()
        output_path = filedialog.asksaveasfilename(
            initialfile="SAVEDATA.BIN",
            filetypes=[("Save Files", "*.BIN"), ("All Files", "*.*")]
        )
        if output_path:
            with open(output_path, "wb") as f:
                f.write(final_data)
            messagebox.showinfo("Success", "Saved to {}".format(output_path))

# ==================== INVENTORY PARSING ====================
def parse_equipment(offset):
    base = offset
    item = {}
    item["item_id"]               = read_u16(data, base + 0x00)
    item["appearance_id"]         = read_u16(data, base + 0x02)
    item["quantity"]              = read_u16(data, base + 0x04)
    item["item_level"]            = read_u16(data, base + 0x06)
    item["item_level_pre_forge"]  = read_u16(data, base + 0x08)
    item["plus_value"]            = read_u16(data, base + 0x0A)
    item["familiarity_raw"]       = read_u32(data, base + 0x14)
    item["equipment_flags_1"]     = read_u8(data,  base + 0x18)
    item["equipment_flags_2"]     = read_u8(data,  base + 0x1A)
    item["inv_index"]             = read_u16(data, base + 0x1C)
    item["ui_1"]                  = read_u8(data,  base + 0x28)
    item["ui_2"]                  = read_u8(data,  base + 0x29)
    item["rarity"]                = read_u8(data,  base + 0x30)

    effects = []
    se_base = base + 0x38
    for i in range(7):
        eo = se_base + i * 0x08
        effects.append({
            "effect_id":            read_u16(data, eo + 0x00),
            "effect_value":         read_u16(data, eo + 0x04),
            "category_effect_icon": read_u8(data,  eo + 0x09),
            "effect_extra":         read_u8(data,  eo + 0x0A),
        })
    item["effects"] = effects

    item["equipped_flag_1"] = read_u8(data, base + 0xE8)
    item["equipped_flag_2"] = read_u8(data, base + 0xEC)
    item["padding"]         = bytes(data[base + 0xED: base + 0xF0])
    return item


def parse_usable(offset):
    base = offset
    item = {}
    item["item_id"]               = read_u16(data, base + 0x00)
    item["appearance_id"]         = read_u16(data, base + 0x02)
    item["quantity"]              = read_u16(data, base + 0x04)
    item["item_level"]            = read_u16(data, base + 0x06)
    item["item_level_pre_forge"]  = read_u16(data, base + 0x08)
    item["plus_value"]            = read_u16(data, base + 0x0A)
    item["familiarity_raw"]       = read_u32(data, base + 0x14)
    item["equipment_flags_1"]     = read_u8(data,  base + 0x18)
    item["inv_index"]             = read_u16(data, base + 0x1C)
    item["ui_1"]                  = read_u8(data,  base + 0x28)
    item["ui_2"]                  = read_u8(data,  base + 0x29)
    item["rarity"]                = read_u8(data,  base + 0x30)
    effects = []
    se_base = base + 0x38
    for i in range(7):
        eo = se_base + i * 0x08
        effects.append({
            "effect_id":            read_u16(data, eo + 0x00),
            "effect_value":         read_u16(data, eo + 0x04),
            "category_effect_icon": read_u8(data,  eo + 0x09),
            "effect_extra":         read_u8(data,  eo + 0x0A),
        })
    item["effects"] = effects
    return item

def player_items():
    global items
    items = []
    for slot in range(ITEM_SLOTS):
        item = parse_equipment(ITEM_START + slot * ITEM_SIZE)
        item["slot"] = slot
        items.append(item)

def player_usables():
    global usables
    usables = []
    for slot in range(USABLE_SLOTS):
        item = parse_usable(USABLE_START + slot * USABLE_SIZE)
        item["slot"] = slot
        usables.append(item)

def player_storage():
    global storage
    storage = []
    for slot in range(STORAGE_SLOTS):
        item = parse_usable(STORAGE_START + slot * STORAGE_SIZE)
        item["slot"] = slot
        storage.append(item)


def _write_slot(item, base, has_equipped):
    data[base:base+2]         = write_le(item["item_id"],              2)
    data[base+2:base+4]       = write_le(item["appearance_id"],        2)
    data[base+4:base+6]       = write_le(item["quantity"],             2)
    data[base+6:base+8]       = write_le(item["item_level"],           2)
    data[base+8:base+10]      = write_le(item["item_level_pre_forge"], 2)
    data[base+10:base+12]     = write_le(item["plus_value"],           2)
    data[base+0x14:base+0x18] = write_le(item["familiarity_raw"],      4)
    data[base+0x18]            = item["equipment_flags_1"] & 0xFF
    if has_equipped:
        data[base+0x1A]        = item.get("equipment_flags_2", 0) & 0xFF
    data[base+0x1C:base+0x1E] = write_le(item["inv_index"],            2)
    data[base+0x30]            = item["rarity"]            & 0xFF

    se_base = base + 0x38
    for i, eff in enumerate(item["effects"]):
        eo = se_base + i * 0x08
        data[eo:eo+2]       = write_le(eff["effect_id"],            2)
        data[eo+4:eo+6]     = write_le(eff["effect_value"],         2)
        data[eo+0x09]        = eff.get("category_effect_icon", 0) & 0xFF
        data[eo+0x0A]        = eff.get("effect_extra",          0) & 0xFF

    if has_equipped:
        data[base+0xE8] = item["equipped_flag_1"] & 0xFF
        data[base+0xEC] = item["equipped_flag_2"] & 0xFF

def write_items_to_data():
    for item in items:
        _write_slot(item, ITEM_START    + item["slot"] * ITEM_SIZE,    has_equipped=True)

def write_usables_to_data():
    for item in usables:
        _write_slot(item, USABLE_START  + item["slot"] * USABLE_SIZE,  has_equipped=False)

def write_storage_to_data():
    for item in storage:
        _write_slot(item, STORAGE_START + item["slot"] * STORAGE_SIZE, has_equipped=False)


_EQUIP_TEMPLATE_HEX = ('3E AD 3E AD 01 00 A0 00 A0 00 03 00 00 00 00 40 00 00 00 00 00 00 00 00 82 00 00 00 D0 35 00 00 01 00 DB B4 00 00 00 00 B0 7D 0F 00 00 00 00 00 04 00 00 00 4F 68 00 00 04 10 00 00 0F 00 00 00 64 5F 00 00 00 00 00 00 00 00 00 00 E9 9A 00 00 4B 79 00 00 0F 00 00 00 5B 92 80 FF 00 00 00 00 00 00 00 00 3F 87 00 00 26 99 00 00 2A 00 00 00 54 83 00 07 00 00 00 00 00 00 00 00 F7 86 00 00 07 51 00 00 10 00 00 00 57 12 00 00 00 00 58 1B 00 00 00 00 02 2B 00 00 EB E8 00 00 00 00 00 00 00 0C 02 00 00 00 00 00 00 00 00 00 00 00 00 00 FF FF FF FF 00 00 00 00 00 00 00 07 00 00 00 00 00 00 00 00 00 00 8C 3F FF FF FF FF 00 00 00 00 00 00 80 3F 00 00 80 3F 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 11 00 00 00 11 00 00 00')
_EQUIP_TEMPLATE  = bytearray.fromhex(_EQUIP_TEMPLATE_HEX.replace(" ", ""))

_USABLE_TEMPLATE= bytearray.fromhex('7A FC 7A FC 0A 00 00 00 00 00 00 00 00 00 00 40 00 00 00 00 00 00 00 00 00 00 20 00 EC 1D 00 00 00 00 00 00 00 00 00 00 80 B1 02 00 00 00 00 00 02 00 00 00 00 00 00 00 FF FF FF FF 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 FF FF FF FF 00 00 00 00 00 00 00 44 00 00 00 00 00 00 00 00 00 00 00 00 FF FF FF FF 00 00 00 00 00 00 00 5D 00 00 00 00 00 00 00 00 00 00 00 00 FF FF FF FF 00 00 00 00 00 00 00 51 00 00 00 00 00 00 00 00 00 00 00 00 FF FF FF FF 00 00 00 00 00 00 00 80 00 00 00 80 00 00 00 00 00 00 00 80 FF FF FF FF 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 FF FF FF FF 00 00 00 00 00 00 80 51 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00')


def _display_id_to_le_bytes(id_hex_str):
    raw     = int(id_hex_str, 16)
    swapped = ((raw & 0xFF) << 8) | (raw >> 8)
    return swapped.to_bytes(2, "little")


def _find_empty_slot_raw(start_offset, slot_size, slot_count):
    for i in range(slot_count):
        base = start_offset + i * slot_size
        if data[base: base + 4] == b'\x00\x00\x00\x00':
            return i
    return None


def _collect_used_inv_indices():
    used = set()
    for start, size, count in [
        (ITEM_START,    ITEM_SIZE,    ITEM_SLOTS),
        (USABLE_START,  USABLE_SIZE,  USABLE_SLOTS),
        (STORAGE_START, STORAGE_SIZE, STORAGE_SLOTS),
    ]:
        for i in range(count):
            base = start + i * size
            if data[base: base + 4] == b'\x00\x00\x00\x00':
                continue
            idx = read_u16(data, base + 0x1C)
            if idx != 0:
                used.add(idx)
    return used


def _generate_unique_inv_index():
    used = _collect_used_inv_indices()
    import random
    candidates = [
        (hi << 8) | lo
        for hi in range(0x01, 0x100)
        for lo in range(0x01, 0x100)
    ]
    random.shuffle(candidates)
    for val in candidates:
        if val not in used:
            return val
    raise RuntimeError("No unique inventory index available (all 65025 slots used).")


def spawn_equipment(item_name):
    global data
    if data is None:
        raise RuntimeError("No save file loaded.")
    name_to_id = {v["name"]: k for k, v in items_json.items()}
    id_hex = name_to_id.get(item_name)
    if id_hex is None:
        raise ValueError("Item not found: {}".format(item_name))

    slot_idx = _find_empty_slot_raw(ITEM_START, ITEM_SIZE, ITEM_SLOTS)
    if slot_idx is None:
        raise RuntimeError("No empty equipment slot available.")

    slot_bytes      = bytearray(_EQUIP_TEMPLATE)
    le_id           = _display_id_to_le_bytes(id_hex)
    slot_bytes[0:2] = le_id
    slot_bytes[2:4] = le_id
    inv_idx = _generate_unique_inv_index()
    slot_bytes[0x1C:0x1E] = inv_idx.to_bytes(2, "little")

    base = ITEM_START + slot_idx * ITEM_SIZE
    data[base: base + ITEM_SIZE] = slot_bytes

    new_item         = parse_equipment(base)
    new_item["slot"] = slot_idx
    items[slot_idx]  = new_item

    increment_inventory_counter(data)

    return slot_idx


def spawn_usable(item_name):
    global data
    if data is None:
        raise RuntimeError("No save file loaded.")
    name_to_id = {v["name"]: k for k, v in items_json.items()}
    id_hex = name_to_id.get(item_name)
    if id_hex is None:
        raise ValueError("Item not found: {}".format(item_name))

    slot_idx = _find_empty_slot_raw(USABLE_START, USABLE_SIZE, USABLE_SLOTS)
    if slot_idx is None:
        raise RuntimeError("No empty usable slot available.")

    slot_bytes      = bytearray(_USABLE_TEMPLATE)
    le_id           = _display_id_to_le_bytes(id_hex)
    slot_bytes[0:2] = le_id
    slot_bytes[2:4] = le_id
    inv_idx = _generate_unique_inv_index()
    slot_bytes[0x1C:0x1E] = inv_idx.to_bytes(2, "little")

    base = USABLE_START + slot_idx * USABLE_SIZE
    data[base: base + USABLE_SIZE] = slot_bytes

    new_item           = parse_usable(base)
    new_item["slot"]   = slot_idx
    usables[slot_idx]  = new_item

    increment_inventory_counter(data)
    return slot_idx


_effect_name_by_id = {e["id"]: e["Effect"] for e in effects_json}


def _resolve_effect_name(raw_effect_id):
    if raw_effect_id == 0:
        return ""
    display_id = "{:04X}".format(raw_effect_id)
    return _effect_name_by_id.get(display_id, display_id)


def export_to_excel(sheet_name, dataset):
    if not dataset:
        messagebox.showwarning("No Data", "Nothing to export.")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    header = ["Slot", "Item ID", "Name", "Type", "Quantity", "Level", "Rarity"]
    for n in range(1, 8):
        header += [
            "Effect {} (ID - Name)".format(n),
            "Effect {} Value".format(n),
            "Effect {} Category Icon".format(n),
            "Effect {} Extra".format(n),
        ]
    ws.append(header)

    for item in dataset:
        if item["item_id"] == 0:
            continue
        hex_id      = swap_endian_hex(item["item_id"])
        name, type_ = lookup_item(hex_id)

        row = [
            item["slot"],
            hex_id,
            name  or "Unknown",
            type_ or "?",
            item["quantity"],
            item.get("item_level", 0),
            item.get("rarity",     0),
        ]

        for eff in item.get("effects", []):
            raw_id = eff["effect_id"]
            if raw_id == 0:
                row.append("Empty")
                row.append("")
            else:
                display_id  = "{:04X}".format(swap_u16(raw_id))
                effect_name = _effect_name_by_id.get(display_id, display_id)
                row.append(f"{display_id} - {effect_name}")
                row.append(eff["effect_value"])
            # Category Effect Icon
            cei = eff.get("category_effect_icon", 0)
            row.append(CEI_LABELS.get(cei, "0x{:02X}".format(cei)))
            # Effect Extra
            ee = eff.get("effect_extra", 0)
            row.append(EE_LABELS.get(ee, "0x{:02X}".format(ee)))

        ws.append(row)

    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        initialfile="{}.xlsx".format(sheet_name),
        filetypes=[("Excel Files", "*.xlsx")]
    )
    if save_path:
        wb.save(save_path)
        messagebox.showinfo("Success", "{} exported successfully!".format(sheet_name))

# ==================== HELPERS ====================

def _make_flag_combo(parent, row, col, label_text, options_dict):
    """Create a label + OptionMenu pair for a flag field. Returns (var, menu)."""
    tk.Label(parent, text=label_text, anchor="w").grid(row=row, column=col, sticky="w", padx=4, pady=2)
    choices = ["{} – {}".format("0x{:02X}".format(k), v) for k, v in options_dict.items()]
    var = tk.StringVar(value=choices[0])
    menu = ttk.Combobox(parent, textvariable=var, values=choices, width=28, state="readonly")
    menu.grid(row=row, column=col + 1, sticky="w", padx=4, pady=2)
    return var, menu


def _flag_combo_get_int(var, options_dict):
    """Parse the hex value back out of a flag combo string."""
    text = var.get()
    hex_part = text.split("–")[0].strip()
    return int(hex_part, 16)


def _flag_combo_set(var, options_dict, int_val):
    choices = ["{} – {}".format("0x{:02X}".format(k), v) for k, v in options_dict.items()]
    target  = "0x{:02X}".format(int_val)
    for c in choices:
        if c.startswith(target):
            var.set(c)
            return
    # Unknown value — show raw hex
    var.set("0x{:02X} – Unknown".format(int_val))


# ==================== SEARCHABLE COMBOBOX ====================
class SearchableCombobox(ttk.Frame):
    def __init__(self, master=None, values=None, width=40, **kwargs):
        super().__init__(master, **kwargs)
        self.full_values     = values if values else []
        self.filtered_values = self.full_values.copy()
        self._loading        = False

        self.var   = tk.StringVar()
        self.entry = ttk.Entry(self, textvariable=self.var, width=width)
        self.entry.pack(side="left", fill="x", expand=True)
        self.btn = ttk.Button(self, text="v", width=2, command=self.toggle_dropdown)
        self.btn.pack(side="right")

        self.listbox_frame = tk.Toplevel(self)
        self.listbox_frame.withdraw()
        self.listbox_frame.overrideredirect(True)
        self.listbox = tk.Listbox(self.listbox_frame, height=10, width=width)
        self.listbox.pack(fill="both", expand=True)
        sb = ttk.Scrollbar(self.listbox_frame, orient="vertical", command=self.listbox.yview)
        sb.pack(side="right", fill="y")
        self.listbox.config(yscrollcommand=sb.set)

        self.var.trace_add("write", self._on_type)
        self.entry.bind("<Down>",     self._on_arrow_down)
        self.entry.bind("<Up>",       self._on_arrow_up)
        self.entry.bind("<Return>",   self._on_return)
        self.entry.bind("<Escape>",   self._on_escape)
        self.entry.bind("<FocusOut>", self._on_focus_out)
        self.listbox.bind("<<ListboxSelect>>", self._on_select)
        self.listbox.bind("<Return>",           self._on_return)
        self.listbox.bind("<Escape>",           self._on_escape)
        self.listbox.bind("<Double-Button-1>",  self._on_select)
        self.dropdown_visible = False

    def _on_type(self, *args):
        if self._loading:
            return
        typed = self.var.get().lower()
        self.filtered_values = (
            self.full_values.copy() if typed == ""
            else [v for v in self.full_values if typed in v.lower()]
        )
        self._update_listbox()
        if not self.dropdown_visible and self.filtered_values:
            self.show_dropdown()

    def _update_listbox(self):
        self.listbox.delete(0, tk.END)
        for value in self.filtered_values:
            self.listbox.insert(tk.END, value)

    def show_dropdown(self):
        if not self.filtered_values:
            return
        self.dropdown_visible = True
        x     = self.entry.winfo_rootx()
        y     = self.entry.winfo_rooty() + self.entry.winfo_height()
        width = self.entry.winfo_width() + self.btn.winfo_width()
        self.listbox_frame.geometry("{}x200+{}+{}".format(width, x, y))
        self.listbox_frame.deiconify()
        self.listbox_frame.lift()

    def hide_dropdown(self):
        self.dropdown_visible = False
        self.listbox_frame.withdraw()

    def toggle_dropdown(self):
        if self.dropdown_visible:
            self.hide_dropdown()
        else:
            self.filtered_values = self.full_values.copy()
            self._update_listbox()
            self.show_dropdown()
            self.entry.focus_set()

    def _on_arrow_down(self, event):
        if not self.dropdown_visible:
            self.show_dropdown()
        else:
            cur = self.listbox.curselection()
            if not cur:
                self.listbox.selection_set(0)
            elif cur[0] < self.listbox.size() - 1:
                self.listbox.selection_clear(cur)
                self.listbox.selection_set(cur[0] + 1)
                self.listbox.see(cur[0] + 1)
        return "break"

    def _on_arrow_up(self, event):
        if self.dropdown_visible:
            cur = self.listbox.curselection()
            if cur and cur[0] > 0:
                self.listbox.selection_clear(cur)
                self.listbox.selection_set(cur[0] - 1)
                self.listbox.see(cur[0] - 1)
        return "break"

    def _on_return(self, event):
        if self.dropdown_visible:
            cur = self.listbox.curselection()
            if cur:
                self._loading = True
                self.var.set(self.listbox.get(cur[0]))
                self._loading = False
            self.hide_dropdown()
        return "break"

    def _on_escape(self, event):
        self.hide_dropdown()
        return "break"

    def _on_select(self, event):
        cur = self.listbox.curselection()
        if cur:
            self._loading = True
            self.var.set(self.listbox.get(cur[0]))
            self._loading = False
            self.hide_dropdown()

    def _on_focus_out(self, event):
        self.after(200, lambda: self.hide_dropdown() if not self.listbox.focus_get() else None)

    def get(self):
        return self.var.get()

    def set(self, value):
        self._loading = True
        self.var.set(value)
        self._loading = False

    def configure(self, **kwargs):
        if "values" in kwargs:
            self.full_values     = list(kwargs.pop("values"))
            self.filtered_values = self.full_values.copy()
        if kwargs:
            self.entry.configure(**kwargs)

    def __setitem__(self, key, value):
        if key == "values":
            self.full_values     = list(value)
            self.filtered_values = self.full_values.copy()

    def __getitem__(self, key):
        return self.full_values if key == "values" else None

# ==================== ITEM EDITOR ====================
class ItemEditor(ttk.Frame):
    """
    Edits a single equipment or usable slot.

    Equipment has extra fields: equipment_flags_2, equipped_flag_1/2.
    Both types now also show per-effect Category Icon and Effect Extra dropdowns.
    """

    # Simple numeric entry fields
    _EQUIP_PROPS = [
        ("item_id",              "Item ID (hex LE)"),
        ("appearance_id",        "Appearance ID (hex LE)"),
        ("quantity",             "Quantity"),
        ("item_level",           "Item Level  [max 170]"),
        ("item_level_pre_forge", "Item Level (Pre-Forge)"),
        ("plus_value",           "Plus Value  [max +15]"),
        ("familiarity_raw",      "Familiarity (raw u32)"),
        ("inv_index",            "Placement ID (inv index)"),
        ("ui_1",                 "UI Object 1  (0x28)"),
        ("ui_2",                 "UI Object 2  (0x29)"),
    ]
    _USABLE_PROPS = [
        ("item_id",              "Item ID (hex LE)"),
        ("appearance_id",        "Appearance ID (hex LE)"),
        ("quantity",             "Quantity"),
        ("item_level",           "Item Level"),
        ("item_level_pre_forge", "Item Level (Pre-Forge)"),
        ("plus_value",           "Plus Value"),
        ("familiarity_raw",      "Familiarity (raw u32)"),
        ("inv_index",            "Placement ID (inv index)"),
        ("ui_1",                 "UI Object 1  (0x28)"),
        ("ui_2",                 "UI Object 2  (0x29)"),
    ]

    def __init__(self, master, has_equipped, on_apply, **kwargs):
        super().__init__(master, **kwargs)
        self._on_apply    = on_apply
        self._has_equipped = has_equipped
        self._props       = self._EQUIP_PROPS if has_equipped else self._USABLE_PROPS
        self._build()

    # ------------------------------------------------------------------
    def _build(self):
        canvas = tk.Canvas(self)
        sb     = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        frame  = ttk.Frame(canvas)
        frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=frame, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # ── Properties ────────────────────────────────────────────────
        self._entries = {}
        pf = ttk.LabelFrame(frame, text="Properties", padding=8)
        pf.pack(fill="x", padx=5, pady=5)

        for i, (key, lbl) in enumerate(self._props):
            ttk.Label(pf, text=lbl).grid(row=i, column=0, sticky="w", padx=4, pady=2)
            e = ttk.Entry(pf, width=18)
            e.grid(row=i, column=1, sticky="w", padx=4, pady=2)
            self._entries[key] = e

        row_offset = len(self._props)

        # Rarity dropdown
        ttk.Label(pf, text="Rarity").grid(row=row_offset, column=0, sticky="w", padx=4, pady=2)
        self._rarity_var = tk.StringVar()
        rarity_choices   = ["{} – {}".format("0x{:02X}".format(k), v)
                            for k, v in RARITY_OPTIONS.items()]
        self._rarity_cb  = ttk.Combobox(pf, textvariable=self._rarity_var,
                                         values=rarity_choices, width=28, state="readonly")
        self._rarity_cb.grid(row=row_offset, column=1, sticky="w", padx=4, pady=2)
        row_offset += 1

        # Equipment Flags 1
        ttk.Label(pf, text="Equipment Flags 1  (0x18)").grid(
            row=row_offset, column=0, sticky="w", padx=4, pady=2)
        self._flags1_var = tk.StringVar()
        flags1_choices   = ["{} – {}".format("0x{:02X}".format(k), v)
                            for k, v in EQUIPMENT_FLAGS1_OPTIONS.items()]
        self._flags1_cb  = ttk.Combobox(pf, textvariable=self._flags1_var,
                                          values=flags1_choices, width=28, state="readonly")
        self._flags1_cb.grid(row=row_offset, column=1, sticky="w", padx=4, pady=2)
        row_offset += 1

        if self._has_equipped:
            # Equipment Flags 2 (favourite / crucible)
            ttk.Label(pf, text="Equipment Flags 2  (0x1A)").grid(
                row=row_offset, column=0, sticky="w", padx=4, pady=2)
            self._flags2_var = tk.StringVar()
            flags2_choices   = ["{} – {}".format("0x{:02X}".format(k), v)
                                for k, v in EQUIPMENT_FLAGS2_OPTIONS.items()]
            self._flags2_cb  = ttk.Combobox(pf, textvariable=self._flags2_var,
                                              values=flags2_choices, width=28, state="readonly")
            self._flags2_cb.grid(row=row_offset, column=1, sticky="w", padx=4, pady=2)
            row_offset += 1

            # Equipped flag 1
            ttk.Label(pf, text="Equipped Flag 1  (0xE8)").grid(
                row=row_offset, column=0, sticky="w", padx=4, pady=2)
            self._eq1_var = tk.StringVar()
            eq_choices    = ["{} – {}".format("0x{:02X}".format(k), v)
                             for k, v in EQUIPPED_FLAG_OPTIONS.items()]
            self._eq1_cb  = ttk.Combobox(pf, textvariable=self._eq1_var,
                                           values=eq_choices, width=28, state="readonly")
            self._eq1_cb.grid(row=row_offset, column=1, sticky="w", padx=4, pady=2)
            row_offset += 1

            # Equipped flag 2
            ttk.Label(pf, text="Equipped Flag 2  (0xEC)").grid(
                row=row_offset, column=0, sticky="w", padx=4, pady=2)
            self._eq2_var = tk.StringVar()
            self._eq2_cb  = ttk.Combobox(pf, textvariable=self._eq2_var,
                                           values=eq_choices, width=28, state="readonly")
            self._eq2_cb.grid(row=row_offset, column=1, sticky="w", padx=4, pady=2)

        # ── Effects ───────────────────────────────────────────────────
        ef = ttk.LabelFrame(frame, text="Effects  (7 slots)", padding=8)
        ef.pack(fill="x", padx=5, pady=5)

        # Column headers
        for col, hdr in enumerate(["Slot", "Effect ID – Name", "Value",
                                    "Category Icon  (0x41)", "Extra  (0x42)"]):
            ttk.Label(ef, text=hdr, font=("TkDefaultFont", 8, "bold")).grid(
                row=0, column=col, sticky="w", padx=4, pady=2)

        self._eff_combos = []
        self._eff_mags   = []
        self._eff_cei    = []   # category_effect_icon combobox vars
        self._eff_ee     = []   # effect_extra combobox vars

        cei_choices = ["{} – {}".format("0x{:02X}".format(k), v)
                       for k, v in CEI_LABELS.items()]
        ee_choices  = ["{} – {}".format("0x{:02X}".format(k), v)
                       for k, v in EE_LABELS.items()]

        for i in range(7):
            r = i + 1
            ttk.Label(ef, text="{}:".format(i + 1)).grid(row=r, column=0, padx=4, pady=2)

            combo = SearchableCombobox(ef, width=34, values=effect_dropdown_list)
            combo.grid(row=r, column=1, sticky="w", padx=4, pady=2)
            self._eff_combos.append(combo)

            mag = ttk.Entry(ef, width=8)
            mag.grid(row=r, column=2, sticky="w", padx=4, pady=2)
            self._eff_mags.append(mag)

            cei_var = tk.StringVar(value=cei_choices[0])
            cei_cb  = ttk.Combobox(ef, textvariable=cei_var,
                                    values=cei_choices, width=26, state="readonly")
            cei_cb.grid(row=r, column=3, sticky="w", padx=4, pady=2)
            self._eff_cei.append(cei_var)

            ee_var = tk.StringVar(value=ee_choices[0])
            ee_cb  = ttk.Combobox(ef, textvariable=ee_var,
                                   values=ee_choices, width=26, state="readonly")
            ee_cb.grid(row=r, column=4, sticky="w", padx=4, pady=2)
            self._eff_ee.append(ee_var)

        # ── Apply button ──────────────────────────────────────────────
        bf = ttk.Frame(frame)
        bf.pack(pady=8)
        ttk.Button(bf, text="Apply Changes", command=self._do_apply).pack()

    # ------------------------------------------------------------------
    def _do_apply(self):
        self._on_apply(
            self._entries,
            self._eff_combos,
            self._eff_mags,
            self._eff_cei,
            self._eff_ee,
            self._rarity_var,
            self._flags1_var,
            getattr(self, "_flags2_var", None),
            getattr(self, "_eq1_var",    None),
            getattr(self, "_eq2_var",    None),
        )

    # ------------------------------------------------------------------
    @staticmethod
    def _parse_combo_hex(var):
        """Extract the leading 0xNN int from a flag/option combo string."""
        text = var.get()
        hex_part = text.split("–")[0].strip()
        try:
            return int(hex_part, 16)
        except ValueError:
            return 0

    @staticmethod
    def _set_combo(var, cb, options_dict, int_val):
        choices = ["{} – {}".format("0x{:02X}".format(k), v)
                   for k, v in options_dict.items()]
        target  = "0x{:02X}".format(int_val)
        for c in choices:
            if c.startswith(target):
                var.set(c)
                return
        var.set("0x{:02X} – Unknown".format(int_val))

    # ------------------------------------------------------------------
    def load(self, item):
        # Numeric fields
        for key, entry in self._entries.items():
            entry.delete(0, tk.END)
            entry.insert(0, item.get(key, 0))

        # Rarity
        self._set_combo(self._rarity_var, self._rarity_cb,
                        RARITY_OPTIONS, item.get("rarity", 0))

        # Flags 1
        self._set_combo(self._flags1_var, self._flags1_cb,
                        EQUIPMENT_FLAGS1_OPTIONS, item.get("equipment_flags_1", 0))

        if self._has_equipped:
            self._set_combo(self._flags2_var, self._flags2_cb,
                            EQUIPMENT_FLAGS2_OPTIONS, item.get("equipment_flags_2", 0))
            self._set_combo(self._eq1_var, self._eq1_cb,
                            EQUIPPED_FLAG_OPTIONS, item.get("equipped_flag_1", 0))
            self._set_combo(self._eq2_var, self._eq2_cb,
                            EQUIPPED_FLAG_OPTIONS, item.get("equipped_flag_2", 0))

        # Effect dropdowns filtered to item type
        eff_list    = _get_effect_list_for_item(item)
        cei_choices = ["{} – {}".format("0x{:02X}".format(k), v)
                       for k, v in CEI_LABELS.items()]
        ee_choices  = ["{} – {}".format("0x{:02X}".format(k), v)
                       for k, v in EE_LABELS.items()]

        for combo in self._eff_combos:
            combo.configure(values=eff_list)

        for i, eff in enumerate(item.get("effects", [])):
            raw_id = eff["effect_id"]
            if raw_id == 0:
                self._eff_combos[i].set("Empty")
            else:
                display_id = "{:04X}".format(raw_id)
                matched    = next((c for c in eff_list if c.startswith(display_id)), display_id)
                self._eff_combos[i].set(matched)

            self._eff_mags[i].delete(0, tk.END)
            self._eff_mags[i].insert(0, eff["effect_value"])

            # Category Effect Icon
            cei_val = eff.get("category_effect_icon", 0)
            self._set_combo(self._eff_cei[i], None, CEI_LABELS, cei_val)

            # Effect Extra
            ee_val  = eff.get("effect_extra", 0)
            self._set_combo(self._eff_ee[i], None, EE_LABELS, ee_val)

    def clear(self):
        for e in self._entries.values():
            e.delete(0, tk.END)
        for combo in self._eff_combos:
            combo.configure(values=effect_dropdown_list)
            combo.set("")
        for mag in self._eff_mags:
            mag.delete(0, tk.END)

# ==================== SPAWN DIALOG ====================
class SpawnDialog(tk.Toplevel):
    def __init__(self, master, title_label, panel, allowed_types=None):
        super().__init__(master)
        self._panel         = panel
        self._allowed_types = allowed_types
        self.title("Spawn {}".format(title_label))
        self.geometry("520x480")
        self.resizable(False, True)
        self.grab_set()

        if allowed_types is not None:
            self._names = sorted(
                v["name"] for v in items_json.values()
                if v.get("type", "") in allowed_types and v.get("name", "")
            )
        else:
            self._names = sorted(
                v["name"] for v in items_json.values()
                if v.get("name", "")
            )

        self._build()

    def _build(self):
        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")
        ttk.Label(top, text="Search:").pack(side="left")
        self._search_var = tk.StringVar()
        se = ttk.Entry(top, textvariable=self._search_var, width=36)
        se.pack(side="left", padx=6)
        se.bind("<KeyRelease>", lambda e: self._filter())
        se.focus_set()

        list_frame = ttk.Frame(self)
        list_frame.pack(fill="both", expand=True, padx=8, pady=4)
        self._listbox = tk.Listbox(list_frame, activestyle="dotbox", selectmode="single")
        sb = ttk.Scrollbar(list_frame, orient="vertical", command=self._listbox.yview)
        self._listbox.configure(yscrollcommand=sb.set)
        self._listbox.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self._listbox.bind("<Double-Button-1>", lambda e: self._spawn())
        self._listbox.bind("<Return>",          lambda e: self._spawn())

        btn_frame = ttk.Frame(self, padding=(0, 0, 8, 8))
        btn_frame.pack(fill="x", side="bottom")
        ttk.Button(btn_frame, text="Cancel", command=self.destroy).pack(side="right", padx=4)
        ttk.Button(btn_frame, text="Spawn",  command=self._spawn).pack(side="right")

        self._populate(self._names)

    def _populate(self, names):
        self._listbox.delete(0, tk.END)
        for n in names:
            self._listbox.insert(tk.END, n)

    def _filter(self):
        q = self._search_var.get().lower()
        self._populate([n for n in self._names if q in n.lower()] if q else self._names)

    def _spawn(self):
        sel = self._listbox.curselection()
        if not sel:
            messagebox.showwarning("No Selection", "Select an item to spawn.", parent=self)
            return
        item_name = self._listbox.get(sel[0])
        try:
            if self._panel._has_equipped:
                slot = spawn_equipment(item_name)
            else:
                slot = spawn_usable(item_name)
            self._panel.refresh()
            messagebox.showinfo("Spawned",
                                "Spawned \"{}\" into slot {}.".format(item_name, slot),
                                parent=self)
            self.destroy()
        except (ValueError, RuntimeError) as exc:
            messagebox.showerror("Spawn Failed", str(exc), parent=self)

import struct

def increment_inventory_counter(data: bytearray, offset=0x33E515):
    current = struct.unpack_from("<I", data, offset)[0]

    # Optional: prevent overflow (uint32 max)
    if current >= 0xFFFFFFFF:
        current = 0

    new_value = current + 1
    struct.pack_into("<I", data, offset, new_value)

    return new_value
# ==================== IMPORT INVENTORY DIALOG ====================
class ImportInventoryDialog(tk.Toplevel):
    """
    Lets the user pick a source save file, choose which inventory
    sections to copy (Equipment / Consumables / Storage), then
    applies the raw byte blocks into the current save's data buffer.

    All inv_index values from the source are remapped to unique IDs
    that don't collide with anything already in the current save so
    the game never sees duplicate placement IDs.
    """

    # Section metadata: (label, src_start, dst_start, slot_size, slot_count)
    _SECTIONS = [
        ("Equipment (Weapons & Armour)",
         ITEM_START,    ITEM_START,    ITEM_SIZE,    ITEM_SLOTS),
        ("Consumables / Materials",
         USABLE_START,  USABLE_START,  USABLE_SIZE,  USABLE_SLOTS),
        ("Storage Box",
         STORAGE_START, STORAGE_START, STORAGE_SIZE, STORAGE_SLOTS),
    ]

    def __init__(self, master, app_ref):
        super().__init__(master)
        self._app  = app_ref
        self.title("Import Inventory from Save")
        self.geometry("480x320")
        self.resizable(False, False)
        self.grab_set()
        self._src_path_var = tk.StringVar(value="No file selected")
        self._src_data     = None
        self._checks       = []
        self._build()

    # ------------------------------------------------------------------
    def _build(self):
        pad = dict(padx=12, pady=6)

        # ── File picker row ──────────────────────────────────────────
        file_frame = ttk.LabelFrame(self, text="Source Save File", padding=8)
        file_frame.pack(fill="x", **pad)

        ttk.Label(file_frame, textvariable=self._src_path_var,
                  width=44, anchor="w", relief="sunken").pack(side="left", padx=(0, 6))
        ttk.Button(file_frame, text="Browse…",
                   command=self._browse).pack(side="left")

        # ── Section checkboxes ───────────────────────────────────────
        sec_frame = ttk.LabelFrame(self, text="Sections to Import", padding=8)
        sec_frame.pack(fill="x", **pad)

        self._checks = []
        for label, *_ in self._SECTIONS:
            var = tk.BooleanVar(value=True)
            cb  = ttk.Checkbutton(sec_frame, text=label, variable=var)
            cb.pack(anchor="w", pady=2)
            self._checks.append(var)

        # "Select all / none" convenience buttons
        btn_row = ttk.Frame(sec_frame)
        btn_row.pack(anchor="w", pady=(4, 0))
        ttk.Button(btn_row, text="Select All",
                   command=lambda: [v.set(True)  for v in self._checks]
                   ).pack(side="left", padx=(0, 4))
        ttk.Button(btn_row, text="Select None",
                   command=lambda: [v.set(False) for v in self._checks]
                   ).pack(side="left")

        # ── Action buttons ───────────────────────────────────────────
        act_frame = ttk.Frame(self)
        act_frame.pack(fill="x", side="bottom", padx=12, pady=10)
        ttk.Button(act_frame, text="Cancel",
                   command=self.destroy).pack(side="right", padx=(4, 0))
        ttk.Button(act_frame, text="Import",
                   command=self._do_import).pack(side="right")

    # ------------------------------------------------------------------
    def _browse(self):
        """Decrypt and load a source SAVEDATA.BIN into self._src_data."""
        file_path = filedialog.askopenfilename(
            title="Select Source Save File",
            filetypes=[("Save Files", "*.BIN"), ("All Files", "*.*")]
        )
        if not file_path:
            return
        if os.path.basename(file_path) != "SAVEDATA.BIN":
            messagebox.showerror("Error",
                                 "Please select a SAVEDATA.BIN file.",
                                 parent=self)
            return

        exe_path  = base_dir / "pc_import" / "pc.exe"
        subprocess.run([str(exe_path), file_path], cwd=exe_path.parent,
                       input="\n", text=True, capture_output=True)
        dec_path  = exe_path.parent / "decr_SAVEDATA.BIN"
        try:
            with open(dec_path, "rb") as f:
                self._src_data = bytearray(f.read())
            self._src_path_var.set(os.path.basename(file_path) +
                                   "  ✓  ({:,} bytes)".format(len(self._src_data)))
        except FileNotFoundError:
            messagebox.showerror("Error",
                                 "Decryption failed – decr_SAVEDATA.BIN not found.",
                                 parent=self)
            self._src_data = None

    # ------------------------------------------------------------------
    def _remap_inv_indices(self, start, slot_size, slot_count):
        """
        After copying raw bytes from the source, walk every non-empty slot
        in [start … start + slot_size*slot_count) of the *current* data
        buffer and replace each inv_index (u16 LE @ +0x1C) with a freshly
        generated unique value so there are no collisions with items that
        were already in the destination save.
        """
        for i in range(slot_count):
            base = start + i * slot_size
            if data[base: base + 4] == b'\x00\x00\x00\x00':
                continue
            new_idx = _generate_unique_inv_index()
            data[base + 0x1C: base + 0x1E] = new_idx.to_bytes(2, "little")

    # ------------------------------------------------------------------
    def _do_import(self):
        if data is None:
            messagebox.showerror("Error",
                                 "No save file loaded. Please open your current save first.",
                                 parent=self)
            return
        if self._src_data is None:
            messagebox.showerror("Error",
                                 "No source file selected. Click Browse… first.",
                                 parent=self)
            return

        selected = [i for i, v in enumerate(self._checks) if v.get()]
        if not selected:
            messagebox.showwarning("Nothing Selected",
                                   "Please tick at least one section to import.",
                                   parent=self)
            return

        # Build a human-readable summary for the confirmation dialog
        labels = [self._SECTIONS[i][0] for i in selected]
        summary = "\n  • ".join([""] + labels)
        if not messagebox.askyesno(
                "Confirm Import",
                "The following sections will be REPLACED in your current save:{}\n\n"
                "This cannot be undone (your original save backup was made on load).\n\n"
                "Continue?".format(summary),
                parent=self):
            return

        # ── Perform the copy ─────────────────────────────────────────
        imported = []
        for idx in selected:
            label, src_start, dst_start, slot_size, slot_count = self._SECTIONS[idx]
            byte_len = slot_size * slot_count

            # Bounds-check both buffers
            if src_start + byte_len > len(self._src_data):
                messagebox.showerror("Error",
                                     "Source file is too small for section:\n{}".format(label),
                                     parent=self)
                return
            if dst_start + byte_len > len(data):
                messagebox.showerror("Error",
                                     "Destination file is too small for section:\n{}".format(label),
                                     parent=self)
                return

            # Raw block copy
            data[dst_start: dst_start + byte_len] = \
                self._src_data[src_start: src_start + byte_len]

            # Remap inv_index values to avoid collisions
            self._remap_inv_indices(dst_start, slot_size, slot_count)
            imported.append(label)

        # ── Re-parse affected sections and rebuild UI panels ─────────
        rebuild_keys = []
        for idx in selected:
            label = self._SECTIONS[idx][0]
            if "Equipment" in label:
                player_items()
                rebuild_keys.append("equipment")
            elif "Consumables" in label:
                player_usables()
                rebuild_keys.append("usables")
            elif "Storage" in label:
                player_storage()
                rebuild_keys.append("storage")

        app = self._app
        if "equipment" in rebuild_keys:
            app._rebuild_panel("equipment", app._tab_equipment,
                               items,   write_items_to_data,   True,
                               "Equipment", use_subtabs=True)
        if "usables" in rebuild_keys:
            app._rebuild_panel("usables",  app._tab_usables,
                               usables, write_usables_to_data, False,
                               "Consumable", use_subtabs=False)
        if "storage" in rebuild_keys:
            app._rebuild_panel("storage",  app._tab_storage,
                               storage, write_storage_to_data, False,
                               "Storage Item", use_subtabs=False)

        messagebox.showinfo(
            "Import Complete",
            "Successfully imported:\n  • {}\n\n"
            "Remember to Save File to write the changes to disk.".format(
                "\n  • ".join(imported)),
            parent=self)
        self.destroy()


# ==================== INVENTORY PANEL ====================
class InventoryPanel(ttk.Frame):
    def __init__(self, master, dataset, write_fn,
                 has_equipped=True, label="Item", use_subtabs=False, **kwargs):
        super().__init__(master, **kwargs)
        self._dataset      = dataset
        self._write_fn     = write_fn
        self._has_equipped = has_equipped
        self._label        = label
        self._use_subtabs  = use_subtabs
        self._sel_index    = None
        self._active_tree  = None

        paned = ttk.PanedWindow(self, orient="horizontal")
        paned.pack(fill="both", expand=True)
        left  = ttk.Frame(paned)
        right = ttk.Frame(paned)
        paned.add(left,  weight=2)
        paned.add(right, weight=1)

        self._editor = ItemEditor(right, has_equipped, self._apply)
        self._editor.pack(fill="both", expand=True)

        if use_subtabs:
            self._build_subtab_layout(left)
        else:
            self._active_tree = self._build_tree_pane(left, type_filter=None)

    def _build_subtab_layout(self, parent):
        type_order, seen = [], set()
        for item in self._dataset:
            if item["item_id"] == 0:
                continue
            _, type_ = lookup_item(swap_endian_hex(item["item_id"]))
            if type_ not in seen:
                seen.add(type_)
                type_order.append(type_)

        nb = ttk.Notebook(parent)
        nb.pack(fill="both", expand=True)
        self._sub_trees = {}
        known   = sorted(t for t in type_order if t != "?")
        unknown = ["?"] if "?" in type_order else []
        for tab_type in ["All"] + known + unknown:
            frame = ttk.Frame(nb)
            nb.add(frame, text=tab_type)
            self._sub_trees[tab_type] = self._build_tree_pane(frame, type_filter=tab_type)

        self._active_tree = self._sub_trees["All"]
        nb.bind("<<NotebookTabChanged>>", lambda e, n=nb: self._on_subtab_change(n))

    def _on_subtab_change(self, nb):
        tab_text = nb.tab(nb.select(), "text")
        self._active_tree = self._sub_trees.get(tab_text, self._active_tree)

    def _build_tree_pane(self, parent, type_filter):
        fbar = ttk.Frame(parent)
        fbar.pack(fill="x", padx=5, pady=5)
        ttk.Label(fbar, text="Filter:").pack(side="left")
        fvar = tk.StringVar()
        fe   = ttk.Entry(fbar, textvariable=fvar, width=26)
        fe.pack(side="left", padx=4)

        cols = ("slot", "item_id", "name", "type", "quantity", "level", "rarity", "flags1", "flags2")
        tree = ttk.Treeview(parent, columns=cols, show="headings", height=24)
        vsb  = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        headers = {
            "slot":     ("Slot",     45),
            "item_id":  ("Item ID",  70),
            "name":     ("Name",    200),
            "type":     ("Type",    110),
            "quantity": ("Qty",      50),
            "level":    ("Level",    50),
            "rarity":   ("Rarity",   70),
            "flags1":   ("Flags1",   70),
            "flags2":   ("Flags2",   70),
        }
        for col, (hdr, w) in headers.items():
            tree.heading(col, text=hdr)
            tree.column(col, width=w)

        def populate():
            tree.delete(*tree.get_children())
            ft = fvar.get().lower()
            for item in self._dataset:
                if item["item_id"] == 0:
                    continue
                hex_id = swap_endian_hex(item["item_id"])
                name, type_ = lookup_item(hex_id)
                if type_filter and type_filter != "All" and type_ != type_filter:
                    continue
                if ft and ft not in name.lower() and ft not in type_.lower():
                    continue
                rarity_str = RARITY_OPTIONS.get(item.get("rarity", 0),
                                                "0x{:02X}".format(item.get("rarity", 0)))
                f1 = item.get("equipment_flags_1", 0)
                f2 = item.get("equipment_flags_2", 0) if self._has_equipped else "-"
                flags1_str = EQUIPMENT_FLAGS1_OPTIONS.get(f1, "0x{:02X}".format(f1))
                flags2_str = (EQUIPMENT_FLAGS2_OPTIONS.get(f2, "0x{:02X}".format(f2))
                              if self._has_equipped else "-")
                tree.insert("", "end", iid=item["slot"], values=(
                    item["slot"], hex_id, name, type_,
                    item["quantity"], item["item_level"],
                    rarity_str, flags1_str, flags2_str,
                ))

        fe.bind("<KeyRelease>", lambda e: populate())
        ttk.Button(fbar, text="Clear",
                   command=lambda: [fvar.set(""), populate()]
                   ).pack(side="left", padx=4)

        if self._has_equipped:
            equip_types = frozenset(
                lookup_item(swap_endian_hex(it["item_id"]))[1]
                for it in self._dataset if it["item_id"] != 0
            )
            ttk.Button(
                fbar, text="Spawn Equipment (WIP)",
                command=lambda t=equip_types: SpawnDialog(
                    self.winfo_toplevel(), "Equipment", self, allowed_types=t)
            ).pack(side="left", padx=6)
        else:
            usable_types = frozenset(
                lookup_item(swap_endian_hex(it["item_id"]))[1]
                for it in self._dataset if it["item_id"] != 0
            )
            ttk.Button(
                fbar, text="Spawn Item (WIP)",
                command=lambda t=usable_types: SpawnDialog(
                    self.winfo_toplevel(), "Consumable / Material", self, allowed_types=t)
            ).pack(side="left", padx=6)

        tree._populate = populate
        tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        populate()

        bbar = ttk.Frame(parent)
        bbar.pack(fill="x", padx=5, pady=4)
        if not self._has_equipped:
            ttk.Button(bbar, text="Max Quantity", command=self._max_all).pack(side="left", padx=4)

        return tree

    def _on_tree_select(self, event):
        tree = event.widget
        sel  = tree.selection()
        if not sel:
            return
        self._sel_index   = int(sel[0])
        self._active_tree = tree
        self._editor.load(self._dataset[self._sel_index])

    def refresh(self):
        if self._use_subtabs:
            for tree in self._sub_trees.values():
                tree._populate()
        else:
            if self._active_tree:
                self._active_tree._populate()

    def _max_all(self):
        if messagebox.askyesno("Confirm", "Set all {} quantities to 9999?".format(self._label)):
            count = 0
            for item in self._dataset:
                if item["item_id"] != 0 and item["quantity"] != 0:
                    item["quantity"] = 9999
                    count += 1
            self.refresh()
            messagebox.showinfo("Done", "Maxed {} {}(s) to 9999.".format(count, self._label))

    def _apply(self, entries, eff_combos, eff_mags,
               eff_cei, eff_ee,
               rarity_var, flags1_var, flags2_var, eq1_var, eq2_var):
        if self._sel_index is None:
            messagebox.showwarning("No Selection", "Select a {} first.".format(self._label))
            return
        item = self._dataset[self._sel_index]

        # Numeric fields
        for key, entry in entries.items():
            try:
                item[key] = int(entry.get())
            except ValueError:
                messagebox.showerror("Error", "Invalid value for '{}'".format(key))
                return

        # Flag / dropdown fields
        item["rarity"]           = ItemEditor._parse_combo_hex(rarity_var)
        item["equipment_flags_1"] = ItemEditor._parse_combo_hex(flags1_var)
        if self._has_equipped and flags2_var is not None:
            item["equipment_flags_2"] = ItemEditor._parse_combo_hex(flags2_var)
        if self._has_equipped and eq1_var is not None:
            item["equipped_flag_1"] = ItemEditor._parse_combo_hex(eq1_var)
        if self._has_equipped and eq2_var is not None:
            item["equipped_flag_2"] = ItemEditor._parse_combo_hex(eq2_var)

        # Effects
        for i in range(7):
            chosen = eff_combos[i].get()
            if chosen and chosen != "Empty":
                try:
                    val = int(chosen.split(" ")[0], 16)
                    item["effects"][i]["effect_id"] = val
                except (ValueError, IndexError):
                    pass
            mag_val = eff_mags[i].get()
            if mag_val:
                try:
                    item["effects"][i]["effect_value"] = int(mag_val)
                except ValueError:
                    messagebox.showerror("Error", "Invalid effect value for slot {}".format(i + 1))
                    return

            item["effects"][i]["category_effect_icon"] = ItemEditor._parse_combo_hex(eff_cei[i])
            item["effects"][i]["effect_extra"]          = ItemEditor._parse_combo_hex(eff_ee[i])

        self.refresh()
        messagebox.showinfo("Success", "{} updated!".format(self._label))

# ==================== MAIN GUI ====================
class Nioh3Editor:
    def __init__(self, root):
        self.root = root
        self.root.title("Nioh 3 Save Editor")
        self.root.geometry("1400x780")
        ttk.Style().theme_use("clam")

        menubar   = tk.Menu(root)
        root.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open Save File", command=self.load_file)
        file_menu.add_command(label="Save File",      command=save_file)
        file_menu.add_separator()
        file_menu.add_command(label="Exit",           command=root.quit)

        export_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Export to Excel", menu=export_menu)
        export_menu.add_command(label="Export Equipment",
                                command=lambda: export_to_excel("Equipment",   items))
        export_menu.add_command(label="Export Consumables",
                                command=lambda: export_to_excel("Consumables", usables))
        export_menu.add_command(label="Export Storage",
                                command=lambda: export_to_excel("Storage",     storage))

        import_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Import Save", menu=import_menu)
        import_menu.add_command(label="Import Full Save",      command=self.import_save)


        import_inv_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Import Inventory", menu=import_inv_menu)
        import_inv_menu.add_command(label="Import Inventory …",   command=self.import_inventory)

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=5, pady=5)

        self._tab_equipment = ttk.Frame(self.notebook)
        self._tab_usables   = ttk.Frame(self.notebook)
        self._tab_storage   = ttk.Frame(self.notebook)
        self.notebook.add(self._tab_equipment, text="Equipment")
        self.notebook.add(self._tab_usables,   text="Consumables / Materials")
        self.notebook.add(self._tab_storage,   text="Storage Box")

        self._panels = {}
        self.create_stats_tab()
        self.file_loaded = False

    def load_file(self):
        if not open_file():
            return
        player_items()
        player_usables()
        player_storage()
        self._rebuild_panel("equipment", self._tab_equipment,
                            items,   write_items_to_data,    True,  "Equipment",    use_subtabs=True)
        self._rebuild_panel("usables",  self._tab_usables,
                            usables, write_usables_to_data,  False, "Consumable",   use_subtabs=False)
        self._rebuild_panel("storage",  self._tab_storage,
                            storage, write_storage_to_data,  False, "Storage Item", use_subtabs=False)
        self.update_stats_display()
        self.file_loaded = True

    def _rebuild_panel(self, key, tab, dataset, write_fn, has_equipped, label, use_subtabs=False):
        for child in tab.winfo_children():
            child.destroy()
        panel = InventoryPanel(tab, dataset, write_fn,
                               has_equipped=has_equipped, label=label, use_subtabs=use_subtabs)
        panel.pack(fill="both", expand=True)
        self._panels[key] = panel

    def import_save(self):
        global data
        if not open_file_import():
            return
        if data is None:
            messagebox.showerror("Error", "Please load your current save file first, then click Import.")
            return
        if messagebox.askyesno("Confirm", "This will replace your current character. Continue?"):
            data = data[:0x15F] + import_data[0x15F:]
            if len(data) != 0x9001B0:
                messagebox.showerror("Import Error", "Size mismatch after import.")
                return

            player_items()
            player_usables()
            player_storage()

            self._rebuild_panel("equipment", self._tab_equipment,
                                items,   write_items_to_data,   True,  "Equipment",    use_subtabs=True)
            self._rebuild_panel("usables",  self._tab_usables,
                                usables, write_usables_to_data, False, "Consumable",   use_subtabs=False)
            self._rebuild_panel("storage",  self._tab_storage,
                                storage, write_storage_to_data, False, "Storage Item", use_subtabs=False)
            self.update_stats_display()

            messagebox.showinfo("Success", "File imported. Load in-game to apply the character.")

    def import_inventory(self):
        """Open the Import Inventory dialog."""
        if data is None:
            messagebox.showerror(
                "Error",
                "Please open your current save file first (File → Open Save File)."
            )
            return
        ImportInventoryDialog(self.root, self)

    def create_stats_tab(self):
        self.tab_stats = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_stats, text="Character Stats")

        stats_frame = ttk.LabelFrame(self.tab_stats, text="Character Stats", padding=10)
        stats_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.stat_entries = {}
        stats = [
            ("Amrita",       AMRITA_OFFSET, 8),
            ("Gold",         GOLD_OFFSET,   8),
            ("Constitution", CONSTITUTION,  2),
            ("Heart",        HEART,         2),
            ("Stamina",      STAMINA,       2),
            ("Strength",     STRENGTH,      2),
            ("Skill",        SKILL,         2),
            ("Intellect",    INTELLECT,     2),
            ("Magic",        MAGIC,         2),
        ]
        half = (len(stats) + 1) // 2
        for i, (name, offset, size) in enumerate(stats[:half]):
            ttk.Label(stats_frame, text=name).grid(row=i, column=0, sticky="w", padx=10, pady=5)
            e = ttk.Entry(stats_frame, width=20)
            e.grid(row=i, column=1, sticky="w", padx=10, pady=5)
            self.stat_entries[name] = (e, offset, size)
        for i, (name, offset, size) in enumerate(stats[half:]):
            ttk.Label(stats_frame, text=name).grid(row=i, column=2, sticky="w", padx=20, pady=5)
            e = ttk.Entry(stats_frame, width=20)
            e.grid(row=i, column=3, sticky="w", padx=10, pady=5)
            self.stat_entries[name] = (e, offset, size)

        bf = ttk.Frame(self.tab_stats)
        bf.pack(pady=10)
        ttk.Button(bf, text="Load Stats", command=self.update_stats_display).pack(side="left", padx=20)
        ttk.Button(bf, text="Save Stats", command=self.save_stats).pack(side="left", padx=20)

    def update_stats_display(self):
        if data is None:
            return
        for name, (entry, offset, size) in self.stat_entries.items():
            value = find_value_at_offset(data, offset, size)
            entry.delete(0, tk.END)
            entry.insert(0, value if value is not None else 0)

    def save_stats(self):
        if data is None:
            messagebox.showwarning("Warning", "No save file loaded")
            return
        for name, (entry, offset, size) in self.stat_entries.items():
            try:
                data[offset:offset + size] = write_le(int(entry.get()), size)
            except ValueError:
                messagebox.showerror("Error", "Invalid value for {}".format(name))
                return
        messagebox.showinfo("Success", "Stats updated in memory!")


if __name__ == "__main__":
    root = tk.Tk()
    app  = Nioh3Editor(root)
    root.mainloop()
