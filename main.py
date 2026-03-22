import subprocess
from tkinter import Tk, filedialog, ttk, messagebox
import os
from pathlib import Path
import json
import tkinter as tk
from checksum import patch_checksum
from openpyxl import Workbook
import shutil

AMRITA_OFFSET  = 0x3ADE25
GOLD_OFFSET    = 0x3ADE35
CONSTITUTION   = 0x3ADF63 - 0x24
HEART          = 0x3ADF67 - 0x24
STAMINA        = 0x3ADF6B - 0x24
STRENGTH       = 0x3ADF6F - 0x24
SKILL          = 0x3ADF73 - 0x24
INTELLECT      = 0x3ADF77 - 0x24
MAGIC          = 0x3ADF7B - 0x24

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

# Merge: manual overrides take priority over auto-mapped entries
_ITEM_TYPE_TO_EFFECT_TYPE = {**_auto_map, **_MANUAL_OVERRIDES}

# Debug: print all item types that could NOT be mapped (will use full effect list as fallback)
_unmapped = sorted(
    itype for itype in set(v.get("type", "") for v in items_json.values())
    if itype and itype not in _ITEM_TYPE_TO_EFFECT_TYPE
)
if _unmapped:
    print("[effect filter] Unmapped item types (will show full effect list):", _unmapped)

_mapped = sorted(set(_ITEM_TYPE_TO_EFFECT_TYPE.items()))
print("[effect filter] Active mappings:", _mapped)


def _get_effect_list_for_item(item):
    """Return the filtered effect list for this item, or the full list as fallback."""
    hex_id   = swap_endian_hex(item.get("item_id", 0))
    _, itype = lookup_item(hex_id)
    etype    = _ITEM_TYPE_TO_EFFECT_TYPE.get(itype)
    if etype and etype in _effects_by_type:
        return _effects_by_type[etype]
    # Fallback — also print so you can see which type is missing a mapping
    if itype and itype != "?":
        print("[effect filter] No mapping for item type '{}' (itype={}, etype={})".format(
            itype, itype, etype))
    return effect_dropdown_list


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
        # Copy the file to backup

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
    item["equipment_flags"]       = data[base + 0x18]
    item["inv_index"]             = read_u16(data, base + 0x1C)
    item["ui_1"]                  = data[base + 0x28]
    item["ui_2"]                  = data[base + 0x29]
    item["rarity"]                = data[base + 0x30]
    effects = []
    se_base = base + 0x38
    for i in range(7):
        eo = se_base + i * 0x08
        effects.append({"effect_id": read_u16(data, eo), "effect_value": read_u16(data, eo + 0x04)})
    item["effects"]         = effects
    item["equipped_flag_1"] = data[base + 0xE8]
    item["equipped_flag_2"] = data[base + 0xEC]
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
    item["equipment_flags"]       = data[base + 0x18]
    item["inv_index"]             = read_u16(data, base + 0x1C)
    item["ui_1"]                  = data[base + 0x28]
    item["ui_2"]                  = data[base + 0x29]
    item["rarity"]                = data[base + 0x30]
    effects = []
    se_base = base + 0x38
    for i in range(7):
        eo = se_base + i * 0x08
        effects.append({"effect_id": read_u16(data, eo), "effect_value": read_u16(data, eo + 0x04)})
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
    data[base+0x18]            = item["equipment_flags"] & 0xFF
    data[base+0x1C:base+0x1E] = write_le(item["inv_index"],            2)
    data[base+0x30]            = item["rarity"]          & 0xFF
    se_base = base + 0x38
    for i, eff in enumerate(item["effects"]):
        eo = se_base + i * 0x08
        data[eo:eo+2]   = write_le(eff["effect_id"],    2)
        data[eo+4:eo+6] = write_le(eff["effect_value"], 2)
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
    # items_json keys are big-endian display hex (e.g. "A12B").
    # Save file stores little-endian (bytes swapped).
    raw     = int(id_hex_str, 16)
    swapped = ((raw & 0xFF) << 8) | (raw >> 8)
    return swapped.to_bytes(2, "little")


def _find_empty_slot_raw(start_offset, slot_size, slot_count):
    """
    Scan `data` directly and return the index of the first slot whose
    first 4 bytes are all 0x00 (the actual empty-slot marker in the save).
    Returns None if no empty slot is found.
    """
    for i in range(slot_count):
        base = start_offset + i * slot_size
        if data[base: base + 4] == b'\x00\x00\x00\x00':
            return i
    return None


def _collect_used_inv_indices():
    """
    Scan ALL inventory sections in raw data and return a set of every
    inv_index value (u16 at offset 0x1C of each slot) that is already in use.
    Only reads slots whose first 4 bytes are non-zero (i.e. occupied slots).
    """
    used = set()
    for start, size, count in [
        (ITEM_START,    ITEM_SIZE,    ITEM_SLOTS),
        (USABLE_START,  USABLE_SIZE,  USABLE_SLOTS),
        (STORAGE_START, STORAGE_SIZE, STORAGE_SLOTS),
    ]:
        for i in range(count):
            base = start + i * size
            if data[base: base + 4] == b'\x00\x00\x00\x00':
                continue   # empty slot — skip
            idx = read_u16(data, base + 0x1C)
            if idx != 0:
                used.add(idx)
    return used


def _generate_unique_inv_index():
    """
    Generate a unique 2-byte inventory index that:
      - Is not already used by any occupied slot across all sections.
      - Has no zero byte (both low byte and high byte are >= 0x01).
    Raises RuntimeError if the entire valid space is exhausted (unlikely).
    """
    used = _collect_used_inv_indices()
    # Iterate over all u16 values where neither byte is 0x00:
    # low byte in 0x01..0xFF, high byte in 0x01..0xFF  -> 255*255 = 65025 candidates
    import random
    candidates = [
        (hi << 8) | lo
        for hi in range(0x01, 0x100)
        for lo in range(0x01, 0x100)
    ]
    # Shuffle to avoid always picking the same low values
    random.shuffle(candidates)
    for val in candidates:
        if val not in used:
            return val
    raise RuntimeError("No unique inventory index available (all 65025 slots used).")


def spawn_equipment(item_name):
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
    # Write unique inventory index at 0x1C (both bytes must be non-zero)
    inv_idx = _generate_unique_inv_index()
    slot_bytes[0x1C:0x1E] = inv_idx.to_bytes(2, "little")

    base = ITEM_START + slot_idx * ITEM_SIZE
    data[base: base + ITEM_SIZE] = slot_bytes

    # Refresh the in-memory dict so the treeview reflects the change
    new_item         = parse_equipment(base)
    new_item["slot"] = slot_idx
    items[slot_idx]  = new_item
    return slot_idx


def spawn_usable(item_name):
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
    # Write unique inventory index at 0x1C (both bytes must be non-zero)
    inv_idx = _generate_unique_inv_index()
    slot_bytes[0x1C:0x1E] = inv_idx.to_bytes(2, "little")

    base = USABLE_START + slot_idx * USABLE_SIZE
    data[base: base + USABLE_SIZE] = slot_bytes

    new_item           = parse_usable(base)
    new_item["slot"]   = slot_idx
    usables[slot_idx]  = new_item
    return slot_idx


_effect_name_by_id = {e["id"]: e["Effect"] for e in effects_json}


def _resolve_effect_name(raw_effect_id):
    """Convert a raw effect_id (read as LE u16 from save) to its big-endian display name."""
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

    # Header: base columns + 7 effect pairs
    header = ["Slot", "Item ID", "Name", "Type", "Quantity", "Level", "Rarity"]
    for n in range(1, 8):
        header += ["Effect {} (ID - Name)".format(n), "Effect {} Value".format(n)]
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
            else:
                display_id = "{:04X}".format(swap_u16(raw_id))
                effect_name = _effect_name_by_id.get(display_id, display_id)
                row.append(f"{display_id} - {effect_name}")
            row.append(eff["effect_value"] if raw_id != 0 else "")

        ws.append(row)

    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        initialfile="{}.xlsx".format(sheet_name),
        filetypes=[("Excel Files", "*.xlsx")]
    )
    if save_path:
        wb.save(save_path)
        messagebox.showinfo("Success", "{} exported successfully!".format(sheet_name))

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
    _EQUIP_PROPS = [
        ("item_id",              "Item ID"),
        ("appearance_id",        "Appearance ID"),
        ("quantity",             "Quantity"),
        ("item_level",           "Item Level"),
        ("item_level_pre_forge", "Item Level (Pre-Forge)"),
        ("plus_value",           "Plus Value"),
        ("familiarity_raw",      "Familiarity"),
        ("equipment_flags",      "Equipment Flags"),
        ("rarity",               "Rarity"),
        ("equipped_flag_1",      "Equipped Flag 1"),
        ("equipped_flag_2",      "Equipped Flag 2"),
    ]
    _USABLE_PROPS = [
        ("item_id",              "Item ID"),
        ("appearance_id",        "Appearance ID"),
        ("quantity",             "Quantity"),
        ("item_level",           "Item Level"),
        ("item_level_pre_forge", "Item Level (Pre-Forge)"),
        ("plus_value",           "Plus Value"),
        ("familiarity_raw",      "Familiarity"),
        ("equipment_flags",      "Flags"),
        ("rarity",               "Rarity"),
    ]

    def __init__(self, master, has_equipped, on_apply, **kwargs):
        super().__init__(master, **kwargs)
        self._on_apply = on_apply
        self._props    = self._EQUIP_PROPS if has_equipped else self._USABLE_PROPS
        self._build()

    def _build(self):
        canvas = tk.Canvas(self)
        sb     = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        frame  = ttk.Frame(canvas)
        frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=frame, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self._entries = {}
        pf = ttk.LabelFrame(frame, text="Properties", padding=8)
        pf.pack(fill="x", padx=5, pady=5)
        for i, (key, lbl) in enumerate(self._props):
            ttk.Label(pf, text=lbl).grid(row=i, column=0, sticky="w", padx=4, pady=2)
            e = ttk.Entry(pf, width=18)
            e.grid(row=i, column=1, sticky="w", padx=4, pady=2)
            self._entries[key] = e

        ef = ttk.LabelFrame(frame, text="Effects", padding=8)
        ef.pack(fill="x", padx=5, pady=5)
        self._eff_combos = []
        self._eff_mags   = []
        for i in range(7):
            ttk.Label(ef, text="Effect {}:".format(i + 1)).grid(row=i, column=0, sticky="w", padx=4, pady=2)
            combo = SearchableCombobox(ef, width=34, values=effect_dropdown_list)
            combo.grid(row=i, column=1, sticky="w", padx=4, pady=2)
            self._eff_combos.append(combo)
            ttk.Label(ef, text="Val:").grid(row=i, column=2, sticky="w", padx=2, pady=2)
            mag = ttk.Entry(ef, width=8)
            mag.grid(row=i, column=3, sticky="w", padx=2, pady=2)
            self._eff_mags.append(mag)

        bf = ttk.Frame(frame)
        bf.pack(pady=8)
        ttk.Button(bf, text="Apply Changes",
                   command=lambda: self._on_apply(self._entries, self._eff_combos, self._eff_mags)
                   ).pack()

    def load(self, item):
        for key, entry in self._entries.items():
            entry.delete(0, tk.END)
            entry.insert(0, item.get(key, 0))

        # Filter effect dropdowns to only show effects compatible with this item type.
        # _get_effect_list_for_item returns the full list if no mapping is found.
        eff_list = _get_effect_list_for_item(item)
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
        self._allowed_types = allowed_types   # None = show everything
        self.title("Spawn {}".format(title_label))
        self.geometry("520x480")
        self.resizable(False, True)
        self.grab_set()

        # Build name list from items_json filtered to only the types
        # that actually exist in the calling panel.  No hardcoded guesses.
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

        cols = ("slot", "item_id", "name", "type", "quantity", "level", "rarity")
        tree = ttk.Treeview(parent, columns=cols, show="headings", height=24)
        vsb  = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        for col in cols:
            tree.heading(col, text=col.capitalize())
        tree.column("slot",     width=45)
        tree.column("item_id",  width=70)
        tree.column("name",     width=220)
        tree.column("type",     width=120)
        tree.column("quantity", width=65)
        tree.column("level",    width=50)
        tree.column("rarity",   width=50)

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
                tree.insert("", "end", iid=item["slot"], values=(
                    item["slot"], hex_id, name, type_,
                    item["quantity"], item["item_level"], item["rarity"]
                ))

        fe.bind("<KeyRelease>", lambda e: populate())
        ttk.Button(fbar, text="Clear",
                   command=lambda: [fvar.set(""), populate()]
                   ).pack(side="left", padx=4)

        # Spawn button — allowed_types is built from what items_json entries
        # actually appear in this panel's dataset, so the dialog list matches
        # exactly what the treeview shows.
        if self._has_equipped:
            # Collect all type strings seen in the equipment dataset
            equip_types = frozenset(
                lookup_item(swap_endian_hex(it["item_id"]))[1]
                for it in self._dataset if it["item_id"] != 0
            )
            ttk.Button(
                fbar, text="Spawn Equipment (WIP) (FOR SOME REASON IT SHOWS 3 DUBLICATE OF THE SAME ITEM INGAME)",
                command=lambda t=equip_types: SpawnDialog(
                    self.winfo_toplevel(), "Equipment", self, allowed_types=t)
            ).pack(side="left", padx=6)
        else:
            # Collect all type strings seen in the usable/storage dataset
            usable_types = frozenset(
                lookup_item(swap_endian_hex(it["item_id"]))[1]
                for it in self._dataset if it["item_id"] != 0
            )
            ttk.Button(
                fbar, text="Spawn Item (WIP) (FOR SOME REASON IT SHOWS 3 DUBLICATE OF THE SAME ITEM INGAME)",
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

    def _apply(self, entries, eff_combos, eff_mags):
        if self._sel_index is None:
            messagebox.showwarning("No Selection", "Select a {} first.".format(self._label))
            return
        item = self._dataset[self._sel_index]
        for key, entry in entries.items():
            try:
                item[key] = int(entry.get())
            except ValueError:
                messagebox.showerror("Error", "Invalid value for '{}'".format(key))
                return
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
        import_menu.add_command(label="Import Save", command=import_save)

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
