import subprocess
from tkinter import Tk, filedialog, ttk, messagebox
import os
from pathlib import Path
import json
import tkinter as tk
from checksum import patch_checksum
# ==================== CONSTANTS ====================
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

# ==================== GLOBALS ====================
data       = None
MODE       = None
decrypted_path = None
items      = []
usables    = []
storage    = []

base_dir = Path(__file__).parent

# ==================== JSON LOADING ====================
def load_json(file_name):
    with open(base_dir / file_name, "r") as f:
        return json.load(f)

items_json   = load_json("items.json")
effects_json = load_json("effects.json")
effect_dropdown_list = [
    "{} - {}".format(e["id"], e["Effect"]) for e in effects_json
]

# ==================== HELPERS ====================
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

def save_file():
    global data, decrypted_path
    if data is None or decrypted_path is None:
        messagebox.showwarning("Warning", "No data to save. Load a file first.")
        return
    
    
    write_items_to_data()
    write_usables_to_data()
    write_storage_to_data()
    data, checksum = patch_checksum(data)
    print(f"Final checksum: {checksum:08X}")
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
    item["unknown_1"]             = data[base + 0x1C]
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
    item["unknown_1"]             = data[base + 0x1C]
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
    data[base:base+2]     = write_le(item["item_id"],              2)
    data[base+2:base+4]   = write_le(item["appearance_id"],        2)
    data[base+4:base+6]   = write_le(item["quantity"],             2)
    data[base+6:base+8]   = write_le(item["item_level"],           2)
    data[base+8:base+10]  = write_le(item["item_level_pre_forge"], 2)
    data[base+10:base+12] = write_le(item["plus_value"],           2)
    data[base+0x14:base+0x18] = write_le(item["familiarity_raw"],  4)
    data[base+0x18] = item["equipment_flags"] & 0xFF
    data[base+0x30] = item["rarity"]          & 0xFF
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
        _write_slot(item, ITEM_START + item["slot"] * ITEM_SIZE, has_equipped=True)

def write_usables_to_data():
    for item in usables:
        _write_slot(item, USABLE_START + item["slot"] * USABLE_SIZE, has_equipped=False)

def write_storage_to_data():
    for item in storage:
        _write_slot(item, STORAGE_START + item["slot"] * STORAGE_SIZE, has_equipped=False)

# ==================== SEARCHABLE COMBOBOX ====================
class SearchableCombobox(ttk.Frame):
    """
    Drop-down combobox with live search filtering.
    Key fix: self._loading flag suppresses _on_type when set() is called
    programmatically, preventing the dropdown from auto-opening.
    """
    def __init__(self, master=None, values=None, width=40, **kwargs):
        super().__init__(master, **kwargs)
        self.full_values     = values if values else []
        self.filtered_values = self.full_values.copy()
        self._loading        = False  # True during programmatic set() -> no dropdown

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
        if self._loading:          # suppress when value is set programmatically
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
        self._loading = True     # prevents _on_type from opening the dropdown
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
    """Scrollable property + effect editor, shared by all InventoryPanel instances."""

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
        """Populate all fields. Uses SearchableCombobox._loading guard -> no dropdown."""
        for key, entry in self._entries.items():
            entry.delete(0, tk.END)
            entry.insert(0, item.get(key, 0))
        for i, eff in enumerate(item.get("effects", [])):
            hex_id  = "{:04X}".format(eff["effect_id"])
            matched = next((c for c in effect_dropdown_list if c.startswith(hex_id)), "")
            self._eff_combos[i].set(matched)   # _loading=True inside set() -> no dropdown
            self._eff_mags[i].delete(0, tk.END)
            self._eff_mags[i].insert(0, eff["effect_value"])

    def clear(self):
        for e in self._entries.values():
            e.delete(0, tk.END)
        for combo in self._eff_combos:
            combo.set("")
        for mag in self._eff_mags:
            mag.delete(0, tk.END)

# ==================== INVENTORY PANEL ====================
class InventoryPanel(ttk.Frame):
    """
    Split-pane panel: left = filterable treeview, right = ItemEditor.
    For equipment the left pane uses sub-tabs grouped by item type.
    Clicking a row immediately populates the editor on the right.
    """

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
            tree = self._build_tree_pane(left, type_filter=None)
            self._active_tree = tree

    # ── Sub-tab layout (equipment by type) ───────────────────────────────────

    def _build_subtab_layout(self, parent):
        type_order = []
        seen = set()
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
            tree = self._build_tree_pane(frame, type_filter=tab_type)
            self._sub_trees[tab_type] = tree

        self._active_tree = self._sub_trees["All"]
        nb.bind("<<NotebookTabChanged>>",
                lambda e, n=nb: self._on_subtab_change(n))

    def _on_subtab_change(self, nb):
        tab_text = nb.tab(nb.select(), "text")
        self._active_tree = self._sub_trees.get(tab_text, self._active_tree)

    # ── Tree pane builder ─────────────────────────────────────────────────────

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

        tree._populate = populate
        tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        populate()

        bbar = ttk.Frame(parent)
        bbar.pack(fill="x", padx=5, pady=4)
        ttk.Button(bbar, text="Delete",       command=self._delete).pack(side="left", padx=4)
        ttk.Button(bbar, text="Max Quantity", command=self._max_all).pack(side="left", padx=4)

        return tree

    # ── Events & actions ─────────────────────────────────────────────────────

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

    def _delete(self):
        if self._sel_index is None:
            messagebox.showwarning("No Selection", "Select a {} first.".format(self._label))
            return
        if messagebox.askyesno("Confirm", "Delete this {}?".format(self._label)):
            self._dataset[self._sel_index]["item_id"]  = 0
            self._dataset[self._sel_index]["quantity"] = 0
            self._sel_index = None
            self._editor.clear()
            self.refresh()

    def _max_all(self):
        if messagebox.askyesno("Confirm", "Set all {} quantities to 9999?".format(self._label)):
            count = 0
            for item in self._dataset:
                if item["item_id"] != 0:
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
            if chosen:
                try:
                    item["effects"][i]["effect_id"] = int(chosen.split(" ")[0], 16)
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
                            items,   write_items_to_data,   True,  "Equipment",    use_subtabs=True)
        self._rebuild_panel("usables",  self._tab_usables,
                            usables, write_usables_to_data, False, "Consumable",   use_subtabs=False)
        self._rebuild_panel("storage",  self._tab_storage,
                            storage, write_storage_to_data, False, "Storage Item", use_subtabs=False)

        self.update_stats_display()
        self.file_loaded = True

        # eq_cnt = sum(1 for i in items   if i["item_id"] != 0)
        # us_cnt = sum(1 for i in usables if i["item_id"] != 0)
        # st_cnt = sum(1 for i in storage if i["item_id"] != 0)
        # # messagebox.showinfo("Loaded",
        # #     "Mode: {}\nEquipment: {}\nConsumables: {}\nStorage: {}".format(
        # #         MODE, eq_cnt, us_cnt, st_cnt))

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


# ==================== MAIN ====================
if __name__ == "__main__":
    root = tk.Tk()
    app  = Nioh3Editor(root)
    root.mainloop()