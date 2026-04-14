# ⚔️ Nioh 3 Save Editor


A powerful and easy-to-use save editor for **Nioh 3**, designed to let you modify your save data safely without needing a hex editor.

---

## 🔥 Features

- Export your inventory to excel
- Import other PC saves
- Item spawn (WIP)
- Import another player inventory

### 🧍 Character & Stats Editing
Modify all core character stats directly:

- Amrita  
- Gold  
- Level  
- Constitution  
- Heart  
- Courage  
- Stamina  
- Strength  
- Skill  
- Dexterity  
- Magic  

---

### 🗡️ Weapon Editing
Fully customize your weapons:

- Edit Weapon IDs  
- Change weapon modifiers / special effects  
- Adjust rarity  
- Modify +Level values  
- Customize weapon slots  

---

### 📦 Item Editing
Complete control over inventory:

- Modify any in-game item  
- Add / remove items  
- Change item quantities  
- Supports `items.json` for fast lookup  
- ID → Name mapping for easy editing  


---

## 🛠️ How to Use

### 💻 PC Version

1. Open the editor  
2. Select your save file  
3. Choose the character slot  
4. Edit anything you want  
5. Click **Save**  

Your save file will be updated automatically.

---

## 🎯 Project Goals

This editor is built to make save editing:

- Safe  
- Simple  
- Fast  
- Accessible  

Everything is handled through a clean UI — no hex editing required.

---

## 📌 Requirements

- Windows 10 / 11  
- Python 3.12+ (if running from source)  
- PyInstaller (for building executable)  

---

## How to Build

```bash
pyinstaller ^
--onefile ^
--noconsole ^
--name "Nioh3_Save_Editor" ^
--add-data "items_little_endian.json;." ^
--add-data "effects_big_endian.json;." ^
--add-data "pc\pc.exe;pc" ^
--add-data "PC_import\pc.exe;PC_import" ^
main.py
```


## ⚠️ Disclaimer

- Always **backup your save file** before editing  
- Use at your own risk  
- Online use may carry risks depending on game protections  

---

## 💬 Issues / Feature Requests

If you encounter a bug or want a new feature:

1. Open an issue  
2. Describe:
   - What you edited  
   - Which slot  
   - What went wrong  
3. Include screenshots if possible  

---

## Credit 

- Data sheet used from Darkksss1 at https://www.nexusmods.com/nioh3/mods/6
- A rework of [[pawREP](https://github.com/pawREP)] decryption tool https://github.com/pawREP/Nioh-Savedata-Decryption-Tool
