import tkinter as tk
from tkinter import ttk
import winreg

def find_autocad_versions():
    """Find installed AutoCAD versions from the registry."""
    base_key_path = r"Software\Autodesk\AutoCAD"
    versions = []

    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, base_key_path) as base_key:
            i = 0
            while True:
                try:
                    version = winreg.EnumKey(base_key, i)
                    versions.append(version)
                    i += 1
                except OSError:
                    break
    except FileNotFoundError:
        pass

    return versions

def find_product_lang_codes(version):
    """Find product and language codes for a given AutoCAD version."""
    base_key_path = fr"Software\Autodesk\AutoCAD\{version}"
    product_lang_codes = []

    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, base_key_path) as base_key:
            i = 0
            while True:
                try:
                    product_lang = winreg.EnumKey(base_key, i)
                    product_lang_codes.append(product_lang)
                    i += 1
                except OSError:
                    break
    except FileNotFoundError:
        pass

    return product_lang_codes

def on_version_select(event):
    """Handle version selection and update product/lang codes."""
    selected_version = version_var.get()
    product_lang_codes = find_product_lang_codes(selected_version)
    product_lang_var.set("")
    product_lang_combo['values'] = product_lang_codes
    state_var.set("")
    toggle_button['state'] = 'disabled'

def on_product_lang_select(event):
    """Handle product/lang code selection and update InfoCenter state."""
    version = version_var.get()
    product_lang = product_lang_var.get()
    reg_path = fr"Software\Autodesk\AutoCAD\{version}\{product_lang}\InfoCenter"

    try:
        reg_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_READ)
        current_state, _ = winreg.QueryValueEx(reg_key, "InfoCenterOn")
        reg_key.Close()
        state_var.set(f"Current InfoCenterOn state: {current_state}")
        update_toggle_button(current_state)
    except FileNotFoundError:
        state_var.set("InfoCenterOn state: Not found")
        toggle_button['state'] = 'disabled'
    except Exception as e:
        state_var.set(f"Error: {e}")
        toggle_button['state'] = 'disabled'

def update_toggle_button(current_state):
    """Update the toggle button based on the current state of InfoCenterOn."""
    if current_state == 0:
        toggle_button.config(text="Enable InfoCenter", command=enable_infocenter, state='normal')
    else:
        toggle_button.config(text="Disable InfoCenter", command=disable_infocenter, state='normal')

def enable_infocenter():
    """Enable InfoCenter and update the GUI."""
    modify_registry(1)
    state_var.set("Successfully enabled InfoCenter.")
    update_toggle_button(1)

def disable_infocenter():
    """Disable InfoCenter and update the GUI."""
    modify_registry(0)
    state_var.set("Successfully disabled InfoCenter.")
    update_toggle_button(0)

def modify_registry(new_state):
    """Modify the InfoCenterOn registry value."""
    version = version_var.get()
    product_lang = product_lang_var.get()
    reg_path = fr"Software\Autodesk\AutoCAD\{version}\{product_lang}\InfoCenter"

    try:
        reg_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(reg_key, "InfoCenterOn", 0, winreg.REG_DWORD, new_state)
        reg_key.Close()
    except FileNotFoundError:
        result_var.set("Registry path not found.")
    except PermissionError:
        result_var.set("Permission denied. Try running as administrator.")
    except Exception as e:
        result_var.set(f"Error: {e}")

# Create the main application window
app = tk.Tk()
app.title("AutoCAD InfoCenter Toggle")

# Create and set the message text variables
result_var = tk.StringVar()
result_var.set("")

state_var = tk.StringVar()
state_var.set("")

# Create the GUI components
version_label = ttk.Label(app, text="AutoCAD Version:")
version_label.grid(column=0, row=0, padx=10, pady=5, sticky='W')

version_var = tk.StringVar()
version_combo = ttk.Combobox(app, textvariable=version_var)
version_combo.grid(column=1, row=0, padx=10, pady=5)
version_combo.bind("<<ComboboxSelected>>", on_version_select)

product_lang_label = ttk.Label(app, text="Product and Language Code:")
product_lang_label.grid(column=0, row=1, padx=10, pady=5, sticky='W')

product_lang_var = tk.StringVar()
product_lang_combo = ttk.Combobox(app, textvariable=product_lang_var)
product_lang_combo.grid(column=1, row=1, padx=10, pady=5)
product_lang_combo.bind("<<ComboboxSelected>>", on_product_lang_select)

toggle_button = ttk.Button(app, text="Enable/Disable InfoCenter", state='disabled')
toggle_button.grid(column=0, row=2, columnspan=2, padx=10, pady=10)

state_label = ttk.Label(app, textvariable=state_var)
state_label.grid(column=0, row=3, columnspan=2, padx=10, pady=5)

result_label = ttk.Label(app, textvariable=result_var, foreground="red")
result_label.grid(column=0, row=4, columnspan=2, padx=10, pady=5)

# Populate AutoCAD versions
versions = find_autocad_versions()
version_combo['values'] = versions

# Run the application
app.mainloop()
