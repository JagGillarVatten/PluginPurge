import os
import sys
import glob
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import winreg
from concurrent.futures import ThreadPoolExecutor
from functools import lru_cache
import threading

# TODO: Consider using a configuration file for paths and other settings
VST_PATHS = [
    r"C:\Program Files\VSTPlugins",
    r"C:\Program Files (x86)\VSTPlugins",
    r"C:\Program Files\Common Files\VST2",
    r"C:\Program Files\Common Files\VST3",
]

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

@lru_cache(maxsize=None)
def get_vst_paths_from_registry() -> list:
    """Retrieve VST paths from Windows registry"""
    paths = []
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\VST") as key:
            i = 0
            while True:
                try:
                    paths.append(winreg.EnumValue(key, i)[1])
                    i += 1
                except WindowsError:
                    break
    except WindowsError:
        pass
    return paths

# TODO: Consider making this a class method to avoid global variables
VST_PATHS.extend(get_vst_paths_from_registry())

class PluginPurgeApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("PluginPurge - Audio Plugin Uninstaller")
        self.plugins = []
        self.current_sort = {"column": "Name", "reverse": False}
        self.search_term = tk.StringVar()
        self.filtered_plugins = []
        self.selected_items = []
        self.setup_ui()
        self.load_plugins()

    def setup_ui(self) -> None:
        """Set up the main user interface"""
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        self.setup_search_bar()
        self.setup_treeview()
        self.setup_buttons()
        self.setup_status_bar()
        self.setup_context_menu()
        self.setup_credits()

    def setup_search_bar(self) -> None:
        """Set up the search bar"""
        search_frame = ttk.Frame(self.main_frame)
        search_frame.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        search_frame.columnconfigure(0, weight=1)
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_term)
        self.search_entry.grid(row=0, column=0, sticky="ew")
        self.search_entry.bind("<KeyRelease>", self.filter_plugins)
        ttk.Button(search_frame, text="Search", command=self.filter_plugins).grid(row=0, column=1, padx=(5, 0))

    def setup_treeview(self) -> None:
        """Set up the treeview for displaying plugins"""
        self.treeview = ttk.Treeview(self.main_frame, columns=("Name", "Vendor", "Version", "Size", "Path"), show='headings')
        for col in self.treeview["columns"]:
            self.treeview.heading(col, text=col, command=lambda _col=col: self.treeview_sort_column(_col, False))
            self.treeview.column(col, anchor="w")
        self.treeview.grid(row=1, column=0, columnspan=2, sticky="nsew")
        self.scrollbar = ttk.Scrollbar(self.main_frame, orient=tk.VERTICAL, command=self.treeview.yview)
        self.treeview.configure(yscroll=self.scrollbar.set)
        self.scrollbar.grid(row=1, column=2, sticky="ns")
        self.main_frame.grid_rowconfigure(1, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)

    def setup_buttons(self) -> None:
        """Set up the action buttons"""
        button_frame = ttk.Frame(self.main_frame)
        button_frame.grid(row=2, column=0, columnspan=3, pady=(10, 0))
        ttk.Button(button_frame, text="Uninstall Selected", command=self.uninstall_selected).grid(row=0, column=0, padx=(0, 5))
        ttk.Button(button_frame, text="Refresh", command=self.refresh_plugins).grid(row=0, column=1, padx=(5, 0))

    def setup_status_bar(self) -> None:
        """Set up the status bar"""
        self.status_var = tk.StringVar()
        ttk.Label(self.main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W).grid(row=3, column=0, columnspan=3, sticky="ew", pady=(10, 0))

    def setup_context_menu(self) -> None:
        """Set up the context menu for right-click actions"""
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Open Folder", command=self.open_plugin_folder)
        self.treeview.bind("<Button-3>", self.show_context_menu)

    def setup_credits(self) -> None:
        """Set up the credits label"""
        credits_label = ttk.Label(self.main_frame, text="Created by JagGillarVatten/Pixelody", font=("Arial", 8, "italic"))
        credits_label.grid(row=4, column=0, columnspan=3, sticky="se", pady=(5, 0))

    def show_context_menu(self, event: tk.Event) -> None:
        """Display the context menu on right-click"""
        item = self.treeview.identify_row(event.y)
        if item:
            self.treeview.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)

    def open_plugin_folder(self) -> None:
        """Open the folder containing the selected plugin"""
        selected_item = self.treeview.selection()
        if selected_item:
            os.startfile(os.path.dirname(self.treeview.item(selected_item[0], 'values')[-1]))

    def load_plugins(self) -> None:
        """Start loading plugins in a separate thread"""
        self.update_status("Loading plugins...")
        threading.Thread(target=self._load_plugins_thread, daemon=True).start()

    def _load_plugins_thread(self) -> None:
        """Load plugins using a thread pool"""
        with ThreadPoolExecutor() as executor:
            self.plugins = [item for sublist in executor.map(self.find_plugins_fast, VST_PATHS) for item in sublist]
        self.root.after(0, self._update_ui_after_load)

    def _update_ui_after_load(self) -> None:
        """Update UI after plugins are loaded"""
        self.display_plugins(self.plugins)
        self.update_status(f"Total plugins: {len(self.plugins)}")

    def find_plugins_fast(self, path: str) -> list:
        """Find plugins in the given path"""
        return [
            p for ext in ['*.dll', '*.vst3', '*.clap'] for p in glob.glob(os.path.join(path, ext))
        ] if os.path.exists(path) else []

    @lru_cache(maxsize=1000)
    def get_plugin_details(self, plugin_path: str) -> dict:
        """Get details of a plugin"""
        name = os.path.basename(plugin_path)
        version = vendor = "Unknown"
        size = os.path.getsize(plugin_path)
        try:
            with open(plugin_path, 'rb') as f:
                content = f.read(4096)
                vendor_index = content.find(b'Vender: ')
                if vendor_index != -1:
                    vendor = content[vendor_index+8:vendor_index+50].split(b'\0', 1)[0].decode('utf-8', errors='ignore')
                version_index = content.find(b'Version: ')
                if version_index != -1:
                    version = content[version_index+9:version_index+50].split(b'\0', 1)[0].decode('utf-8', errors='ignore')
        except Exception:
            pass
        return {'path': plugin_path, 'name': name, 'version': version, 'vendor': vendor, 'size': f"{size / (1024 * 1024):.2f} MB"}

    def display_plugins(self, plugins: list) -> None:
        """Display plugins in the treeview"""
        self.treeview.delete(*self.treeview.get_children())
        for detail in map(self.get_plugin_details, plugins):
            self.treeview.insert('', 'end', values=(
                detail['name'],
                detail['vendor'],
                detail['version'],
                detail['size'],
                detail['path'],
            ))
        self.update_status(f"Displaying {len(plugins)} plugins")

    def uninstall_selected(self) -> None:
        """Uninstall selected plugins"""
        selected_items = self.treeview.selection()
        if not selected_items:
            messagebox.showwarning("Warning", "No plugins selected!")
            return
        if messagebox.askyesno("Confirm Uninstall", f"Are you sure you want to uninstall {len(selected_items)} plugin(s)?"):
            uninstalled_count = sum(1 for item in selected_items if self.uninstall_plugin(self.treeview.item(item, 'values')[-1]))
            self.update_status(f"Uninstalled {uninstalled_count} plugin(s)")
            self.refresh_plugins()

    def uninstall_plugin(self, plugin_path: str) -> bool:
        """Uninstall a single plugin"""
        try:
            os.remove(plugin_path)
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to uninstall {plugin_path}: {e}")
            return False

    def filter_plugins(self, event: tk.Event | None = None) -> None:
        """Filter plugins based on search term"""
        search_term = self.search_term.get().lower()
        self.display_plugins([p for p in self.plugins if search_term in os.path.basename(p).lower()])

    def refresh_plugins(self) -> None:
        """Refresh the plugin list"""
        self.plugins.clear()
        self.get_plugin_details.cache_clear()
        self.load_plugins()

    def update_status(self, message: str) -> None:
        """Update the status bar message"""
        self.status_var.set(message)

    def treeview_sort_column(self, col: str, reverse: bool) -> None:
        """Sort treeview column"""
        l = sorted(((self.treeview.set(k, col), k) for k in self.treeview.get_children('')), reverse=reverse)
        for index, (_, k) in enumerate(l):
            self.treeview.move(k, '', index)
        self.treeview.heading(col, command=lambda: self.treeview_sort_column(col, not reverse))


if __name__ == "__main__":
    root = tk.Tk()
    root.style = ttk.Style()
    root.style.theme_use("clam")
    app = PluginPurgeApp(root)
    root.mainloop()

# TODO: Add error handling for file operations
# TODO: Implement logging for better debugging
# TODO: Consider adding a progress bar for long operations
# TODO: Add option to backup plugins before uninstalling
# TODO: Implement a way to restore uninstalled plugins
# TODO: Add support for additional plugin formats
# TODO: Improve plugin metadata extraction
# TODO: Add unit tests for core functionality
# TODO: Implement a plugin scanning scheduler
# TODO: Add multi-language support
