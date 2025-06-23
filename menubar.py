import os
import json
import tkinter as tk
from tkinter import simpledialog, messagebox, ttk
from screeninfo import get_monitors
from urllib.parse import urlparse
from PIL import Image, ImageTk
import requests
import webbrowser
import ctypes
import win32com.client
import threading
import pystray
from pystray import MenuItem as item
from PIL import Image as PILImage
import shutil  # For backup/restore file operations

# Ensure config directory exists
CONFIG_PATH = os.path.join(os.getenv("APPDATA"), "CustomTaskbar")
CONFIG_FILE = os.path.join(CONFIG_PATH, "config.json")
BACKUP_FILE = os.path.join(CONFIG_PATH, "config_backup.json")
SHORTCUT_PATH = os.path.join(
    os.getenv("APPDATA"), "Microsoft\\Windows\\Start Menu\\Programs\\Startup\\CustomTaskbar.lnk"
)

if not os.path.exists(CONFIG_PATH):
    os.makedirs(CONFIG_PATH)


class CustomTaskbar(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Custom Taskbar")
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.configure(bg="#1e1e1e")

        # Default properties
        self.links = []
        self.icons = {}
        self.snap_position = "top"
        self.monitor_index = 0
        self.dragging = False

        # New settings for window size and icon size
        self.window_width = None
        self.window_height = None
        self.icon_size = 24  # default icon size in pixels

        # Load settings and build UI
        self.load_settings()
        self.create_widgets()
        self.snap_to_edge()
        self.add_drag_support()
        self.create_tray_icon()
        self.auto_launch_on_startup()

    def load_settings(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r") as f:
                data = json.load(f)
                self.links = data.get("links", [])
                self.snap_position = data.get("snap_position", "top")
                self.monitor_index = data.get("monitor_index", 0)
                # Load new settings with fallbacks
                self.window_width = data.get("window_width")
                self.window_height = data.get("window_height")
                self.icon_size = data.get("icon_size", 24)

    def save_settings(self):
        with open(CONFIG_FILE, "w") as f:
            json.dump(
                {
                    "links": self.links,
                    "snap_position": self.snap_position,
                    "monitor_index": self.monitor_index,
                    "window_width": self.window_width,
                    "window_height": self.window_height,
                    "icon_size": self.icon_size,
                },
                f,
            )

    def create_widgets(self):
        self.canvas = tk.Canvas(self, bg="#2d2d2d", highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.frame = tk.Frame(self.canvas, bg="#2d2d2d")
        self.frame_id = self.canvas.create_window((0, 0), window=self.frame, anchor="nw")

        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side="left", fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", self.adjust_canvas_window)

        self.render_links()
        self.bind("<Button-3>", self.show_menu)  # Right-click menu

    def adjust_canvas_window(self, event):
        self.canvas.itemconfig(self.frame_id, width=event.width)

    def render_links(self):
        # Clear icon cache to reload icons with current icon size
        self.icons.clear()

        for widget in self.frame.winfo_children():
            widget.destroy()

        for i, item in enumerate(self.links):
            icon = self.get_favicon(item["url"])
            btn = tk.Button(
                self.frame,
                text=item["name"],
                image=icon,
                compound=tk.LEFT,
                command=lambda u=item["url"]: webbrowser.open(u),
                bg="#3c3f41",
                fg="white",
                activebackground="#5c5f61",
                relief=tk.FLAT,
                padx=10,
                pady=5,
                anchor="w",
            )
            btn.image = icon
            if self.snap_position == "top":
                btn.grid(row=0, column=i, padx=4, pady=2)
            else:
                btn.grid(row=i, column=0, padx=4, pady=2, sticky="ew")

    def get_favicon(self, url):
        if url in self.icons:
            return self.icons[url]

        icon_url = f"https://www.google.com/s2/favicons?domain={urlparse(url).netloc}&sz=64"
        try:
            response = requests.get(icon_url, stream=True, timeout=3)
            img = Image.open(response.raw).resize((self.icon_size, self.icon_size))
            icon = ImageTk.PhotoImage(img)
            self.icons[url] = icon
            return icon
        except Exception:
            fallback = Image.new("RGB", (self.icon_size, self.icon_size), color="gray")
            icon = ImageTk.PhotoImage(fallback)
            self.icons[url] = icon
            return icon

    def show_menu(self, event):
        menu = tk.Menu(self, tearoff=0, bg="#3c3f41", fg="white")
        menu.add_command(label="Add Link", command=self.add_link)
        menu.add_command(label="Edit Link", command=self.edit_link)
        menu.add_command(label="Delete Link", command=self.delete_link)
        menu.add_separator()
        menu.add_command(label="Snap Top", command=lambda: self.set_snap("top"))
        menu.add_command(label="Snap Left", command=lambda: self.set_snap("left"))
        menu.add_command(label="Snap Right", command=lambda: self.set_snap("right"))
        menu.add_separator()
        menu.add_command(label="Configure Size & Icon", command=self.configure_settings)
        menu.add_command(label="Backup Settings", command=self.backup_settings)
        menu.add_command(label="Restore Settings", command=self.restore_settings)
        menu.add_separator()
        menu.add_command(label="Exit", command=self.quit_app)
        menu.tk_popup(event.x_root, event.y_root)

    def add_link(self):
        url = simpledialog.askstring("Add URL", "Enter URL:")
        if not url:
            return
        name = simpledialog.askstring("Button Name", "Enter name for button:")
        if not name:
            return
        self.links.append({"url": url, "name": name})
        self.render_links()
        self.save_settings()

    def edit_link(self):
        if not self.links:
            messagebox.showinfo("No Links", "No links to edit.")
            return
        choices = [f"{i + 1}. {l['name']}" for i, l in enumerate(self.links)]
        index = simpledialog.askinteger("Edit Link", "Select index:\n" + "\n".join(choices))
        if not index or index < 1 or index > len(self.links):
            return
        link = self.links[index - 1]
        new_url = simpledialog.askstring("New URL", "Enter new URL:", initialvalue=link["url"])
        new_name = simpledialog.askstring("New Name", "Enter new name:", initialvalue=link["name"])
        if new_url:
            link["url"] = new_url
        if new_name:
            link["name"] = new_name
        self.render_links()
        self.save_settings()

    def delete_link(self):
        if not self.links:
            messagebox.showinfo("No Links", "No links to delete.")
            return
        choices = [f"{i + 1}. {l['name']}" for i, l in enumerate(self.links)]
        index = simpledialog.askinteger("Delete Link", "Select index:\n" + "\n".join(choices))
        if not index or index < 1 or index > len(self.links):
            return
        del self.links[index - 1]
        self.render_links()
        self.save_settings()

    def set_snap(self, position):
        self.snap_position = position
        self.snap_to_edge()

    def snap_to_edge(self):
        monitor = get_monitors()[self.monitor_index]

        # Use saved or default sizes
        if self.window_width and self.window_height:
            width = self.window_width
            height = self.window_height
        else:
            if self.snap_position == "top":
                width = int(monitor.width * 0.9)
                height = int(monitor.height * 0.065) + 10
            elif self.snap_position in ("left", "right"):
                width = int(monitor.width * 0.13) + 10
                height = monitor.height

        # Position window accordingly
        if self.snap_position == "top":
            x = monitor.x + int((monitor.width - width) / 2)
            y = monitor.y
            self.geometry(f"{width}x{height}+{x}+{y}")
        elif self.snap_position == "left":
            x = monitor.x
            y = monitor.y
            self.geometry(f"{width}x{height}+{x}+{y}")
        elif self.snap_position == "right":
            x = monitor.x + monitor.width - width
            y = monitor.y
            self.geometry(f"{width}x{height}+{x}+{y}")

        self.window_width = width
        self.window_height = height
        self.render_links()
        self.save_settings()

    def add_drag_support(self):
        self.bind("<B1-Motion>", self.drag_window)
        self.bind("<ButtonRelease-1>", self.detect_snap_position)

    def drag_window(self, event):
        self.dragging = True
        x = self.winfo_pointerx()
        y = self.winfo_pointery()
        self.geometry(f"+{x}+{y}")

    def detect_snap_position(self, event):
        if self.dragging:
            x = self.winfo_x()
            y = self.winfo_y()
            monitor = get_monitors()[self.monitor_index]
            if y <= monitor.y + 50:
                self.set_snap("top")
            elif x <= monitor.x + 50:
                self.set_snap("left")
            elif x >= monitor.x + monitor.width - 100:
                self.set_snap("right")
        self.dragging = False

    def create_tray_icon(self):
        def on_exit():
            self.quit_app()

        icon_image = PILImage.new("RGB", (64, 64), color=(60, 63, 65))
        self.tray_icon = pystray.Icon(
            "CustomTaskbar",
            icon_image,
            menu=pystray.Menu(item("Show", lambda: self.deiconify()), item("Exit", on_exit)),
        )
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def auto_launch_on_startup(self):
        if not os.path.exists(SHORTCUT_PATH):
            target = os.path.abspath(__file__)
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(SHORTCUT_PATH)
            shortcut.TargetPath = target
            shortcut.WorkingDirectory = os.path.dirname(target)
            shortcut.IconLocation = target
            shortcut.save()

    def quit_app(self):
        self.tray_icon.stop()
        self.destroy()

    # New method to configure window and icon sizes with dialogs
    def configure_settings(self):
        width = simpledialog.askinteger(
            "Window Width", "Enter window width (px):", initialvalue=self.window_width or 800, minvalue=100
        )
        if width is None:
            return
        height = simpledialog.askinteger(
            "Window Height", "Enter window height (px):", initialvalue=self.window_height or 60, minvalue=20
        )
        if height is None:
            return
        icon_size = simpledialog.askinteger(
            "Icon Size", "Enter icon size (px):", initialvalue=self.icon_size, minvalue=12, maxvalue=64
        )
        if icon_size is None:
            return

        self.window_width = width
        self.window_height = height
        self.icon_size = icon_size

        self.snap_to_edge()
        self.render_links()
        self.save_settings()

    # Backup current config to backup file
    def backup_settings(self):
        if not os.path.exists(CONFIG_FILE):
            messagebox.showerror("Backup Error", "No configuration file found to backup.")
            return
        try:
            shutil.copy2(CONFIG_FILE, BACKUP_FILE)
            messagebox.showinfo(
                "Backup",
                f"Backup successful.\nBackup file created at:\n{BACKUP_FILE}",
            )
        except Exception as e:
            messagebox.showerror("Backup Error", f"Failed to create backup.\nError: {e}")

    # Restore config from backup file
    def restore_settings(self):
        if not os.path.exists(BACKUP_FILE):
            messagebox.showerror("Restore Error", "No backup file found.")
            return
        if messagebox.askyesno(
            "Confirm Restore", "Are you sure you want to restore from backup? This will overwrite current settings."
        ):
            try:
                shutil.copy2(BACKUP_FILE, CONFIG_FILE)
                self.load_settings()
                self.snap_to_edge()
                self.render_links()
                messagebox.showinfo("Restore", "Settings restored from backup. Restart the app if necessary.")
            except Exception as e:
                messagebox.showerror("Restore Error", f"Failed to restore backup.\nError: {e}")


if __name__ == "__main__":
    app = CustomTaskbar()
    app.mainloop()
