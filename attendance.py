import cv2
from pyzbar.pyzbar import decode
import pandas as pd
import datetime
import os
import time
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk
import threading
from openpyxl import load_workbook
import re


# ─────────────────────────────────────────────
#  Color / Style constants
# ─────────────────────────────────────────────
BG          = "#F8F7F4"
SIDEBAR_BG  = "#FFFFFF"
HEADER_BG   = "#FFFFFF"
ACCENT      = "#1D9E75"
ACCENT_DIM  = "#0F6E56"
WARN        = "#E24B4A"
TEXT_PRI    = "#1A1A1A"
TEXT_SEC    = "#6B6B68"
BORDER      = "#E5E4DE"
ROW_ALT     = "#F3F2EE"
ROW_HOVER   = "#EAF3EE"
BADGE_COLORS = {
    "CSE": ("#E6F1FB", "#185FA5"),
    "EEE": ("#EEEDFE", "#534AB7"),
    "ME":  ("#FAEEDA", "#854F0B"),
    "BBA": ("#FBEAF0", "#993556"),
    "CE":  ("#EAF3DE", "#3B6D11"),
}

FONT_HEAD  = ("Segoe UI", 13, "bold")
FONT_BODY  = ("Segoe UI", 11)
FONT_SMALL = ("Segoe UI", 9)
FONT_MONO  = ("Consolas", 10)
FONT_BIG   = ("Segoe UI", 22, "bold")


# ─────────────────────────────────────────────
#  Main Application
# ─────────────────────────────────────────────
class QRAttendanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("QR Attendance System")
        self.root.geometry("1100x680")
        self.root.minsize(900, 580)
        self.root.configure(bg=BG)

        # State
        self.scanning       = False
        self.cap            = None
        self.cam_thread     = None
        self.scanned_ids    = set()
        self.last_scan_time = {}
        self.records        = []
        self.session_start  = None
        self.session_running = False
        self.file_name      = None
        self._photo         = None   # prevent GC

        self._build_ui()
        self._new_session()

    # ──────────────────────────────────────────
    #  UI construction
    # ──────────────────────────────────────────
    def _build_ui(self):
        # ── Top bar ──────────────────────────
        topbar = tk.Frame(self.root, bg=HEADER_BG, height=52)
        topbar.pack(fill="x", side="top")
        topbar.pack_propagate(False)

        tk.Frame(topbar, bg=ACCENT, width=4).pack(side="left", fill="y")

        tk.Label(topbar, text="QR Attendance System",
                 font=("Segoe UI", 13, "bold"),
                 bg=HEADER_BG, fg=TEXT_PRI).pack(side="left", padx=16, pady=12)

        self.lbl_session = tk.Label(topbar, text="",
                                    font=FONT_SMALL, bg=HEADER_BG, fg=TEXT_SEC)
        self.lbl_session.pack(side="left", padx=4)

        # Status indicator (right)
        self.status_canvas = tk.Canvas(topbar, width=10, height=10,
                                       bg=HEADER_BG, highlightthickness=0)
        self.status_canvas.pack(side="right", padx=(0, 8))
        self.status_dot = self.status_canvas.create_oval(1, 1, 9, 9, fill=TEXT_SEC, outline="")

        self.lbl_status = tk.Label(topbar, text="Camera off",
                                   font=FONT_SMALL, bg=HEADER_BG, fg=TEXT_SEC)
        self.lbl_status.pack(side="right", padx=(0, 6))

        # separator
        tk.Frame(self.root, bg=BORDER, height=1).pack(fill="x")

        # ── Body (sidebar + main) ─────────────
        body = tk.Frame(self.root, bg=BG)
        body.pack(fill="both", expand=True)

        # Sidebar
        sidebar = tk.Frame(body, bg=SIDEBAR_BG, width=270)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)
        tk.Frame(body, bg=BORDER, width=1).pack(side="left", fill="y")

        self._build_sidebar(sidebar)

        # Main panel
        main = tk.Frame(body, bg=BG)
        main.pack(side="left", fill="both", expand=True)
        self._build_main(main)

    def _build_sidebar(self, parent):
        # Camera preview
        cam_outer = tk.Frame(parent, bg=BORDER, bd=0)
        cam_outer.pack(fill="x", padx=14, pady=(14, 0))

        self.cam_label = tk.Label(cam_outer, bg="#E8E7E2",
                                  width=240, height=180)
        self.cam_label.pack()

        # Placeholder text inside camera
        self.cam_placeholder = tk.Label(
            self.cam_label, text="📷  Camera preview",
            font=FONT_SMALL, bg="#E8E7E2", fg=TEXT_SEC)
        self.cam_placeholder.place(relx=0.5, rely=0.5, anchor="center")

        # Start / Stop button
        self.btn_scan = tk.Button(
            parent, text="▶  Start Scanner",
            font=FONT_BODY, bg=ACCENT, fg="white",
            activebackground=ACCENT_DIM, activeforeground="white",
            relief="flat", cursor="hand2", bd=0, pady=8,
            command=self.toggle_scanner)
        self.btn_scan.pack(fill="x", padx=14, pady=10)

        # Stats
        stats_frame = tk.Frame(parent, bg=SIDEBAR_BG)
        stats_frame.pack(fill="x", padx=14, pady=(0, 6))

        left_stat = tk.Frame(stats_frame, bg="#F3F2EE",
                             relief="flat", bd=0)
        left_stat.pack(side="left", fill="both", expand=True,
                       padx=(0, 5), ipady=8, ipadx=6)
        tk.Label(left_stat, text="Total scanned",
                 font=FONT_SMALL, bg="#F3F2EE", fg=TEXT_SEC).pack()
        self.lbl_total = tk.Label(left_stat, text="0",
                                  font=FONT_BIG, bg="#F3F2EE", fg=TEXT_PRI)
        self.lbl_total.pack()

        right_stat = tk.Frame(stats_frame, bg="#F3F2EE",
                              relief="flat", bd=0)
        right_stat.pack(side="left", fill="both", expand=True,
                        padx=(5, 0), ipady=8, ipadx=6)
        tk.Label(right_stat, text="Unique students",
                 font=FONT_SMALL, bg="#F3F2EE", fg=TEXT_SEC).pack()
        self.lbl_unique = tk.Label(right_stat, text="0",
                                   font=FONT_BIG, bg="#F3F2EE", fg=TEXT_PRI)
        self.lbl_unique.pack()

        # Session timer
        timer_frame = tk.Frame(parent, bg="#F3F2EE")
        timer_frame.pack(fill="x", padx=14, pady=(0, 6), ipady=6, ipadx=6)
        tk.Label(timer_frame, text="Session duration",
                 font=FONT_SMALL, bg="#F3F2EE", fg=TEXT_SEC).pack(side="left", padx=8)
        self.lbl_timer = tk.Label(timer_frame, text="00:00",
                                  font=FONT_MONO, bg="#F3F2EE", fg=TEXT_PRI)
        self.lbl_timer.pack(side="right", padx=8)

        # Last scan card
        self.last_card = tk.Frame(parent, bg="#EAF3EE", bd=0)
        self.last_card.pack(fill="x", padx=14, pady=(0, 10), ipady=8, ipadx=8)
        tk.Label(self.last_card, text="Last scanned",
                 font=FONT_SMALL, bg="#EAF3EE", fg=ACCENT_DIM).pack(anchor="w", padx=8)
        self.lbl_last_name = tk.Label(self.last_card, text="—",
                                      font=("Segoe UI", 11, "bold"),
                                      bg="#EAF3EE", fg=TEXT_PRI)
        self.lbl_last_name.pack(anchor="w", padx=8)
        self.lbl_last_meta = tk.Label(self.last_card, text="—",
                                      font=FONT_SMALL, bg="#EAF3EE", fg=TEXT_SEC)
        self.lbl_last_meta.pack(anchor="w", padx=8)

        # New session / Export buttons
        btn_frame = tk.Frame(parent, bg=SIDEBAR_BG)
        btn_frame.pack(fill="x", padx=14, pady=(0, 14))
        tk.Button(btn_frame, text="New Session",
                  font=FONT_SMALL, bg=SIDEBAR_BG, fg=TEXT_SEC,
                  relief="flat", cursor="hand2", bd=1,
                  highlightbackground=BORDER,
                  command=self._new_session).pack(side="left", fill="x",
                                                  expand=True, padx=(0, 4))
        tk.Button(btn_frame, text="Export Excel",
                  font=FONT_SMALL, bg=SIDEBAR_BG, fg=ACCENT,
                  relief="flat", cursor="hand2", bd=1,
                  highlightbackground=BORDER,
                  command=self.export_excel).pack(side="left", fill="x",
                                                  expand=True, padx=(4, 0))

    def _build_main(self, parent):
        # Toolbar
        toolbar = tk.Frame(parent, bg=HEADER_BG, height=46)
        toolbar.pack(fill="x")
        toolbar.pack_propagate(False)
        tk.Frame(parent, bg=BORDER, height=1).pack(fill="x")

        tk.Label(toolbar, text="🔍", font=("Segoe UI", 11),
                 bg=HEADER_BG, fg=TEXT_SEC).pack(side="left", padx=(12, 2))
        self.search_var = tk.StringVar()
        search_entry = tk.Entry(toolbar, textvariable=self.search_var,
                                font=FONT_BODY, relief="flat",
                                bg="#F3F2EE", fg=TEXT_PRI,
                                insertbackground=TEXT_PRI, width=20)
        search_entry.pack(side="left", padx=(0, 12), ipady=4)
        search_entry.insert(0, "Search name or ID…")
        search_entry.bind("<FocusIn>",
            lambda e: search_entry.delete(0, "end")
            if search_entry.get() == "Search name or ID…" else None)

        tk.Label(toolbar, text="Dept:", font=FONT_SMALL,
                 bg=HEADER_BG, fg=TEXT_SEC).pack(side="left")
        self.dept_var = tk.StringVar(value="All")
        dept_cb = ttk.Combobox(toolbar, textvariable=self.dept_var,
                               values=["All", "CSE", "EEE", "ME", "BBA", "CE"],
                               state="readonly", width=8, font=FONT_SMALL)
        dept_cb.pack(side="left", padx=6)
        dept_cb.bind("<<ComboboxSelected>>", lambda _: self._refresh_table())
        self.search_var.trace_add("write", lambda *_: self._refresh_table())

        tk.Button(toolbar, text="Export CSV",
                  font=FONT_SMALL, bg=HEADER_BG, fg=TEXT_SEC,
                  relief="flat", cursor="hand2",
                  command=self.export_csv).pack(side="right", padx=12)

        # Table
        table_frame = tk.Frame(parent, bg=BG)
        table_frame.pack(fill="both", expand=True, padx=0, pady=0)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Custom.Treeview",
                        background=HEADER_BG, fieldbackground=HEADER_BG,
                        foreground=TEXT_PRI, font=FONT_BODY,
                        rowheight=32, borderwidth=0)
        style.configure("Custom.Treeview.Heading",
                        background=HEADER_BG, foreground=TEXT_SEC,
                        font=("Segoe UI", 9, "bold"),
                        relief="flat", borderwidth=0)
        style.map("Custom.Treeview",
                  background=[("selected", "#D3EDE4")],
                  foreground=[("selected", TEXT_PRI)])

        cols = ("#", "ID", "Name", "Department", "Semester", "Date", "Time")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings",
                                 style="Custom.Treeview")

        widths = [40, 90, 160, 100, 80, 100, 80]
        for col, w in zip(cols, widths):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, minwidth=w, anchor="w")

        self.tree.tag_configure("odd",  background=HEADER_BG)
        self.tree.tag_configure("even", background=ROW_ALT)

        vsb = ttk.Scrollbar(table_frame, orient="vertical",
                            command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)

        # Status bar
        self.statusbar = tk.Label(parent, text="Ready",
                                  font=FONT_SMALL, bg=BORDER, fg=TEXT_SEC,
                                  anchor="w", pady=3)
        self.statusbar.pack(fill="x", side="bottom")

    # ──────────────────────────────────────────
    #  Session management
    # ──────────────────────────────────────────
    def _new_session(self):
        if self.scanning:
            self.toggle_scanner()
        ts = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        self.file_name   = f"attendance_{ts}.xlsx"
        self.scanned_ids = set()
        self.last_scan_time = {}
        self.records     = []
        df = pd.DataFrame(columns=["ID", "Name", "Department", "Semester", "Date", "Time"])
        df.to_excel(self.file_name, index=False)
        self.lbl_session.config(text=f"File: {self.file_name}")
        self._refresh_table()
        self._update_stats()
        self.statusbar.config(text=f"New session started → {self.file_name}")

    # ──────────────────────────────────────────
    #  Scanner control
    # ──────────────────────────────────────────
    def toggle_scanner(self):
        if not self.scanning:
            self._start_scanner()
        else:
            self._stop_scanner()

    def _start_scanner(self):
        self.cap = cv2.VideoCapture(0)
        if not self.cap.isOpened():
            messagebox.showerror("Camera Error",
                                 "Cannot open camera.\nMake sure a webcam is connected.")
            return
        self.scanning = True
        self.session_start = time.time()
        self.session_running = True
        self.btn_scan.config(text="■  Stop Scanner", bg=WARN,
                             activebackground="#A32D2D")
        self._set_status("Scanning…", ACCENT)
        self.cam_placeholder.place_forget()
        self._tick_timer()
        self.cam_thread = threading.Thread(target=self._scan_loop, daemon=True)
        self.cam_thread.start()

    def _stop_scanner(self):
        self.scanning = False
        self.session_running = False
        if self.cap:
            self.cap.release()
            self.cap = None
        self.btn_scan.config(text="▶  Start Scanner", bg=ACCENT,
                             activebackground=ACCENT_DIM)
        self._set_status("Camera off", TEXT_SEC)
        self.cam_label.config(image="")
        self.cam_placeholder.place(relx=0.5, rely=0.5, anchor="center")
        self._photo = None

    def _set_status(self, text, color):
        self.lbl_status.config(text=text)
        self.status_canvas.itemconfig(self.status_dot, fill=color)

    # ──────────────────────────────────────────
    #  Camera / QR loop (runs in thread)
    # ──────────────────────────────────────────
    def _scan_loop(self):
        while self.scanning and self.cap and self.cap.isOpened():
            ret, frame = self.cap.read()
            if not ret:
                break

            # Decode QR codes
            for barcode in decode(frame):
                qr_data = barcode.data.decode("utf-8").strip()
                x, y, w, h = barcode.rect
                cv2.rectangle(frame, (x, y), (x+w, y+h), (29, 158, 117), 2)
                cv2.putText(frame, qr_data[:30], (x, y - 8),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.45, (29, 158, 117), 2)
                self.root.after(0, self._process_qr, qr_data)

            # Resize for preview
            preview = cv2.resize(frame, (240, 180))
            preview = cv2.cvtColor(preview, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(preview)
            photo = ImageTk.PhotoImage(image=img)
            self.root.after(0, self._update_preview, photo)
            time.sleep(0.04)

    def _update_preview(self, photo):
        self._photo = photo
        self.cam_label.config(image=photo)

    # ──────────────────────────────────────────
    #  QR processing
    # ──────────────────────────────────────────
    def _process_qr(self, qr_data):
        try:
            fields = {}
            for part in qr_data.split(";"):
                if ":" not in part:
                    raise ValueError("Bad format")
                key, value = part.split(":", 1)
                fields[key.strip()] = value.strip()

            for r in ["ID", "Name", "Dept", "Semester"]:
                if r not in fields:
                    raise KeyError(f"Missing: {r}")

            sid  = fields["ID"]
            name = fields["Name"]
            dept = fields["Dept"]
            sem  = fields["Semester"]

            now  = time.time()
            if sid in self.last_scan_time and now - self.last_scan_time[sid] < 5:
                return
            self.last_scan_time[sid] = now

            if sid in self.scanned_ids:
                self._flash_status(f"⚠  Already scanned: {name}", WARN)
                return

            current = datetime.datetime.now()
            new_row = {
                "ID": str(sid), "Name": str(name),
                "Department": str(dept), "Semester": str(sem),
                "Date": current.strftime("%Y-%m-%d"),
                "Time": current.strftime("%H:%M:%S"),
            }
            self.records.append(new_row)
            self.scanned_ids.add(sid)

            # Write to Excel
            wb = load_workbook(self.file_name)
            ws = wb.active
            ws.append([str(v) for v in new_row.values()])
            wb.save(self.file_name)

            self._update_stats()
            self._refresh_table()
            self.lbl_last_name.config(text=name)
            self.lbl_last_meta.config(text=f"{dept}  ·  Sem {sem}  ·  {new_row['Time']}")
            self._flash_status(f"✓  Saved: {name}  ({sid})", ACCENT)

        except (KeyError, ValueError):
            pass
        except Exception as e:
            self.statusbar.config(text=f"Error: {e}")

    # ──────────────────────────────────────────
    #  Table rendering
    # ──────────────────────────────────────────
    def _refresh_table(self):
        q    = self.search_var.get().lower().strip()
        dept = self.dept_var.get()
        if q == "search name or id…":
            q = ""

        filtered = [
            r for r in self.records
            if (not q or q in r["Name"].lower() or q in r["ID"])
            and (dept == "All" or r["Department"] == dept)
        ]

        self.tree.delete(*self.tree.get_children())
        for i, r in enumerate(filtered):
            tag = "even" if i % 2 == 0 else "odd"
            self.tree.insert("", "end", tags=(tag,),
                             values=(i+1, r["ID"], r["Name"],
                                     r["Department"], r["Semester"],
                                     r["Date"], r["Time"]))

        total = len(filtered)
        self.statusbar.config(
            text=f"Showing {total} record{'s' if total != 1 else ''}" +
                 (f"  (filtered from {len(self.records)})" if total != len(self.records) else ""))

    def _update_stats(self):
        self.lbl_total.config(text=str(len(self.records)))
        self.lbl_unique.config(text=str(len(self.scanned_ids)))

    # ──────────────────────────────────────────
    #  Session timer
    # ──────────────────────────────────────────
    def _tick_timer(self):
        if not self.session_running:
            return
        elapsed = int(time.time() - self.session_start)
        m, s = divmod(elapsed, 60)
        self.lbl_timer.config(text=f"{m:02d}:{s:02d}")
        self.root.after(1000, self._tick_timer)

    # ──────────────────────────────────────────
    #  Status flash
    # ──────────────────────────────────────────
    def _flash_status(self, msg, color):
        self.statusbar.config(text=msg, fg=color)
        self.root.after(3000, lambda: self.statusbar.config(text="", fg=TEXT_SEC))

    # ──────────────────────────────────────────
    #  Export
    # ──────────────────────────────────────────
    def export_excel(self):
        if not self.records:
            messagebox.showinfo("Export", "No records to export.")
            return
        dest = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=self.file_name)
        if dest:
            import shutil
            shutil.copy2(self.file_name, dest)
            messagebox.showinfo("Exported", f"Saved to:\n{dest}")

    def export_csv(self):
        if not self.records:
            messagebox.showinfo("Export", "No records to export.")
            return
        dest = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile=self.file_name.replace(".xlsx", ".csv"))
        if dest:
            pd.DataFrame(self.records).to_csv(dest, index=False)
            messagebox.showinfo("Exported", f"CSV saved to:\n{dest}")

    # ──────────────────────────────────────────
    #  Cleanup on close
    # ──────────────────────────────────────────
    def on_close(self):
        self.scanning = False
        self.session_running = False
        if self.cap:
            self.cap.release()
        self.root.destroy()


# ─────────────────────────────────────────────
#  Entry point
# ─────────────────────────────────────────────
if __name__ == "__main__":
    root = tk.Tk()
    app  = QRAttendanceApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.mainloop()