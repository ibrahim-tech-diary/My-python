import cv2
from pyzbar.pyzbar import decode
import pandas as pd
import datetime
import time
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk
from openpyxl import load_workbook

# ── Colors ──────────────────────────────────────────────────────────
BG         = "#F5F4F0"
SIDEBAR_BG = "#FFFFFF"
HEADER_BG  = "#FFFFFF"
ACCENT     = "#1D9E75"
ACCENT_DIM = "#0F6E56"
WARN       = "#E24B4A"
TEXT_PRI   = "#1A1A1A"
TEXT_SEC   = "#6B6B68"
BORDER     = "#E5E4DE"
ROW_ALT    = "#F9F8F5"

FONT_HEAD  = ("Segoe UI", 12, "bold")
FONT_BODY  = ("Segoe UI", 10)
FONT_SMALL = ("Segoe UI", 9)
FONT_MONO  = ("Consolas", 10)
FONT_NUM   = ("Segoe UI", 20, "bold")

CAM_W, CAM_H = 234, 176   # camera canvas pixel size


class QRAttendanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("QR Attendance System")
        self.root.geometry("1050x660")
        self.root.minsize(860, 540)
        self.root.configure(bg=BG)

        self.scanning        = False
        self.cap             = None
        self.cam_thread      = None
        self.scanned_ids     = set()
        self.last_scan_time  = {}
        self.records         = []
        self.session_start   = None
        self.session_running = False
        self.file_name       = None
        self._photo          = None

        self._build_ui()
        self._new_session()

    # ── Build UI ────────────────────────────────────────────────────

    def _build_ui(self):
        # Top bar
        topbar = tk.Frame(self.root, bg=HEADER_BG, height=48)
        topbar.pack(fill="x")
        topbar.pack_propagate(False)

        tk.Frame(topbar, bg=ACCENT, width=4).pack(side="left", fill="y")
        tk.Label(topbar, text="QR Attendance System",
                 font=("Segoe UI", 12, "bold"),
                 bg=HEADER_BG, fg=TEXT_PRI).pack(side="left", padx=14)

        self.lbl_session = tk.Label(topbar, text="",
                                    font=FONT_SMALL, bg=HEADER_BG, fg=TEXT_SEC)
        self.lbl_session.pack(side="left")

        self.lbl_status = tk.Label(topbar, text="● Camera off",
                                   font=FONT_SMALL, bg=HEADER_BG, fg=TEXT_SEC)
        self.lbl_status.pack(side="right", padx=14)

        tk.Frame(self.root, bg=BORDER, height=1).pack(fill="x")

        # Body
        body = tk.Frame(self.root, bg=BG)
        body.pack(fill="both", expand=True)

        sidebar = tk.Frame(body, bg=SIDEBAR_BG, width=258)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        tk.Frame(body, bg=BORDER, width=1).pack(side="left", fill="y")

        main = tk.Frame(body, bg=BG)
        main.pack(side="left", fill="both", expand=True)

        self._build_sidebar(sidebar)
        self._build_main(main)

    def _build_sidebar(self, p):
        tk.Frame(p, bg=SIDEBAR_BG, height=10).pack()

        # Camera canvas (fixed pixel dimensions)
        cam_border = tk.Frame(p, bg=BORDER, padx=1, pady=1)
        cam_border.pack(padx=12)

        self.cam_canvas = tk.Canvas(cam_border,
                                    width=CAM_W, height=CAM_H,
                                    bg="#EAEAE6", highlightthickness=0)
        self.cam_canvas.pack()
        self._draw_cam_placeholder()

        # Start/Stop button
        self.btn_scan = tk.Button(
            p, text="▶  Start Scanner",
            font=("Segoe UI", 10, "bold"),
            bg=ACCENT, fg="white",
            activebackground=ACCENT_DIM, activeforeground="white",
            relief="flat", cursor="hand2", pady=7, bd=0,
            command=self.toggle_scanner)
        self.btn_scan.pack(fill="x", padx=12, pady=(8, 0))

        tk.Frame(p, bg=BORDER, height=1).pack(fill="x", padx=12, pady=(8, 0))

        # Stats row
        stats = tk.Frame(p, bg=SIDEBAR_BG)
        stats.pack(fill="x", padx=12, pady=(6, 0))

        lf = tk.Frame(stats, bg="#F3F2EE")
        lf.pack(side="left", fill="both", expand=True, padx=(0, 4), ipady=6, ipadx=4)
        tk.Label(lf, text="Scanned", font=FONT_SMALL,
                 bg="#F3F2EE", fg=TEXT_SEC).pack()
        self.lbl_total = tk.Label(lf, text="0", font=FONT_NUM,
                                  bg="#F3F2EE", fg=TEXT_PRI)
        self.lbl_total.pack()

        rf = tk.Frame(stats, bg="#F3F2EE")
        rf.pack(side="left", fill="both", expand=True, padx=(4, 0), ipady=6, ipadx=4)
        tk.Label(rf, text="Unique", font=FONT_SMALL,
                 bg="#F3F2EE", fg=TEXT_SEC).pack()
        self.lbl_unique = tk.Label(rf, text="0", font=FONT_NUM,
                                   bg="#F3F2EE", fg=TEXT_PRI)
        self.lbl_unique.pack()

        # Session timer
        timer_f = tk.Frame(p, bg="#F3F2EE")
        timer_f.pack(fill="x", padx=12, pady=(6, 0), ipady=5)
        tk.Label(timer_f, text="Session",
                 font=FONT_SMALL, bg="#F3F2EE", fg=TEXT_SEC).pack(side="left", padx=8)
        self.lbl_timer = tk.Label(timer_f, text="00:00",
                                  font=FONT_MONO, bg="#F3F2EE", fg=TEXT_PRI)
        self.lbl_timer.pack(side="right", padx=8)

        # Last scan card
        tk.Frame(p, bg=BORDER, height=1).pack(fill="x", padx=12, pady=(8, 0))
        last_f = tk.Frame(p, bg="#EAF3EE")
        last_f.pack(fill="x", padx=12, ipady=6, ipadx=8)
        tk.Label(last_f, text="Last scanned",
                 font=FONT_SMALL, bg="#EAF3EE", fg=ACCENT_DIM).pack(anchor="w")
        self.lbl_last_name = tk.Label(last_f, text="—",
                                      font=("Segoe UI", 10, "bold"),
                                      bg="#EAF3EE", fg=TEXT_PRI)
        self.lbl_last_name.pack(anchor="w")
        self.lbl_last_meta = tk.Label(last_f, text="—",
                                      font=FONT_SMALL, bg="#EAF3EE", fg=TEXT_SEC)
        self.lbl_last_meta.pack(anchor="w")

        # Bottom action buttons
        tk.Frame(p, bg=BORDER, height=1).pack(fill="x", padx=12, pady=(8, 0))
        btn_row = tk.Frame(p, bg=SIDEBAR_BG)
        btn_row.pack(fill="x", padx=12, pady=(6, 10))

        tk.Button(btn_row, text="New Session",
                  font=FONT_SMALL, bg=SIDEBAR_BG, fg=TEXT_SEC,
                  relief="groove", cursor="hand2",
                  command=self._new_session
                  ).pack(side="left", fill="x", expand=True, padx=(0, 4))

        tk.Button(btn_row, text="Export Excel",
                  font=FONT_SMALL, bg=SIDEBAR_BG, fg=ACCENT,
                  relief="groove", cursor="hand2",
                  command=self.export_excel
                  ).pack(side="left", fill="x", expand=True, padx=(4, 0))

    def _draw_cam_placeholder(self):
        self.cam_canvas.delete("all")
        cx, cy = CAM_W // 2, CAM_H // 2
        self.cam_canvas.create_rectangle(cx-30, cy-20, cx+30, cy+20,
                                         outline="#BBBBBB", width=2)
        self.cam_canvas.create_oval(cx-10, cy-10, cx+10, cy+10,
                                    outline="#BBBBBB", width=2)
        self.cam_canvas.create_text(cx, cy+38,
                                    text="Camera preview",
                                    font=FONT_SMALL, fill="#BBBBBB")

    def _build_main(self, p):
        # Toolbar
        toolbar = tk.Frame(p, bg=HEADER_BG, height=42)
        toolbar.pack(fill="x")
        toolbar.pack_propagate(False)
        tk.Frame(p, bg=BORDER, height=1).pack(fill="x")

        self.search_var = tk.StringVar()
        tk.Entry(toolbar, textvariable=self.search_var,
                 font=FONT_BODY, relief="flat",
                 bg="#F3F2EE", fg=TEXT_PRI,
                 insertbackground=TEXT_PRI, width=22
                 ).pack(side="left", padx=(12, 8), pady=7, ipady=3)

        tk.Label(toolbar, text="Dept:",
                 font=FONT_SMALL, bg=HEADER_BG, fg=TEXT_SEC).pack(side="left")

        self.dept_var = tk.StringVar(value="All")
        dept_cb = ttk.Combobox(toolbar, textvariable=self.dept_var,
                               values=["All", "CSE", "EEE", "ME", "BBA", "CE"],
                               state="readonly", width=7, font=FONT_SMALL)
        dept_cb.pack(side="left", padx=6)

        # Bind AFTER both vars exist
        self.search_var.trace_add("write", lambda *_: self._refresh_table())
        dept_cb.bind("<<ComboboxSelected>>", lambda _: self._refresh_table())

        tk.Button(toolbar, text="Export CSV",
                  font=FONT_SMALL, bg=HEADER_BG, fg=TEXT_SEC,
                  relief="groove", cursor="hand2",
                  command=self.export_csv).pack(side="right", padx=12)

        # Treeview
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("A.Treeview",
                        background=HEADER_BG, fieldbackground=HEADER_BG,
                        foreground=TEXT_PRI, font=FONT_BODY,
                        rowheight=30, borderwidth=0)
        style.configure("A.Treeview.Heading",
                        background=HEADER_BG, foreground=TEXT_SEC,
                        font=("Segoe UI", 8, "bold"), relief="flat")
        style.map("A.Treeview",
                  background=[("selected", "#C8E6D8")],
                  foreground=[("selected", TEXT_PRI)])

        cols = ("#", "ID", "Name", "Department", "Semester", "Date", "Time")
        tree_frame = tk.Frame(p, bg=BG)
        tree_frame.pack(fill="both", expand=True)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                                 style="A.Treeview", yscrollcommand=vsb.set)
        vsb.configure(command=self.tree.yview)

        for col, w in zip(cols, [36, 85, 155, 95, 75, 95, 75]):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, minwidth=w, anchor="w")

        self.tree.tag_configure("odd",  background=HEADER_BG)
        self.tree.tag_configure("even", background=ROW_ALT)

        vsb.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)

        self.statusbar = tk.Label(p, text="Ready — scan a QR code to begin",
                                  font=FONT_SMALL, bg=BORDER,
                                  fg=TEXT_SEC, anchor="w", pady=3, padx=8)
        self.statusbar.pack(fill="x", side="bottom")

    # ── Session ─────────────────────────────────────────────────────

    def _new_session(self):
        if self.scanning:
            self.toggle_scanner()
        ts = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        self.file_name       = f"attendance_{ts}.xlsx"
        self.scanned_ids     = set()
        self.last_scan_time  = {}
        self.records         = []
        self.lbl_total.config(text="0")
        self.lbl_unique.config(text="0")
        self.lbl_timer.config(text="00:00")
        self.lbl_last_name.config(text="—")
        self.lbl_last_meta.config(text="—")
        pd.DataFrame(columns=["ID", "Name", "Department",
                               "Semester", "Date", "Time"]
                     ).to_excel(self.file_name, index=False)
        self.lbl_session.config(text=f"  {self.file_name}")
        self._refresh_table()
        self.statusbar.config(text=f"New session → {self.file_name}", fg=TEXT_SEC)

    # ── Scanner ──────────────────────────────────────────────────────

    def toggle_scanner(self):
        if not self.scanning:
            self._start_scanner()
        else:
            self._stop_scanner()

    def _start_scanner(self):
        self.cap = cv2.VideoCapture(0)
        if not self.cap.isOpened():
            messagebox.showerror("Camera Error",
                                 "Cannot open camera.\n"
                                 "Make sure a webcam is connected.")
            self.cap = None
            return
        self.scanning        = True
        self.session_start   = time.time()
        self.session_running = True
        self.btn_scan.config(text="■  Stop Scanner",
                             bg=WARN, activebackground="#A32D2D")
        self.lbl_status.config(text="● Scanning…", fg=ACCENT)
        self._tick_timer()
        self.cam_thread = threading.Thread(target=self._scan_loop, daemon=True)
        self.cam_thread.start()

    def _stop_scanner(self):
        self.scanning        = False
        self.session_running = False
        if self.cap:
            self.cap.release()
            self.cap = None
        self.btn_scan.config(text="▶  Start Scanner",
                             bg=ACCENT, activebackground=ACCENT_DIM)
        self.lbl_status.config(text="● Camera off", fg=TEXT_SEC)
        self._draw_cam_placeholder()
        self._photo = None

    # ── Scan loop (background thread) ───────────────────────────────

    def _scan_loop(self):
        while self.scanning and self.cap and self.cap.isOpened():
            ret, frame = self.cap.read()
            if not ret:
                break

            for barcode in decode(frame):
                qr_data = barcode.data.decode("utf-8").strip()
                x, y, w, h = barcode.rect
                cv2.rectangle(frame, (x, y), (x+w, y+h), (29, 158, 117), 2)
                cv2.putText(frame, qr_data[:28], (x, max(y-6, 14)),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.42, (29, 158, 117), 1)
                self.root.after(0, self._process_qr, qr_data)

            # Scale to fit canvas keeping aspect ratio
            fh, fw = frame.shape[:2]
            ratio   = min(CAM_W / fw, CAM_H / fh)
            new_w   = int(fw * ratio)
            new_h   = int(fh * ratio)
            preview = cv2.resize(frame, (new_w, new_h))
            preview = cv2.cvtColor(preview, cv2.COLOR_BGR2RGB)
            photo   = ImageTk.PhotoImage(image=Image.fromarray(preview))
            self.root.after(0, self._update_cam, photo, new_w, new_h)
            time.sleep(0.033)

    def _update_cam(self, photo, w, h):
        self._photo = photo
        x = (CAM_W - w) // 2
        y = (CAM_H - h) // 2
        self.cam_canvas.delete("all")
        self.cam_canvas.create_image(x, y, anchor="nw", image=photo)

    # ── QR processing ────────────────────────────────────────────────

    def _process_qr(self, qr_data):
        try:
            fields = {}
            for part in qr_data.split(";"):
                part = part.strip()
                if ":" not in part:
                    continue
                key, value = part.split(":", 1)
                fields[key.strip()] = value.strip()

            for req in ["ID", "Name", "Dept", "Semester"]:
                if req not in fields:
                    raise KeyError(req)

            sid  = fields["ID"]
            name = fields["Name"]
            dept = fields["Dept"]
            sem  = fields["Semester"]

            now = time.time()
            if sid in self.last_scan_time and now - self.last_scan_time[sid] < 5:
                return
            self.last_scan_time[sid] = now

            if sid in self.scanned_ids:
                self._set_status(f"⚠  Already scanned: {name}", WARN)
                return

            current = datetime.datetime.now()
            new_row = {
                "ID":         str(sid),
                "Name":       str(name),
                "Department": str(dept),
                "Semester":   str(sem),
                "Date":       current.strftime("%Y-%m-%d"),
                "Time":       current.strftime("%H:%M:%S"),
            }
            self.records.append(new_row)
            self.scanned_ids.add(sid)

            wb = load_workbook(self.file_name)
            ws = wb.active
            ws.append([str(v) for v in new_row.values()])
            wb.save(self.file_name)

            self.lbl_total.config(text=str(len(self.records)))
            self.lbl_unique.config(text=str(len(self.scanned_ids)))
            self.lbl_last_name.config(text=name)
            self.lbl_last_meta.config(
                text=f"{dept}  ·  Sem {sem}  ·  {new_row['Time']}")
            self._refresh_table()
            self._set_status(f"✓  Saved: {name}  ({sid})", ACCENT)

        except KeyError:
            pass
        except Exception as e:
            self._set_status(f"Error: {e}", WARN)

    # ── Table ────────────────────────────────────────────────────────

    def _refresh_table(self):
        q    = self.search_var.get().lower().strip()
        dept = self.dept_var.get() if hasattr(self, "dept_var") else "All"

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
        n   = len(filtered)
        txt = f"Showing {n} record{'s' if n != 1 else ''}"
        if n != len(self.records):
            txt += f"  (filtered from {len(self.records)})"
        self.statusbar.config(text=txt, fg=TEXT_SEC)

    # ── Timer ────────────────────────────────────────────────────────

    def _tick_timer(self):
        if not self.session_running:
            return
        elapsed = int(time.time() - self.session_start)
        m, s = divmod(elapsed, 60)
        self.lbl_timer.config(text=f"{m:02d}:{s:02d}")
        self.root.after(1000, self._tick_timer)

    # ── Status bar helper ────────────────────────────────────────────

    def _set_status(self, msg, color=None):
        self.statusbar.config(text=msg, fg=color or TEXT_SEC)
        self.root.after(3500, lambda: self.statusbar.config(
            text="", fg=TEXT_SEC))

    # ── Export ───────────────────────────────────────────────────────

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

    # ── Cleanup ──────────────────────────────────────────────────────

    def on_close(self):
        self.scanning        = False
        self.session_running = False
        if self.cap:
            self.cap.release()
        self.root.destroy()


# ── Entry point ──────────────────────────────────────────────────────
if __name__ == "__main__":
    root = tk.Tk()
    app  = QRAttendanceApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.mainloop()
