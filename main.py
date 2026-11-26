import os, re, sqlite3, sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk
import openpyxl
from openpyxl.styles import Font, PatternFill

def get_data_dir():
    if getattr(sys, "frozen", False):
        d = Path(sys.executable).resolve().parent
        try:
            t = d / ".test"; t.touch(); t.unlink()
            return d
        except OSError:
            f = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "BAF_Register"
            f.mkdir(parents=True, exist_ok=True)
            return f
    return Path(__file__).resolve().parent

def reg_key(v):
    return "".join((v or "").split()).upper()

def safe_name(s):
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", s.strip()).rstrip(". ") or "Class_Export"

class DB:
    def __init__(self, path):
        self.conn = sqlite3.connect(str(path))
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON;")
        with self.conn:
            self.conn.execute("CREATE TABLE IF NOT EXISTS classes (id INTEGER PRIMARY KEY, name TEXT NOT NULL UNIQUE COLLATE NOCASE);")
            self.conn.execute("CREATE TABLE IF NOT EXISTS students (id INTEGER PRIMARY KEY, class_id INTEGER NOT NULL REFERENCES classes(id) ON DELETE CASCADE, name TEXT, reg TEXT, reg_key TEXT, UNIQUE(class_id, reg_key));")

    def classes(self): return self.conn.execute("SELECT id, name FROM classes ORDER BY name COLLATE NOCASE;").fetchall()
    def add_class(self, name):
        with self.conn: return self.conn.execute("INSERT INTO classes (name) VALUES (?);", (name,)).lastrowid
    def rename_class(self, cid, name):
        with self.conn: self.conn.execute("UPDATE classes SET name=? WHERE id=?;", (name, cid))
    def del_class(self, cid):
        with self.conn: self.conn.execute("DELETE FROM classes WHERE id=?;", (cid,))
    def students(self, cid): return self.conn.execute("SELECT id, name, reg FROM students WHERE class_id=? ORDER BY id;", (cid,)).fetchall()
    def dup_reg(self, cid, rk, excl=None):
        if not rk: return False
        sql = "SELECT 1 FROM students WHERE class_id=? AND reg_key=?" + (" AND id!=?" if excl else "")
        return self.conn.execute(sql, (cid, rk, excl) if excl else (cid, rk)).fetchone() is not None
    def add_student(self, cid, name, reg):
        with self.conn: return self.conn.execute("INSERT INTO students (class_id, name, reg, reg_key) VALUES (?,?,?,?);", (cid, name, reg, reg_key(reg) or None)).lastrowid
    def upd_student(self, sid, name, reg):
        with self.conn: self.conn.execute("UPDATE students SET name=?, reg=?, reg_key=? WHERE id=?;", (name, reg, reg_key(reg) or None, sid))
    def del_student(self, sid):
        with self.conn: self.conn.execute("DELETE FROM students WHERE id=?;", (sid,))

class App(tk.Tk):
    def __init__(self, db, base_dir):
        super().__init__()
        self.db, self.base_dir = db, base_dir
        self.cid, self.cname = None, ""
        self.cls_list, self.rows, self.iid_map = [], [], {}
        self.editor, self.ed_rc = None, None
        self.title("BAF | Simple Student Register")
        self.configure(bg="#0b1220")
        self.geometry("900x700")
        self.protocol("WM_DELETE_WINDOW", self._close)
        self._setup_ui()
        self._load_classes()

    def _setup_ui(self):
        st = ttk.Style(self)
        st.theme_use("clam")
        st.configure("T.Treeview", background="#101e30", foreground="#dfebff", fieldbackground="#101e30", bordercolor="#28415b", rowheight=30, font=("Segoe UI", 11))
        st.map("T.Treeview", background=[("selected", "#20516a")], foreground=[("selected", "#ffffff")])
        st.configure("T.Treeview.Heading", background="#193449", foreground="#63e6df", relief="flat", font=("Segoe UI", 11, "bold"), padding=8)
        root = tk.Frame(self, bg="#0b1220")
        root.pack(fill="both", expand=True, padx=16, pady=16)
        side = tk.Frame(root, bg="#0b1220", width=220)
        side.pack(side="left", fill="y", padx=(0, 16))
        side.pack_propagate(False)
        tk.Label(side, text="BAF\nREGISTER", bg="#0b1220", fg="#63e6df", font=("Segoe UI", 22, "bold"), justify="left", anchor="w").pack(fill="x", pady=(0, 12))
        self.lb = tk.Listbox(side, bg="#101e30", fg="#dfebff", selectbackground="#20516a", selectforeground="#63e6df", font=("Segoe UI", 11), height=10, bd=1, relief="solid", highlightthickness=0, exportselection=False)
        self.lb.pack(fill="x", pady=(0, 12))
        self.lb.bind("<<ListboxSelect>>", self._on_cls_pick)
        self.btns = []
        for txt, cmd in [("+ New class", self._new_cls), ("Rename", self._ren_cls), ("Delete class", self._del_cls), ("Export this class", self._exp_cls)]:
            b = tk.Button(side, text=txt, command=cmd, bg="#163449", fg="#dfebff", activebackground="#20516a", activeforeground="#63e6df", font=("Segoe UI", 11, "bold"), anchor="w", padx=12, pady=8, relief="flat", highlightthickness=1, highlightbackground="#247b91")
            b.pack(fill="x", pady=(0, 8))
            self.btns.append(b)
        main = tk.Frame(root, bg="#0b1220")
        main.pack(side="left", fill="both", expand=True)
        self.hdr = tk.Label(main, text="Select or create a class", bg="#0b1220", fg="#dfebff", font=("Segoe UI", 17, "bold"), anchor="w")
        self.hdr.pack(fill="x", pady=(0, 10))
        self.svar = tk.StringVar(value="Search name or registration number...")
        self.s_ph = True
        self.sent = tk.Entry(main, textvariable=self.svar, bg="#101e30", fg="#6c8bb0", insertbackground="#63e6df", font=("Segoe UI", 11), relief="solid", bd=1)
        self.sent.pack(fill="x", pady=(0, 12), ipady=6)
        self.sent.bind("<FocusIn>", self._s_in)
        self.sent.bind("<FocusOut>", self._s_out)
        self.svar.trace_add("write", lambda *_: self._render() if not self.s_ph else None)
        tf = tk.Frame(main, bg="#101e30", bd=1, relief="solid")
        tf.pack(fill="both", expand=True, pady=(0, 12))
        self.tree = ttk.Treeview(tf, columns=("n", "r"), show="headings", style="T.Treeview", selectmode="browse")
        self.tree.heading("n", text="STUDENT NAME", anchor="w")
        self.tree.heading("r", text="REGISTRATION NUMBER", anchor="w")
        self.tree.column("n", width=360, anchor="w")
        self.tree.column("r", width=220, anchor="w", stretch=False)
        vsb = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self.tree.bind("<Button-1>", self._click)
        self.tree.bind("<Return>", lambda _: self._edit(0, 0))
        self.stat = tk.Label(main, text="Type a name, press Enter, type the registration number, then press Enter for the next student.", bg="#101e30", fg="#9ec0e6", font=("Segoe UI", 10), anchor="w", padx=12, pady=8, relief="solid", bd=1)
        self.stat.pack(fill="x")
    def _s_in(self, _):
        if self.s_ph: self.s_ph = False; self.svar.set(""); self.sent.config(fg="#dfebff")

    def _s_out(self, _):
        if not self.svar.get().strip(): self.s_ph = True; self.sent.config(fg="#6c8bb0"); self.svar.set("Search name or registration number...")

    def _msg(self, txt, err=False):
        self.stat.config(text=txt, bg="#3a1620" if err else "#101e30", fg="#ffd6d6" if err else "#9ec0e6")

    def _load_classes(self, pick_id=None):
        self._kill_ed()
        self.lb.delete(0, tk.END)
        self.cls_list = [(r["id"], r["name"]) for r in self.db.classes()]
        for _, n in self.cls_list: self.lb.insert(tk.END, n)
        if not self.cls_list:
            self.cid, self.cname = None, ""
            self.hdr.config(text="Create your first class.")
            self.rows = [{"id": None, "n": "", "r": ""} for _ in range(20)]
            self._render()
            for b in self.btns[1:]: b.config(state="disabled")
            return
        for b in self.btns[1:]: b.config(state="normal")
        idx = next((i for i, (c, _) in enumerate(self.cls_list) if c == pick_id), 0)
        self.lb.selection_set(idx)
        self._select_cls(idx)

    def _on_cls_pick(self, _):
        sel = self.lb.curselection()
        if sel: self._select_cls(sel)

    def _select_cls(self, idx):
        self._kill_ed()
        self.cid, self.cname = self.cls_list[idx]
        self.hdr.config(text=self.cname)
        sts = self.db.students(self.cid)
        self.rows = [{"id": s["id"], "n": s["name"] or "", "r": s["reg"] or ""} for s in sts]
        while len(self.rows) < 20: self.rows.append({"id": None, "n": "", "r": ""})
        if self.rows[-1]["n"] or self.rows[-1]["r"]: self.rows.append({"id": None, "n": "", "r": ""})
        self._render()
        self._msg("Type a name, press Enter, type the registration number, then press Enter for the next student.")

    def _render(self):
        self.tree.delete(*self.tree.get_children())
        self.iid_map.clear()
        q = "" if self.s_ph else self.svar.get().strip().lower()
        blank_i = next((i for i in range(len(self.rows) - 1, -1, -1) if not self.rows[i]["n"] and not self.rows[i]["r"]), -1)
        for i, r in enumerate(self.rows):
            if q and i != blank_i and q not in r["n"].lower() and q not in r["r"].lower(): continue
            iid = f"r{i}"
            self.tree.insert("", "end", iid=iid, values=(r["n"], r["r"]))
            self.iid_map[iid] = i

    def _click(self, ev):
        if not self.cid or self.tree.identify("region", ev.x, ev.y) != "cell": return
        iid, col = self.tree.identify_row(ev.y), self.tree.identify_column(ev.x)
        if iid in self.iid_map and col in ("#1", "#2"):
            self._save_ed()
            self.after(10, lambda r=self.iid_map[iid], c=(0 if col == "#1" else 1): self._edit(r, c))

    def _kill_ed(self):
        if self.editor:
            try: self.editor.destroy()
            except tk.TclError: pass
            self.editor, self.ed_rc = None, None

    def _edit(self, r, c):
        if not self.cid or r >= len(self.rows): return
        self._kill_ed()
        iid = f"r{r}"
        if not self.tree.exists(iid): return
        self.tree.selection_set(iid); self.tree.see(iid); self.update_idletasks()
        box = self.tree.bbox(iid, "#1" if c == 0 else "#2")
        if not box: return
        e = tk.Entry(self.tree, bg="#193449", fg="#ffffff", insertbackground="#63e6df", font=("Segoe UI", 11), relief="flat", highlightthickness=2, highlightcolor="#63e6df")
        e.place(x=box, y=box, width=box, height=box)
        e.insert(0, self.rows[r]["n"] if c == 0 else self.rows[r]["r"])
        e.select_range(0, tk.END); e.focus_set()
        self.editor, self.ed_rc = e, (r, c)
        e.bind("<Return>", self._enter)
        e.bind("<KP_Enter>", self._enter)
        e.bind("<Escape>", lambda _: self._kill_ed())
        e.bind("<FocusOut>", lambda _: self._save_ed())

    def _enter(self, _):
        if not self.ed_rc: return "break"
        r, c = self.ed_rc
        if not self._save_ed(): return "break"
        if c == 0:
            self.after(10, lambda: self._edit(r, 1))
        else:
            if r + 1 >= len(self.rows):
                self.rows.append({"id": None, "n": "", "r": ""})
                self._render()
            self.after(10, lambda: self._edit(r + 1, 0))
        return "break"

    def _save_ed(self):
        if not self.editor or not self.ed_rc: return True
        r, c = self.ed_rc
        try: val = self.editor.get()
        except tk.TclError: self._kill_ed(); return True
        self._kill_ed()
        row = self.rows[r]
        nn = val.strip() if c == 0 else row["n"].strip()
        nr = row["r"].strip().upper() if c == 0 else val.strip().upper()
        if nn == row["n"] and nr == row["r"]: return True
        if not nn and not nr:
            if row["id"]:
                self.db.del_student(row["id"])
                row.update({"id": None, "n": "", "r": ""})
                self._render(); self._msg("Saved automatically.")
            return True
        if reg_key(nr) and self.db.dup_reg(self.cid, reg_key(nr), row["id"]):
            self._msg("REJECTED: This registration number is already used in this class. Enter a different number.", True)
            row["n"], row["r"] = nn, ""
            self._render()
            self.after(10, lambda: self._edit(r, 1))
            return False
        try:
            if not row["id"]:
                row["id"] = self.db.add_student(self.cid, nn, nr)
                row["n"], row["r"] = nn, nr
                if self.rows[-1]["n"] or self.rows[-1]["r"]: self.rows.append({"id": None, "n": "", "r": ""})
                self._render(); self._msg("Student registered.")
            else:
                self.db.upd_student(row["id"], nn, nr)
                row["n"], row["r"] = nn, nr
                self._render(); self._msg("Saved automatically.")
            return True
        except sqlite3.IntegrityError:
            self._msg("REJECTED: This registration number is already used in this class. Your change was not saved.", True)
            self._select_cls(self.lb.curselection())
            return False

    def _new_cls(self):
        self._save_ed()
        name = simpledialog.askstring("New Class", "Enter class name:", parent=self)
        if not name or not name.strip(): return
        try:
            nid = self.db.add_class(name.strip())
            self._load_classes(nid)
        except sqlite3.IntegrityError:
            messagebox.showwarning("Duplicate Class", "A class with this name already exists.", parent=self)

    def _ren_cls(self):
        if not self.cid: return
        self._save_ed()
        name = simpledialog.askstring("Rename Class", "Enter new class name:", initialvalue=self.cname, parent=self)
        if not name or not name.strip(): return
        try:
            self.db.rename_class(self.cid, name.strip())
            self._load_classes(self.cid)
        except sqlite3.IntegrityError:
            messagebox.showwarning("Duplicate Class", "A class with this name already exists.", parent=self)

    def _del_cls(self):
        if not self.cid: return
        self._kill_ed()
        if messagebox.askyesno("Delete Class", f"Delete '{self.cname}' and all its students?", parent=self):
            self.db.del_class(self.cid)
            self._load_classes()

    def _exp_cls(self):
        if not self.cid: return
        self._save_ed()
        exp = self.base_dir / "Exports"
        exp.mkdir(parents=True, exist_ok=True)
        path = filedialog.asksaveasfilename(parent=self, title="Export Class", initialdir=str(exp), initialfile=f"{safe_name(self.cname)}.xlsx", defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path: return
        if not path.lower().endswith(".xlsx"): path += ".xlsx"
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = safe_name(self.cname)[:31]
            ws.append(["STUDENT NAME", "REGISTRATION NUMBER"])
            for col in (1, 2):
                c = ws.cell(1, col)
                c.font = Font(name="Segoe UI", bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="203957")
            sts = self.db.students(self.cid)
            for s in sts: ws.append([s["name"] or "", s["reg"] or ""])
            ws.column_dimensions["A"].width = 45
            ws.column_dimensions["B"].width = 30
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = f"A1:B{max(1, len(sts) + 1)}"
            wb.save(path)
            self._msg(f"Exported to: {path}")
            messagebox.showinfo("Export Complete", f"Saved to:\n{path}", parent=self)
        except Exception as e:
            messagebox.showerror("Export Failed", f"Could not export (close the file if it is open in Excel):\n{e}", parent=self)

    def _close(self):
        self._save_ed(); self.db.conn.close(); self.destroy()

if __name__ == "__main__":
    base = get_data_dir()
    app = App(DB(base / "baf_student_register.db"), base)
    app.mainloop()
