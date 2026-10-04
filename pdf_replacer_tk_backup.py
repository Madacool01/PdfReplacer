"""
Dad's PDF Keyword Replacer - easy to use app.
- Select PDFs or folder, find/replace one keyword pair
- Keep original font/size or set custom, plus bold/italic/underline/highlight
- Saves to new folder, originals untouched
- Run: python pdf_replacer.py  |  Build EXE: pyinstaller --onefile --windowed pdf_replacer.py
Requires: pymupdf
"""
import os
import csv
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser
from pathlib import Path

import customtkinter as ctk

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

try:
    import pymupdf
except ImportError:
    raise SystemExit("Missing dependency: pip install pymupdf")

# ---------- font helpers ----------

BASE_FONTS = {
    ("helv", False, False): "helv",
    ("helv", True, False): "hebo",
    ("helv", False, True): "heit",
    ("helv", True, True): "hebi",
    ("cour", False, False): "cour",
    ("cour", True, False): "cobo",
    ("cour", False, True): "coit",
    ("cour", True, True): "cobi",
    ("tiro", False, False): "tiro",
    ("tiro", True, False): "tibo",
    ("tiro", False, True): "tiit",
    ("tiro", True, True): "tibi",
}
FAMILY_LABELS = {"helv": "Helvetica", "cour": "Courier", "tiro": "Times"}


def detect_family_and_flags(span_font: str):
    f = (span_font or "").lower()
    if "cour" in f or "courier" in f or "mono" in f:
        fam = "cour"
    elif "time" in f or "tiro" in f or "serif" in f or "georgia" in f:
        fam = "tiro"
    else:
        fam = "helv"
    bold = any(k in f for k in ("bold", "black", "heavy", "demi", "boto", "hebo", "cobo", "tibo"))
    italic = any(k in f for k in ("italic", "oblique", "it ", "heit", "coit", "tiit"))
    return fam, bold, italic


def int_to_rgb(color_int):
    if color_int is None:
        return (0, 0, 0)
    try:
        r = ((color_int >> 16) & 255) / 255
        g = ((color_int >> 8) & 255) / 255
        b = (color_int & 255) / 255
        return (r, g, b)
    except Exception:
        return (0, 0, 0)


def hex_to_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))


def rgb_to_hex(rgb):
    return "#%02x%02x%02x" % tuple(int(max(0, min(1, c)) * 255) for c in rgb)


# ---------- core replace logic ----------

def get_span_style(page, rect):
    """Find the span with max overlap to inherit font/size/color/baseline."""
    try:
        d = page.get_text("dict")
    except Exception:
        return None
    best = None
    best_area = 0
    for block in d.get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                sb = pymupdf.Rect(span["bbox"])
                inter = sb & rect
                if inter.is_empty:
                    continue
                area = abs(inter.get_area())
                if area > best_area:
                    best_area = area
                    best = span
    return best


def matches_filter(clip_text, find_str, match_case, whole_word):
    if not match_case:
        ct, fs = clip_text.lower(), find_str.lower()
    else:
        ct, fs = clip_text, find_str
    idx = ct.find(fs)
    if idx < 0:
        return False
    if whole_word:
        before = ct[idx - 1] if idx > 0 else " "
        after = ct[idx + len(fs)] if idx + len(fs) < len(ct) else " "
        if before.isalnum() or before == "_" or after.isalnum() or after == "_":
            return False
    return True


def find_match_rects(page, find_str, match_case, whole_word):
    """Use search_for for location, then filter by case/whole-word via clip text."""
    try:
        rects = page.search_for(find_str)
    except Exception:
        return []
    out = []
    for r in rects:
        # inflate slightly to capture full text for verification
        clip = pymupdf.Rect(r.x0 - 2, r.y0 - 2, r.x1 + 2, r.y1 + 2)
        try:
            clip_text = page.get_text(clip=clip).strip().replace("\n", " ")
        except Exception:
            clip_text = find_str
        # clip may contain surrounding words; only verify the match exists
        if matches_filter(clip_text, find_str, match_case, whole_word):
            out.append(pymupdf.Rect(r))
        elif not whole_word and not match_case:
            out.append(pymupdf.Rect(r))  # search_for already did fuzzy match
    # if strict filtering removed everything but search found something and user wants loose match, fall back
    return out


def count_matches(pdf_path, find_str, match_case, whole_word):
    total = 0
    try:
        doc = pymupdf.open(pdf_path)
    except Exception:
        return 0
    with doc:
        for page in doc:
            total += len(find_match_rects(page, find_str, match_case, whole_word))
    return total


def replace_in_pdf(src, dst, find_str, replace_str, opts):
    """
    opts: dict with match_case, whole_word, keep_original,
          bold, italic, underline, highlight, highlight_hex,
          custom_family(helv/cour/tiro), custom_size, custom_hex
    Returns (replacements_count, warning_or_None)
    """
    match_case = opts["match_case"]
    whole_word = opts["whole_word"]
    keep_original = opts["keep_original"]

    try:
        doc = pymupdf.open(src)
    except Exception as e:
        return 0, f"Cannot open: {e}"
    if doc.is_encrypted:
        try:
            doc.authenticate("")
        except Exception:
            pass
        if doc.is_encrypted:
            doc.close()
            return 0, "Skipped (encrypted/password protected)"

    total = 0
    try:
        for page in doc:
            rects = find_match_rects(page, find_str, match_case, whole_word)
            if not rects:
                continue
            # gather styles BEFORE redacting
            jobs = []
            for rect in rects:
                span = get_span_style(page, rect)
                if span and keep_original:
                    fam, b, i = detect_family_and_flags(span.get("font", ""))
                    size = span.get("size", 11)
                    rgb = int_to_rgb(span.get("color", 0))
                    base_y = span.get("origin", [rect.x0, rect.y1 - 2])[1]
                else:
                    fam, b, i = opts["custom_family"], False, False
                    size = opts["custom_size"]
                    rgb = hex_to_rgb(opts["custom_hex"])
                    base_y = rect.y1 - 2
                # apply bold/italic toggles on top
                if opts["bold"]:
                    b = True
                if opts["italic"]:
                    i = True
                if keep_original and not opts["bold"] and not opts["italic"]:
                    pass  # keep detected b/i
                elif keep_original:
                    pass  # toggles already merged above
                fontname = BASE_FONTS.get((fam, b, i), "helv")
                if not keep_original:
                    fontname = BASE_FONTS.get((opts["custom_family"], opts["bold"], opts["italic"]), "helv")
                    size = opts["custom_size"]
                    rgb = hex_to_rgb(opts["custom_hex"])
                # shrink-to-fit so a longer replacement doesn't overlap next word
                try:
                    w = pymupdf.get_text_length(replace_str if replace_str else " ", fontname=fontname, fontsize=size)
                    avail = max(rect.width, 5)
                    if w > avail + 1 and len(replace_str) > 0:
                        size = max(6.0, size * (avail / w))
                except Exception:
                    pass
                jobs.append({"rect": rect, "fontname": fontname, "size": size, "rgb": rgb, "base_y": base_y})

            # redact old text
            for j in jobs:
                try:
                    page.add_redact_annot(j["rect"], fill=(1, 1, 1))
                except Exception:
                    pass
            try:
                page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE)
            except TypeError:
                page.apply_redactions()

            # insert new text
            for j in jobs:
                rect = j["rect"]
                x = rect.x0
                y = j["base_y"]
                # keep inside page
                if x < 0:
                    x = 0
                try:
                    page.insert_text(
                        pymupdf.Point(x, y),
                        replace_str,
                        fontname=j["fontname"],
                        fontsize=j["size"],
                        color=j["rgb"],
                    )
                except Exception:
                    # fallback to default font
                    page.insert_text(pymupdf.Point(x, y), replace_str, fontsize=j["size"], color=j["rgb"])
                total += 1
                # compute new rect for annotations
                try:
                    w = pymupdf.get_text_length(replace_str, fontname=j["fontname"], fontsize=j["size"])
                except Exception:
                    w = rect.width
                new_rect = pymupdf.Rect(x, rect.y0, min(page.rect.x1 - 1, x + w + 1), rect.y1)
                if opts["highlight"]:
                    try:
                        a = page.add_highlight_annot(new_rect)
                        hc = hex_to_rgb(opts["highlight_hex"])
                        a.set_colors(stroke=hc)
                        a.set_opacity(0.45)
                        a.update()
                    except Exception:
                        pass
                if opts["underline"]:
                    try:
                        a = page.add_underline_annot(new_rect)
                        a.set_colors(stroke=j["rgb"])
                        a.update()
                    except Exception:
                        # fallback: draw a line
                        try:
                            p1 = pymupdf.Point(new_rect.x0, new_rect.y1 - 1)
                            p2 = pymupdf.Point(new_rect.x1, new_rect.y1 - 1)
                            page.draw_line(p1, p2, color=j["rgb"], width=0.8)
                        except Exception:
                            pass
        # metadata
        try:
            doc.set_metadata({**doc.metadata, "producer": "Dad PDF Replacer"})
        except Exception:
            pass
        Path(dst).parent.mkdir(parents=True, exist_ok=True)
        doc.save(dst, garbage=4, deflate=True)
        doc.close()
        return total, None
    except Exception as e:
        try:
            doc.close()
        except Exception:
            pass
        return total, f"Error: {e}"


# ---------- UI ----------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Dad's PDF Keyword Replacer")
        self.geometry("680x780")
        self.files = []

        self.var_find = tk.StringVar()
        self.var_replace = tk.StringVar()
        self.var_case = tk.BooleanVar(value=False)
        self.var_whole = tk.BooleanVar(value=True)
        self.var_keep = tk.BooleanVar(value=True)
        self.var_bold = tk.BooleanVar(value=False)
        self.var_italic = tk.BooleanVar(value=False)
        self.var_underline = tk.BooleanVar(value=False)
        self.var_highlight = tk.BooleanVar(value=False)
        self.var_family = tk.StringVar(value="Helvetica")
        self.var_size = tk.IntVar(value=11)
        self.var_textcolor = tk.StringVar(value="#000000")
        self.var_hlcolor = tk.StringVar(value="#FFFF00")
        self.var_out = tk.StringVar(value=str(Path.cwd() / "fixed"))

        self.build()

    def build(self):
        pad = {"padx": 10, "pady": 4}
        tk.Label(self, text="1. Choose PDF files", font=("Arial", 12, "bold")).pack(anchor="w", **pad)
        fr = tk.Frame(self)
        fr.pack(fill="x", **pad)
        tk.Button(fr, text="Add PDFs", font=("Arial", 11), command=self.add_files, width=12).pack(side="left", padx=4)
        tk.Button(fr, text="Add Folder", font=("Arial", 11), command=self.add_folder, width=12).pack(side="left", padx=4)
        tk.Button(fr, text="Remove", font=("Arial", 11), command=self.remove_sel).pack(side="left", padx=4)
        tk.Button(fr, text="Clear", font=("Arial", 11), command=self.clear).pack(side="left", padx=4)

        self.listbox = tk.Listbox(self, height=7, font=("Arial", 10), selectmode="extended")
        self.listbox.pack(fill="x", **pad)
        self.lbl_count = tk.Label(self, text="0 files selected", font=("Arial", 10))
        self.lbl_count.pack(anchor="w", padx=12)

        tk.Label(self, text="2. What to find and replace", font=("Arial", 12, "bold")).pack(anchor="w", **pad)
        gf = tk.Frame(self)
        gf.pack(fill="x", **pad)
        tk.Label(gf, text="Find:", font=("Arial", 11), width=12, anchor="e").grid(row=0, column=0)
        tk.Entry(gf, textvariable=self.var_find, font=("Arial", 11), width=40).grid(row=0, column=1, padx=4)
        tk.Label(gf, text="Replace with:", font=("Arial", 11), width=12, anchor="e").grid(row=1, column=0)
        tk.Entry(gf, textvariable=self.var_replace, font=("Arial", 11), width=40).grid(row=1, column=1, padx=4)
        of = tk.Frame(self)
        of.pack(fill="x", padx=10)
        tk.Checkbutton(of, text="Match case", variable=self.var_case, font=("Arial", 10)).pack(side="left")
        tk.Checkbutton(of, text="Whole word only", variable=self.var_whole, font=("Arial", 10)).pack(side="left", padx=12)

        tk.Label(self, text="3. How should the NEW word look?", font=("Arial", 12, "bold")).pack(anchor="w", **pad)
        sf = tk.Frame(self)
        sf.pack(fill="x", padx=10)
        tk.Checkbutton(sf, text="Keep original font/size/color (recommended)", variable=self.var_keep,
                       font=("Arial", 10), command=self.toggle_custom).pack(anchor="w")
        self.custom_frame = tk.Frame(self)
        self.custom_frame.pack(fill="x", padx=10)
        tk.Label(self.custom_frame, text="Font:", font=("Arial", 10)).grid(row=0, column=0, sticky="e")
        ttk.Combobox(self.custom_frame, textvariable=self.var_family,
                     values=["Helvetica", "Times", "Courier"], width=12, state="readonly").grid(row=0, column=1, padx=4)
        tk.Label(self.custom_frame, text="Size:", font=("Arial", 10)).grid(row=0, column=2, sticky="e")
        tk.Spinbox(self.custom_frame, from_=6, to=72, textvariable=self.var_size, width=5, font=("Arial", 10)).grid(row=0, column=3, padx=4)
        tk.Button(self.custom_frame, text="Text color", command=self.pick_text).grid(row=0, column=4, padx=4)
        self.btn_text_prev = tk.Label(self.custom_frame, text="   ", bg="#000000", width=4)
        self.btn_text_prev.grid(row=0, column=5)

        bf = tk.Frame(self)
        bf.pack(fill="x", padx=10, pady=4)
        tk.Checkbutton(bf, text="Bold", variable=self.var_bold, font=("Arial", 11)).pack(side="left")
        tk.Checkbutton(bf, text="Italic", variable=self.var_italic, font=("Arial", 11)).pack(side="left", padx=8)
        tk.Checkbutton(bf, text="Underline", variable=self.var_underline, font=("Arial", 11)).pack(side="left", padx=8)
        tk.Checkbutton(bf, text="Highlight", variable=self.var_highlight, font=("Arial", 11)).pack(side="left", padx=8)
        tk.Button(bf, text="Highlight color", command=self.pick_hl).pack(side="left", padx=4)
        self.btn_hl_prev = tk.Label(bf, text="   ", bg="#FFFF00", width=4)
        self.btn_hl_prev.pack(side="left")

        tk.Label(self, text="4. Replace", font=("Arial", 12, "bold")).pack(anchor="w", **pad)
        outf = tk.Frame(self)
        outf.pack(fill="x", padx=10)
        tk.Label(outf, text="Save to:", font=("Arial", 10)).pack(side="left")
        tk.Entry(outf, textvariable=self.var_out, font=("Arial", 10), width=42).pack(side="left", padx=4)
        tk.Button(outf, text="Browse", command=self.browse_out).pack(side="left")

        btnf = tk.Frame(self)
        btnf.pack(fill="x", padx=10, pady=8)
        tk.Button(btnf, text="Preview (count only)", font=("Arial", 11), command=self.preview).pack(side="left", padx=4)
        self.btn_go = tk.Button(btnf, text="REPLACE IN ALL FILES", font=("Arial", 12, "bold"),
                                bg="#2e7d32", fg="white", command=self.run)
        self.btn_go.pack(side="left", padx=10)

        self.prog = ttk.Progressbar(self, length=600, mode="determinate")
        self.prog.pack(padx=10, pady=4)
        self.status = tk.Label(self, text="Ready. Pick files to start.", font=("Arial", 10), fg="#333")
        self.status.pack(anchor="w", padx=12)
        self.toggle_custom()

    def toggle_custom(self):
        state = "disabled" if self.var_keep.get() else "normal"
        for child in self.custom_frame.winfo_children():
            try:
                child.configure(state=state)
            except Exception:
                pass

    def pick_text(self):
        c = colorchooser.askcolor(self.var_textcolor.get())[1]
        if c:
            self.var_textcolor.set(c)
            self.btn_text_prev.configure(bg=c)

    def pick_hl(self):
        c = colorchooser.askcolor(self.var_hlcolor.get())[1]
        if c:
            self.var_hlcolor.set(c)
            self.btn_hl_prev.configure(bg=c)
            if not self.var_highlight.get():
                self.var_highlight.set(True)

    def browse_out(self):
        d = filedialog.askdirectory()
        if d:
            self.var_out.set(d)

    def refresh_count(self):
        self.lbl_count.config(text=f"{len(self.files)} files selected")

    def add_files(self):
        ps = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
        for p in ps:
            if p not in self.files:
                self.files.append(p)
                self.listbox.insert("end", p)
        self.refresh_count()

    def add_folder(self):
        d = filedialog.askdirectory()
        if not d:
            return
        for p in sorted(Path(d).glob("*.pdf")):
            s = str(p)
            if s not in self.files:
                self.files.append(s)
                self.listbox.insert("end", s)
        self.refresh_count()

    def remove_sel(self):
        for i in reversed(self.listbox.curselection()):
            self.listbox.delete(i)
            del self.files[i]
        self.refresh_count()

    def clear(self):
        self.files.clear()
        self.listbox.delete(0, "end")
        self.refresh_count()

    def get_opts(self):
        fam_map = {"Helvetica": "helv", "Times": "tiro", "Courier": "cour"}
        return {
            "match_case": self.var_case.get(),
            "whole_word": self.var_whole.get(),
            "keep_original": self.var_keep.get(),
            "bold": self.var_bold.get(),
            "italic": self.var_italic.get(),
            "underline": self.var_underline.get(),
            "highlight": self.var_highlight.get(),
            "highlight_hex": self.var_hlcolor.get(),
            "custom_family": fam_map.get(self.var_family.get(), "helv"),
            "custom_size": int(self.var_size.get()),
            "custom_hex": self.var_textcolor.get(),
        }

    def validate(self):
        if not self.files:
            messagebox.showwarning("No files", "Please add PDF files first.")
            return False
        if not self.var_find.get():
            messagebox.showwarning("Missing keyword", "Please type the keyword to FIND.")
            return False
        if self.var_replace.get() == "":
            if not messagebox.askyesno("Empty replacement", "Replace-with is empty. This will DELETE the keyword. Continue?"):
                return False
        return True

    def preview(self):
        if not self.validate():
            return
        find = self.var_find.get()
        opts = self.get_opts()
        total = 0
        per_file = []
        self.status.config(text="Counting...")
        self.update()
        for f in self.files:
            c = count_matches(f, find, opts["match_case"], opts["whole_word"])
            per_file.append((os.path.basename(f), c))
            total += c
        msg = f"Found {total} matches in {len(self.files)} files.\n\n" + "\n".join(f"{n}: {c}" for n, c in per_file[:20])
        if len(per_file) > 20:
            msg += f"\n...and {len(per_file) - 20} more files"
        messagebox.showinfo("Preview", msg)
        self.status.config(text=f"Preview: {total} matches found.")

    def run(self):
        if not self.validate():
            return
        find = self.var_find.get()
        repl = self.var_replace.get()
        opts = self.get_opts()
        outdir = Path(self.var_out.get())
        outdir.mkdir(parents=True, exist_ok=True)

        self.prog["maximum"] = len(self.files)
        self.prog["value"] = 0
        self.btn_go.config(state="disabled")
        grand = 0
        rows = []
        for i, f in enumerate(self.files):
            self.status.config(text=f"Working {i + 1}/{len(self.files)}: {os.path.basename(f)}")
            self.update()
            dst = outdir / Path(f).name
            # avoid overwriting source if same folder: add _fixed suffix
            try:
                if Path(f).resolve() == dst.resolve():
                    dst = outdir / (Path(f).stem + "_fixed.pdf")
            except Exception:
                pass
            n, warn = replace_in_pdf(f, str(dst), find, repl, opts)
            grand += n
            rows.append([f, str(dst), n, warn or "OK"])
            self.prog["value"] = i + 1
            self.update()

        log = outdir / "replace_log.csv"
        with open(log, "w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            w.writerow(["source", "output", "replacements", "status"])
            w.writerows(rows)
        self.btn_go.config(state="normal")
        self.status.config(text=f"Done! {grand} replacements in {len(self.files)} files. Log: {log}")
        messagebox.showinfo("Done", f"Done!\n{grand} replacements in {len(self.files)} files.\nSaved to: {outdir}\nLog: replace_log.csv")


if __name__ == "__main__":
    App().mainloop()
