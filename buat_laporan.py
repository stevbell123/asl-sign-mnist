# -*- coding: utf-8 -*-
"""
buat_laporan.py: Generator laporan proyek IF2121 Pembelajaran Mendalam sesuai
"Format Laporan Pembelajaran Mendalam" dari dosen:
Cover -> Bagian Awal (Daftar Isi, Daftar Gambar, Daftar Tabel, Pembagian Tugas)
-> Bagian A (Laporan Teknis 1-10) -> Bagian B (Tautan).

Semua angka diambil dari outputs/results.json hasil running notebook.
Pemakaian:  python buat_laporan.py [--results outputs/results.json] [--out file.docx] [--draft]
"""
import argparse, json, os

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

# ============================================================================= argumen & data
ap = argparse.ArgumentParser()
ap.add_argument("--results", default="outputs/results.json")
ap.add_argument("--out", default=None)
ap.add_argument("--draft", action="store_true", help="paksa mode draf (tanpa hasil)")
args, _ = ap.parse_known_args()

R = None
if not args.draft and os.path.exists(args.results):
    with open(args.results) as f:
        R = json.load(f)
DRAFT = R is None
MARK = "[terisi otomatis]"
ISI = "[isi manual]"

META = {
    "nama_kelompok": "Kelompok1", "kelas": "IF-A-PAGI",
    "judul": ("Klasifikasi Huruf Bahasa Isyarat Amerika (ASL) pada Dataset Sign Language MNIST "
              "Menggunakan Deep Neural Network dan Convolutional Neural Network"),
    "mata_kuliah": "IF2121 – Pembelajaran Mendalam", "jenis_laporan": "Ujian Tengah Semester",
    "semester": "Ganjil", "tahun_ajaran": "2026-2027",
    "anggota": [
        {"nama": "Steven David H Malau", "nim": "231111408", "peran": "Ketua Tim & Model Engineer (CNN)"},
        {"nama": "Wan Saka Nasa", "nim": "231111488", "peran": "Data Engineer (EDA, Preprocessing, tf.data)"},
        {"nama": "Joel Saroha Sianturi", "nim": "231111954", "peran": "DNN & Optimasi (Perceptron, MLP, Optimizer)"},
        {"nama": "Nama Anggota 4", "nim": "NIM Anggota 4", "peran": "Evaluasi & Fitur Unik (Threshold/OOD, TFLite, Grad-CAM)"},
        {"nama": "Nama Anggota 5", "nim": "NIM Anggota 5", "peran": "Deployment & Dokumentasi (Gradio, HF Spaces, GitHub, Video)"},
    ],
    "batch_size": 128, "val_size": 0.10, "seed": 42, "links": {"github": "", "video": "", "app": ""},
}
if R:
    META.update({k: v for k, v in R.get("meta", {}).items() if v not in (None, "")})
LINKS = dict(META.get("links") or {})
LOGO = "logo_mikroskil.png"
_links_path = os.path.join(os.path.dirname(args.results) or ".", "links.json")
if os.path.exists(_links_path):
    with open(_links_path) as f:
        _l = json.load(f)
    LINKS.update({k: _l[k] for k in ("github", "video", "app") if _l.get(k)})
    if _l.get("jenis_laporan"):
        META["jenis_laporan"] = _l["jenis_laporan"]
    if _l.get("logo"):
        LOGO = _l["logo"]
QUICK = bool(R and R["meta"].get("quick_mode"))
OUT = args.out or os.path.join("outputs", f"Laporan_{META['nama_kelompok']}_{META['kelas']}_Proyek.docx")
os.makedirs(os.path.dirname(OUT) or ".", exist_ok=True)

E = {e["id"]: e for e in R["experiments"]} if R else {}
FIN = R["final"] if R else {}
D = R["data"] if R else {}
F1 = R.get("fitur1", {}) if R else {}
F2 = R.get("fitur2", {}) if R else {}
DEP = R.get("deploy", {}) if R else {}
FIGS = R.get("figures", {}) if R else {}


# ============================================================================= format angka (gaya Indonesia)
def _id(s):
    return s.replace(",", "#").replace(".", ",").replace("#", ".")

def num(x, d=0):
    return _id(f"{x:,.{d}f}")

def pct(x, d=2):
    return _id(f"{x * 100:,.{d}f}") + "%"

def pp(x, d=2):
    return _id(f"{abs(x) * 100:,.{d}f}") + " poin persentase"

def V(fn):
    """Nilai dinamis; di mode draf (atau jika gagal) kembalikan penanda kuning."""
    if DRAFT:
        return MARK
    try:
        return fn()
    except Exception:
        return MARK

def naik_turun(x):
    return "meningkatkan" if x >= 0 else "menurunkan"

def naik(x):
    return "meningkat" if x >= 0 else "menurun"

def gabung(xs):
    xs = list(xs)
    return xs[0] if len(xs) == 1 else ", ".join(xs[:-1]) + " dan " + xs[-1]


# ============================================================================= dokumen & gaya (mengikuti contoh dosen)
doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21.59), Cm(27.94)          # Letter, seperti format & contoh
for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
    setattr(sec, side, Cm(2.54))
LEBAR = 16.5                                                      # lebar area teks (cm)
FONT = "Times New Roman"


def _set_font(style, size, bold=None, italic=None, underline=None):
    style.font.name, style.font.size = FONT, Pt(size)
    if bold is not None: style.font.bold = bold
    if italic is not None: style.font.italic = italic
    if underline is not None: style.font.underline = underline
    style.font.color.rgb = RGBColor(0, 0, 0)
    rpr = style.element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts"); rpr.append(rf)
    for a in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rf.get(qn(a)) is not None:
            del rf.attrib[qn(a)]
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(a), FONT)


st = doc.styles["Normal"]
_set_font(st, 12)
st.paragraph_format.line_spacing = 1.15
st.paragraph_format.space_after = Pt(6)
for name, size, bold, italic, under, center in (("Heading 1", 14, True, False, False, True),
                                                ("Heading 2", 13, True, False, True, False),
                                                ("Heading 3", 12, True, False, False, False),
                                                ("Heading 4", 12, True, True, False, False)):
    hs = doc.styles[name]
    _set_font(hs, size, bold, italic, under)
    hs.paragraph_format.space_before = Pt(0 if name == "Heading 1" else 10)
    hs.paragraph_format.space_after = Pt(12 if name == "Heading 1" else 4)
    hs.paragraph_format.keep_with_next = True
    hs.paragraph_format.line_spacing = 1.15
    if center:
        hs.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
cap_style = doc.styles["Caption"]
_set_font(cap_style, 10, bold=False, italic=True)
cap_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

_upd = OxmlElement("w:updateFields"); _upd.set(qn("w:val"), "true")   # Word memperbarui field saat dibuka
doc.settings.element.append(_upd)

CNT = {"Gambar": 0, "Tabel": 0}


def _runs(p, text, bold=False, italic=False, size=None):
    tokens = [MARK, ISI]
    parts, buf, i = [], "", 0
    while i < len(text):
        hit = next((t for t in tokens if text.startswith(t, i)), None)
        if hit:
            if buf: parts.append((buf, False))
            parts.append((hit, True)); buf = ""; i += len(hit)
        else:
            buf += text[i]; i += 1
    if buf: parts.append((buf, False))
    for part, hl in parts:
        r = p.add_run(part); r.bold = bold or hl; r.italic = italic
        if size: r.font.size = Pt(size)
        if hl: r.font.highlight_color = WD_COLOR_INDEX.YELLOW
    return p


def P(text, align="justify", bold=False, italic=False, size=None, space_after=6):
    p = doc.add_paragraph()
    p.alignment = {"justify": WD_ALIGN_PARAGRAPH.JUSTIFY, "center": WD_ALIGN_PARAGRAPH.CENTER,
                   "left": WD_ALIGN_PARAGRAPH.LEFT}[align]
    p.paragraph_format.space_after = Pt(space_after)
    return _runs(p, text, bold, italic, size)


def RICH(segments):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    for t, s in segments:
        _runs(p, t, bold="b" in s, italic="i" in s)
    return p


def LIST(items, style="num", level=0):
    """style: 'num' (1. 2.), 'alpha' (a. b.), 'bullet'."""
    for i, it in enumerate(items, 1):
        p = doc.add_paragraph()
        ind = 0.75 + 0.75 * level
        p.paragraph_format.left_indent = Cm(ind); p.paragraph_format.first_line_indent = Cm(-0.6)
        p.paragraph_format.space_after = Pt(2); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.tab_stops.add_tab_stop(Cm(ind))
        p.add_run({"num": f"{i}.", "alpha": f"{chr(96 + i)}.", "bullet": "•"}[style] + "\t")
        if isinstance(it, tuple):
            _runs(p, it[0], bold=True); _runs(p, it[1])
        else:
            _runs(p, it)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def H1(t, page_break=True):
    h = doc.add_heading(t, level=1)
    h.paragraph_format.page_break_before = page_break
def H2(t): doc.add_heading(t, level=2)
def H3(t): doc.add_heading(t, level=3)
def H4(t): doc.add_heading(t, level=4)


def _shade(cell, hex_color):
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement("w:shd"); shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def _field(p, instr, cached):
    def fc(kind):
        r = p.add_run(); el = OxmlElement("w:fldChar"); el.set(qn("w:fldCharType"), kind); r._r.append(el)
    fc("begin")
    r = p.add_run(); it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = f" {instr} "; r._r.append(it)
    fc("separate")
    p.add_run(cached)
    fc("end")


def CAP(kind, text):
    CNT[kind] += 1
    p = doc.add_paragraph(style="Caption")
    p.paragraph_format.keep_with_next = kind == "Tabel"
    p.paragraph_format.space_after = Pt(10 if kind == "Gambar" else 3)
    p.add_run(f"{kind} ")
    _field(p, f"SEQ {kind} \\* ARABIC", str(CNT[kind]))
    p.add_run(" - ")
    _runs(p, text, italic=True)
    return f"{kind} {CNT[kind]}"


def REF(kind, offset=1):
    """Nomor gambar/tabel berikutnya (offset=1) atau sesudahnya (offset=2, dst.)."""
    return f"{kind} {CNT[kind] + offset}"


def set_widths(t, widths):
    t.autofit = False
    tblPr = t._tbl.tblPr
    lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); tblPr.append(lay)
    for i, gc in enumerate(t._tbl.tblGrid.findall(qn("w:gridCol"))):
        if i < len(widths):
            gc.set(qn("w:w"), str(int(widths[i] * 567)))
    for row in t.rows:
        for i, w in enumerate(widths):
            if i < len(row.cells):
                row.cells[i].width = Cm(w)


def TABLE(headers, rows, caption, widths=None, size=10, center_cols=None):
    CAP("Tabel", caption)
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]; c.text = ""
        r = c.paragraphs[0].add_run(str(h)); r.bold = True; r.font.size = Pt(size)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        _shade(c, "D9E2F3")
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            lines = v if isinstance(v, list) else [v]
            for k, line in enumerate(lines):
                p = cells[i].paragraphs[0] if k == 0 else cells[i].add_paragraph()
                s = ("• " if isinstance(v, list) else "") + str(line)
                auto = i > 0 and (s[:1].isdigit() or s[:1] in "-−–✓" or s.endswith("%") or len(s) <= 3)
                if (center_cols is not None and i in center_cols) or (center_cols is None and auto):
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                _runs(p, s, size=size)
    for row in t.rows:
        for c in row.cells:
            for p in c.paragraphs:
                p.paragraph_format.line_spacing = 1.0; p.paragraph_format.space_after = Pt(1)
    set_widths(t, widths or [LEBAR / len(headers)] * len(headers))
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


GEN_FIGS = ("fig_arsitektur_cnn",)


def BOX(text, width=15.5, pad_lines=1):
    t = doc.add_table(rows=1, cols=1); t.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_widths(t, (width,)); c = t.rows[0].cells[0]; _shade(c, "F2F2F2")
    q = c.paragraphs[0]; q.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = q.add_run("\n" * pad_lines + text + "\n" * pad_lines)
    r.italic = True; r.font.size = Pt(10); r.font.highlight_color = WD_COLOR_INDEX.YELLOW
    k = doc.add_paragraph(); k.paragraph_format.space_after = Pt(0); k.paragraph_format.keep_with_next = True


def FIG(key, caption, width=15.5, path=None):
    path = path or FIGS.get(key) or os.path.join("outputs", "figures", f"{key}.png")
    if os.path.exists(path) and (not DRAFT or key in GEN_FIGS or key.startswith("ss_")):
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.keep_with_next = True; p.paragraph_format.space_after = Pt(2)
        p.add_run().add_picture(path, width=Cm(width))
    else:
        BOX(f"[Gambar \"{key}.png\" disisipkan otomatis setelah notebook dijalankan]", width)
    return CAP("Gambar", caption)


def SCREENSHOT(no, caption):
    """Sisipkan screenshots/langkah{no}.png jika ada; jika tidak, kotak penanda untuk diisi manual."""
    for ext in ("png", "jpg", "jpeg"):
        path = os.path.join("screenshots", f"langkah{no}.{ext}")
        if os.path.exists(path):
            return FIG(f"ss_{no}", caption, width=14, path=path)
    BOX(f"{ISI} Tempel tangkapan layar di sini (atau simpan sebagai screenshots/langkah{no}.png lalu jalankan ulang generator)", 14, 2)
    return CAP("Gambar", caption)


def CODE(text, size=8.5):
    t = doc.add_table(rows=1, cols=1); t.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_widths(t, (LEBAR,)); c = t.rows[0].cells[0]; _shade(c, "F5F5F5")
    p = c.paragraphs[0]; p.paragraph_format.line_spacing = 1.0; p.paragraph_format.space_after = Pt(0)
    r = p.add_run(text.strip("\n")); r.font.name = "Consolas"; r.font.size = Pt(size)
    r._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:hAnsi"), "Consolas")
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def page_number_footer(section):
    p = section.footer.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _field(p, "PAGE", "1")


def TOC(instr, hint):
    p = doc.add_paragraph()
    _field(p, instr, hint)


def link_or_isi(key, label):
    v = LINKS.get(key) or ""
    return v if v else f"{ISI} {label}"


# ============================================================================= gambar buatan generator
def buat_diagram_arsitektur():
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        from matplotlib.patches import FancyBboxPatch
    except Exception:
        return
    os.makedirs(os.path.join("outputs", "figures"), exist_ok=True)
    blocks = [("Input\n28×28×1", "#eeeeee"),
              ("Conv3×3-32\nBN-ReLU ×2\nMaxPool, Drop .25\n→14×14×32", "#dbe7f5"),
              ("Conv3×3-64\nBN-ReLU ×2\nMaxPool, Drop .25\n→7×7×64", "#c6d9f0"),
              ("Conv3×3-128\nBN-ReLU\nMaxPool, Drop .30\n→3×3×128", "#a9c4e8"),
              ("Flatten\n1152", "#eeeeee"),
              ("Dense 256\nBN-ReLU\nDrop .40", "#fde7d4"),
              ("Dense 24\nSoftmax", "#f9cfa8")]
    fig, ax = plt.subplots(figsize=(13, 2.9)); ax.axis("off"); ax.set_xlim(0, len(blocks)); ax.set_ylim(0, 1.05)
    for i, (t, c) in enumerate(blocks):
        ax.add_patch(FancyBboxPatch((i + 0.07, 0.08), 0.86, 0.78, boxstyle="round,pad=0.02", fc=c, ec="#333", lw=1.2))
        ax.text(i + 0.5, 0.47, t, ha="center", va="center", fontsize=9.5)
        if i < len(blocks) - 1:
            ax.annotate("", xy=(i + 1.07, 0.47), xytext=(i + 0.93, 0.47), arrowprops=dict(arrowstyle="-|>", lw=1.4))
    ax.text(0.07, 1.0, "Augmentasi (rotasi, zoom, translasi, brightness, contrast) diterapkan pada pipeline tf.data, hanya saat latih",
            fontsize=9, ha="left", va="top", style="italic")
    plt.tight_layout(); fig.savefig(os.path.join("outputs", "figures", "fig_arsitektur_cnn.png"), dpi=170); plt.close(fig)


buat_diagram_arsitektur()


# ============================================================================= COVER
def cover():
    def line(txt, size=12, bold=True, before=0, after=0):
        q = doc.add_paragraph(); q.alignment = WD_ALIGN_PARAGRAPH.CENTER
        q.paragraph_format.space_before = Pt(before); q.paragraph_format.space_after = Pt(after)
        q.paragraph_format.line_spacing = 1.3
        _runs(q, txt, bold=bold, size=size)
    line(f"LAPORAN PROYEK {META['jenis_laporan'].upper()}", 14, before=24)
    line(META["judul"].upper(), 14, after=18)
    line(f"Mata Kuliah: {META['mata_kuliah']}", 12, bold=False)
    line(f"Semester: {META['semester']}", 12, bold=False)
    line(f"TA. {META['tahun_ajaran']}", 12, bold=False, after=18)
    line(f"Nama Kelompok: {META['nama_kelompok']}", 12)
    line("NIM & Nama Anggota Kelompok:", 12)
    for a in META["anggota"]:
        line(f"{a['nim']} – {a['nama']}", 12, bold=False)
    q = doc.add_paragraph(); q.alignment = WD_ALIGN_PARAGRAPH.CENTER
    q.paragraph_format.space_before = Pt(18); q.paragraph_format.space_after = Pt(18)
    if os.path.exists(LOGO):
        q.add_run().add_picture(LOGO, width=Cm(4.5))
    else:
        _runs(q, f"{ISI} logo Universitas Mikroskil (unggah logo_mikroskil.png lalu jalankan ulang)", size=10)
    if DRAFT or QUICK:
        line("DRAF: bagian bertanda kuning terisi otomatis setelah notebook dijalankan" if DRAFT else
             "HASIL QUICK MODE: jalankan ulang dengan QUICK_MODE = False sebelum dikumpulkan", 10, after=6)
    line("PROGRAM STUDI S-1 TEKNIK INFORMATIKA", 12)
    line("FAKULTAS INFORMATIKA", 12)
    line("UNIVERSITAS MIKROSKIL", 12)


cover()
page_number_footer(doc.sections[0])

# ============================================================================= BAGIAN AWAL
H1("DAFTAR ISI")
TOC('TOC \\o "1-3" \\h \\z \\u', "Klik kanan di sini → Update Field (atau pilih Yes saat Word menanyakan pembaruan field).")
H1("DAFTAR GAMBAR")
TOC('TOC \\h \\z \\c "Gambar"', "Klik kanan di sini → Update Field untuk menampilkan daftar gambar.")
H1("DAFTAR TABEL")
TOC('TOC \\h \\z \\c "Tabel"', "Klik kanan di sini → Update Field untuk menampilkan daftar tabel.")

H1("PEMBAGIAN TUGAS")
P(f"{REF('Tabel')} merinci peran dan pekerjaan setiap anggota. Setiap anggota memegang satu tahapan utama proyek "
  "(data, model, optimasi, evaluasi dan fitur unik, serta deployment) dengan beban yang seimbang (±20% per anggota). "
  "Setiap anggota juga menjelaskan bagiannya sendiri dalam video demo.")
RINCIAN = [
    ["Koordinasi tim dan konsultasi topik ke dosen", "Merancang arsitektur CNN (E3–E5): konvolusi, pooling, BatchNorm, Dropout",
     "Melatih dan melakukan tuning model CNN akhir", "Integrasi notebook akhir dan review laporan"],
    ["Akuisisi dataset (KaggleHub/mirror) dan pengecekan integritas", "EDA: distribusi kelas, histogram piksel, korelasi piksel, contoh per kelas",
     "Pembersihan, normalisasi, reshape, dan pembagian data tanpa kebocoran", "Representasi tensor, pipeline tf.data, dan augmentasi data"],
    ["Perceptron Rosenblatt, neuron sigmoid, dan MLP manual (GradientTape)", "Model DNN/MLP E1–E2",
     "Studi fungsi aktivasi dan optimizer (SGD, Momentum, RMSProp, Adam)", "Callback: EarlyStopping, ReduceLROnPlateau, Checkpoint, TensorBoard"],
    ["Evaluasi: metrik, classification report, confusion matrix, analisis error", "Fitur unik 1: confidence threshold & deteksi input di luar domain",
     "Fitur unik 2: kompresi model ke TensorFlow Lite", "Interpretasi Grad-CAM"],
    ["Skrip inferensi dan antarmuka Gradio", "Deployment ke Hugging Face Spaces dan pengujian aplikasi",
     "Repositori GitHub dan README", "Video demo, penyusunan laporan, dan pengumpulan"],
]
TABLE(["NIM", "Nama", "Peran", "Rincian pekerjaan"],
      [[a["nim"], a["nama"], a["peran"], RINCIAN[i] if i < len(RINCIAN) else ["Mendukung seluruh tahapan proyek"]]
       for i, a in enumerate(META["anggota"])],
      "Pembagian tugas anggota kelompok", widths=[2.4, 3.3, 3.6, 7.2], size=9, center_cols=[0])

# ============================================================================= BAGIAN A
H1("BAGIAN A. LAPORAN TEKNIS PROYEK")

# ---------------------------------------------------------------- 1. Kompleksitas masalah
H2("1. Kompleksitas Masalah")
P("Proyek ini membangun model yang mengenali 24 huruf statis American Sign Language (ASL) dari citra tangan "
  "grayscale berukuran 28×28 piksel. Pengenalan huruf isyarat (fingerspelling) adalah fondasi sistem penerjemah bahasa "
  "isyarat untuk membantu komunikasi penyandang tunarungu. Menurut WHO, lebih dari 430 juta orang hidup dengan gangguan "
  "pendengaran yang memerlukan rehabilitasi (World Health Organization, 2021). Model dibatasi pada Deep Neural Network "
  "(DNN/MLP) dan Convolutional Neural Network (CNN), sesuai ketentuan tugas.")
H3("a. Mengapa masalah ini dianggap kompleks?")
LIST([("Kemiripan antarkelas sangat tinggi (fine-grained). ", "Huruf berbasis kepalan (A, E, M, N, S, T) hanya berbeda pada "
       "posisi ibu jari, dan huruf dua jari (U, V, R, K) hanya berbeda pada jarak atau silangan jari."),
      ("Resolusi rendah dan dimensi tinggi. ", "Citra hanya 28×28 piksel grayscale sehingga banyak detail jari yang "
       "hilang. Bagi DNN, setiap citra adalah vektor 784 dimensi tanpa informasi spasial."),
      ("Variasi kondisi citra. ", "Pencahayaan, posisi, rotasi, skala, latar, dan bentuk tangan setiap orang berbeda-beda."),
      ("Pergeseran distribusi data latih dan uji. ", "Data latih dibentuk dengan memperbanyak sejumlah kecil foto asli, "
       "sehingga validasi dari data latih cenderung lebih optimistis daripada data uji."),
      ("Ketidakseimbangan kelas pada data uji. ", f"Rasio kelas terbanyak terhadap tersedikit pada data uji mencapai "
       f"{V(lambda: num(D['imbalance_ratio_test'], 2))}."),
      ("Input dunia nyata yang tidak terduga. ", "Pada aplikasi, pengguna dapat mengunggah foto apa saja, padahal "
       "softmax selalu memaksakan salah satu dari 24 huruf.")])
H3("b. Apa tantangan utama dalam menyelesaikan masalah ini?")
LIST(["Membuat model yang menggeneralisasi ke data uji berdistribusi berbeda tanpa overfitting pada data latih yang mirip satu sama lain.",
      "Membedakan pasangan huruf yang secara visual hampir identik pada resolusi 28×28.",
      "Mencegah kebocoran data: data uji tidak boleh dipakai untuk early stopping, pemilihan model, atau penentuan ambang.",
      "Menolak secara andal input yang meragukan atau berada di luar domain, alih-alih memberi jawaban salah dengan yakin.",
      "Menghasilkan model yang ringan dan cepat agar dapat di-deploy sebagai aplikasi web."])
H3("c. Pengembangan dan penyempurnaan dari proyek sebelumnya")
P(f"{REF('Tabel')} merangkum bagian yang ditambah, diperbaiki, atau diganti berdasarkan audit terhadap notebook dan "
  "laporan proyek sebelumnya.")
TABLE(["Aspek", "Proyek sebelumnya", "Proyek ini", "Jenis"], [
    ["Jumlah data", "Subset 1.200 citra (50/kelas)", "Seluruh data: 27.455 latih + 7.172 uji resmi", "Diperbaiki"],
    ["Pembagian data", "Data uji juga dipakai sebagai validasi (kebocoran data)", "Latih/validasi stratified 90/10; data uji resmi terpisah", "Diperbaiki"],
    ["Cakupan metode", "Campuran ML klasik (DT, KNN, LR, K-Means)", "Dibatasi pada DNN dan CNN sesuai ketentuan", "Diganti"],
    ["Input pipeline", "Array NumPy langsung", "tf.data: cache, shuffle, augmentasi, batch, prefetch", "Ditambah"],
    ["DNN & optimasi", "Tidak ada MLP dan studi optimizer", "Perceptron, MLP manual & Keras, studi aktivasi & optimizer, LR scheduling", "Ditambah"],
    ["Augmentasi", "Tidak ada", "Rotasi, zoom, translasi, brightness, contrast", "Ditambah"],
    ["Fitur unik", "Tidak ada", "Confidence threshold + deteksi OOD; kompresi TFLite", "Ditambah"],
    ["Interpretasi", "Tidak ada", "Grad-CAM", "Ditambah"],
    ["Deployment", "Tidak ada", "Aplikasi Gradio di Hugging Face Spaces + repositori GitHub", "Ditambah"],
    ["Hasil CNN", "71,67% pada 240 citra (laporan menulis 93,33%)", f"{V(lambda: pct(FIN['acc']))} pada 7.172 citra uji resmi", "Diperbaiki"],
    ["Konsistensi laporan", "Placeholder [ISI] dan angka tidak cocok dengan output", "Laporan disusun otomatis dari results.json", "Diperbaiki"],
], "Pengembangan dan penyempurnaan dari proyek sebelumnya", widths=[2.9, 4.6, 6.6, 2.4], size=9, center_cols=[3])

# ---------------------------------------------------------------- 2. Dataset
H2("2. Dataset yang Digunakan")
H3("a. Sumber dataset")
P("Dataset yang digunakan adalah Sign Language MNIST (Datamunge, 2017) dari Kaggle: "
  "https://www.kaggle.com/datasets/datamunge/sign-language-mnist. Dataset ini dirancang sebagai pengganti MNIST "
  "(LeCun et al., 1998) untuk pengenalan gestur tangan. Notebook mengunduhnya otomatis melalui KaggleHub; jika gagal, "
  "notebook memakai salinan CSV yang identik di GitHub atau meminta unggah manual"
  + ("" if DRAFT else f" (sumber pada running ini: {D.get('sumber', '-')})") + ".")
H3("b. Jumlah data, kelas/label, format, dan ukuran data")
TABLE(["Atribut", "Keterangan"], [
    ["Jumlah data", V(lambda: f"{num(D['n_train_raw'] + D['n_test_raw'])} citra: {num(D['n_train_raw'])} latih + {num(D['n_test_raw'])} uji resmi")],
    ["Kelas/label", "24 kelas: huruf A–Y tanpa J dan Z (keduanya memerlukan gerakan); label asli 0–24 tanpa 9"],
    ["Format", "CSV; setiap baris = 1 kolom label + 784 kolom piksel (row-major)"],
    ["Ukuran citra", "28×28 piksel, grayscale 1 kanal, intensitas 0–255"],
    ["Data latih per kelas", V(lambda: f"{num(min(D['train_counts'].values()))}–{num(max(D['train_counts'].values()))} citra")],
    ["Data uji per kelas", V(lambda: f"{num(min(D['test_counts'].values()))}–{num(max(D['test_counts'].values()))} citra")],
], "Ringkasan dataset Sign Language MNIST", widths=[4.0, 12.5], center_cols=[])
H3("c. Apakah dataset bersifat real-world atau sintetis?")
P("Dataset ini bersifat real-world yang diperbanyak secara sintetis. Citra dasarnya adalah foto tangan sungguhan dari "
  "beberapa pengguna dengan latar berbeda. Menurut deskripsi di Kaggle, foto tersebut dipotong, diubah ke grayscale, "
  "diperkecil, lalu diperbanyak menjadi banyak variasi melalui pipeline pengolahan citra (filter, pikselasi acak, "
  "perubahan kecerahan/kontras, dan rotasi kecil). Akibatnya, banyak citra latih merupakan variasi dari foto asli yang "
  "sama. Sifat ini membuat validasi dari data latih menjadi optimistis.")
H3("d. Seberapa kompleks dataset tersebut?")
LIST([f"Berukuran sedang ({V(lambda: num(D['n_train_raw'] + D['n_test_raw']))} citra) dengan 24 kelas dan 784 fitur per citra.",
      f"Data latih relatif seimbang (rasio {V(lambda: num(D['imbalance_ratio_train'], 2))}), tetapi data uji timpang (rasio {V(lambda: num(D['imbalance_ratio_test'], 2))}).",
      "Kelas bersifat fine-grained: beberapa pasangan huruf hanya berbeda pada detail kecil.",
      "Terdapat pergeseran distribusi antara data latih dan uji, sehingga tantangan utamanya adalah generalisasi, bukan sekadar menghafal data latih."])

# ---------------------------------------------------------------- 3. EDA
H2("3. Exploratory Data Analysis (EDA)")
H3("a. Distribusi label/kelas")
if not DRAFT:
    ir_tr, ir_te = D["imbalance_ratio_train"], D["imbalance_ratio_test"]
    tc = D["test_counts"]
    kal = (f"Data latih relatif seimbang dengan rasio kelas terbanyak terhadap tersedikit {num(ir_tr, 2)}. Data uji "
           f"{'cukup timpang' if ir_te > 2 else 'juga relatif seimbang'} dengan rasio {num(ir_te, 2)}: kelas terbanyak "
           f"{max(tc, key=tc.get)} ({num(max(tc.values()))} citra) dan tersedikit {min(tc, key=tc.get)} ({num(min(tc.values()))} citra).")
else:
    kal = f"Rasio kelas terbanyak/tersedikit pada data latih dan uji {MARK}."
P(f"{REF('Gambar')} memperlihatkan distribusi jumlah citra per huruf. {kal} Karena itu, macro-F1 dilaporkan di samping "
  "accuracy agar setiap huruf berbobot sama.")
FIG("fig_distribusi_kelas", "Distribusi kelas pada data latih dan uji resmi")
H3("b. Word count")
P("Tidak relevan, karena dataset berupa citra dan tidak mengandung teks.")
H3("c. Histogram fitur numerik")
P(f"Untuk data citra, fitur numeriknya adalah intensitas piksel. {REF('Gambar')} membandingkan histogram intensitas "
  "piksel dan kecerahan rata-rata per citra antara data latih dan uji. Rata-rata ± simpangan baku piksel adalah "
  f"{V(lambda: num(D['pixel_mean_train'], 1) + ' ± ' + num(D['pixel_std_train'], 1))} (latih) dan "
  f"{V(lambda: num(D['pixel_mean_test'], 1) + ' ± ' + num(D['pixel_std_test'], 1))} (uji). Sebaran kecerahan antarcitra "
  "cukup lebar, sehingga augmentasi kecerahan dan kontras diperlukan.")
FIG("fig_histogram_intensitas", "Histogram intensitas piksel dan kecerahan rata-rata citra latih vs uji")
H3("d. Korelasi antarfitur")
P(f"{REF('Gambar')} menunjukkan korelasi Pearson antarpiksel pada baris dan kolom tengah. Piksel bertetangga berkorelasi "
  f"sangat kuat (rata-rata r = {V(lambda: num(D['corr_tetangga'], 3))}), sedangkan piksel berjarak 10 jauh lebih lemah "
  f"(r = {V(lambda: num(D['corr_jarak10'], 3))}). Struktur korelasi lokal inilah alasan utama penggunaan CNN: konvolusi "
  "memanfaatkan keterhubungan lokal, sedangkan DNN memperlakukan setiap piksel sebagai fitur independen.")
FIG("fig_korelasi_piksel", "Korelasi antarpiksel pada baris ke-14 dan kolom ke-14", width=14)
H3("e. Contoh sampel data per kelas")
P(f"{REF('Gambar')} menampilkan satu contoh citra untuk setiap huruf, dan {REF('Gambar', 2)} menampilkan citra rata-rata "
  "per kelas yang memperlihatkan pola bentuk tangan khas tiap huruf.")
FIG("fig_contoh_per_kelas", "Contoh citra untuk setiap huruf")
FIG("fig_rata2_per_kelas", "Citra rata-rata setiap kelas")
H3("f. Temuan penting dari EDA yang memengaruhi preprocessing")
dup_txt = V(lambda: ("tidak ditemukan duplikat persis maupun citra uji yang identik dengan citra latih"
                     if D["dup_train"] == 0 and D["overlap_train_test"] == 0 else
                     f"ditemukan {num(D['dup_train'])} duplikat persis dan {num(D['overlap_train_test'])} citra uji yang identik dengan citra latih"))
TABLE(["Temuan EDA", "Keputusan preprocessing/pemodelan"], [
    [f"Tidak ada missing value dan semua piksel dalam rentang 0–255; {dup_txt}", "Tidak ada baris yang dihapus; cukup validasi rentang dan pemetaan label"],
    ["Data uji timpang", "Macro-F1 dipakai sebagai metrik pendamping accuracy"],
    ["Sebaran kecerahan dan posisi tangan bervariasi", "Augmentasi brightness, contrast, rotasi, zoom, dan translasi pada tf.data"],
    ["Korelasi piksel sangat lokal", "Arsitektur CNN dengan konvolusi 3×3 dan pooling"],
    ["Data latih berupa variasi dari foto asli yang sama", "Validasi tidak dipakai sebagai satu-satunya ukuran; data uji resmi dipisahkan sepenuhnya"],
    ["Label asli 0–24 tanpa 9", "Re-indeks label ke 0–23"],
], "Temuan EDA dan dampaknya pada preprocessing", widths=[7.5, 9.0], size=9.5, center_cols=[])

# ---------------------------------------------------------------- 4. Preprocessing
H2("4. Preprocessing Data")
H3("a. Pembersihan data")
LIST([f"Pemeriksaan missing value: {V(lambda: str(D['missing_train']))} (latih) dan {V(lambda: str(D['missing_test']))} (uji), sehingga tidak ada imputasi.",
      f"Pemeriksaan duplikat dengan hash baris piksel: {V(lambda: num(D['dup_train']))} duplikat di data latih dan {V(lambda: num(D['overlap_train_test']))} citra uji yang identik dengan data latih.",
      "Validasi data rusak: semua nilai piksel berada pada rentang 0–255 dan setiap baris berukuran 784 piksel.",
      "Validasi label: hanya 24 nilai label yang sah (0–24 tanpa 9), lalu dipetakan ulang ke indeks 0–23 agar sesuai dengan 24 neuron softmax."])
H3("b. Normalisasi dan rescale piksel")
P("Piksel diubah ke float32 lalu dibagi 255 (x' = x/255) sehingga berada pada rentang [0, 1]; skala input yang seragam "
  "menstabilkan gradien. Konstanta 255 tidak dihitung dari data, sehingga tidak ada kebocoran statistik. Vektor 784 "
  "diubah menjadi tensor (28, 28, 1), dan satu batch menjadi tensor rank-4 (B, 28, 28, 1). Resize tidak diperlukan karena "
  "dataset sudah 28×28. Pada aplikasi, foto pengguna diubah ke grayscale, dipotong persegi di tengah, di-resize ke 28×28 "
  "(interpolasi area/box), lalu dibagi 255. Data dialirkan dengan pipeline tf.data berikut.")
CODE("""
ds = tf.data.Dataset.from_tensor_slices((X, y)).cache()
ds = ds.shuffle(len(X), seed=SEED, reshuffle_each_iteration=True)   # hanya data latih
ds = ds.map(photometric_augment, num_parallel_calls=AUTOTUNE)       # hanya E5
ds = ds.batch(128)
ds = ds.map(lambda x, y: (geo_augment(x, training=True), y), num_parallel_calls=AUTOTUNE)   # hanya E5
ds = ds.prefetch(AUTOTUNE)
""")
P(f"Throughput pipeline: {V(lambda: num(R['pipeline']['throughput_naive']))} citra/detik tanpa cache/prefetch, "
  f"{V(lambda: num(R['pipeline']['throughput_plain']))} citra/detik dengan cache+shuffle+prefetch, dan "
  f"{V(lambda: num(R['pipeline']['throughput_aug']))} citra/detik dengan augmentasi.")
H3("c. Augmentasi data")
P(f"Augmentasi hanya diterapkan pada data latih skenario E5, dengan parameter pada {REF('Tabel')}. Contoh hasilnya "
  f"ditunjukkan pada {REF('Gambar')}.")
TABLE(["Transformasi", "Parameter", "Alasan"], [
    ["RandomRotation", "faktor 0,03 (±10,8°)", "Variasi kemiringan tangan"],
    ["RandomZoom", "±10%", "Variasi jarak tangan ke kamera"],
    ["RandomTranslation", "±10% tinggi & lebar", "Variasi posisi tangan di bingkai"],
    ["random_brightness", "Δ ≤ 0,10", "Variasi pencahayaan"],
    ["random_contrast", "faktor 0,8–1,2", "Variasi kontras kamera"],
    ["Flip horizontal", "Tidak digunakan", "Membalik citra mengubah tangan kanan menjadi kiri"],
], "Jenis dan parameter augmentasi data", widths=[4.0, 4.5, 8.0], center_cols=[1])
FIG("fig_augmentasi", "Citra asli (baris pertama) dan tiga variasi hasil augmentasi", width=14.5)
H3("d. Penanganan missing value, outlier, dan ketidakseimbangan kelas")
LIST(["Missing value: tidak ada, sehingga tidak diperlukan imputasi.",
      "Outlier: tidak ada nilai piksel di luar rentang. Citra dengan pencahayaan ekstrem tidak dibuang, tetapi diantisipasi dengan augmentasi brightness/contrast agar model lebih tahan.",
      "Ketidakseimbangan kelas: data latih sudah relatif seimbang sehingga tidak diperlukan oversampling atau class weight; ketimpangan data uji ditangani dengan pelaporan macro-F1 dan metrik per kelas."])
H3("e. Pembagian data dan pencegahan kebocoran data")
TABLE(["Subset", "Sumber", "Jumlah", "Proporsi", "Kegunaan"], [
    ["Latih", "Data latih resmi (stratified)", V(lambda: num(D["n_train"])), "90% data latih", "Memperbarui bobot"],
    ["Validasi", "Data latih resmi (stratified)", V(lambda: num(D["n_val"])), "10% data latih", "Early stopping, scheduler, checkpoint, ambang τ"],
    ["Uji", "Data uji resmi", V(lambda: num(D["n_test"])), "Terpisah", "Evaluasi akhir saja"],
], "Pembagian data", widths=[2.0, 4.6, 2.2, 2.9, 4.8], size=9.5, center_cols=[2, 3])
LIST(["Data uji resmi tidak pernah dipakai untuk pelatihan, early stopping, pemilihan hyperparameter, maupun penentuan ambang τ.",
      f"Pembagian latih/validasi dilakukan secara stratified dengan random_state={META['seed']}.",
      "Ambang τ (fitur unik 1) ditentukan dari data validasi yang diaugmentasi, bukan dari data uji.",
      "Model akhir (E5) ditetapkan sejak awal berdasarkan rancangan, bukan dipilih dengan melihat skor uji.",
      "Normalisasi memakai konstanta tetap (/255), bukan statistik yang dihitung dari seluruh data."])

# ---------------------------------------------------------------- 5. Model
H2("5. Model yang Digunakan")
H3("a. Jenis model")
P("Dua keluarga model digunakan sesuai ketentuan. DNN (Multi-Layer Perceptron) menjadi baseline E1–E2, dan CNN "
  "menjadi model utama E3–E5. Model akhir adalah CNN tiga blok konvolusi yang dilatih dari awal (E5). Transfer learning "
  "(ResNet, VGG, MobileNet, EfficientNet) tidak dipakai karena tiga alasan: input berupa citra grayscale 28×28 yang "
  "sangat berbeda dari input model pretrained (RGB ±224×224); CNN kecil sudah cukup untuk mempelajari pola sederhana "
  "ini; dan model kecil (±442 ribu parameter) jauh lebih ringan untuk deployment.")
H3("b. Arsitektur model")
P(f"Diagram arsitektur CNN akhir ditunjukkan pada {REF('Gambar')}, dan rincian setiap lapisannya pada {REF('Tabel')}. "
  "Setiap konvolusi 3×3 (padding same) diikuti BatchNormalization dan ReLU. Setiap blok diakhiri MaxPooling 2×2 dan "
  "Dropout. Lapisan sebelum BatchNorm tidak memakai bias karena perannya sudah digantikan oleh parameter β. Output "
  "berupa softmax 24 kelas.")
FIG("fig_arsitektur_cnn", "Diagram arsitektur CNN akhir (E5)", width=16)
if not DRAFT and "E5" in E:
    rows = [[r[0], r[1], r[2].replace("None", "B"), num(r[3])] for r in E["E5"]["arsitektur"]]
    total = num(E["E5"]["params"])
else:
    arch = [("b1c1_conv", "Conv2D 3×3, 32", "(B, 28, 28, 32)", 288), ("b1c1_bn / relu", "BatchNorm / ReLU", "(B, 28, 28, 32)", 128),
            ("b1c2_conv", "Conv2D 3×3, 32", "(B, 28, 28, 32)", 9216), ("b1c2_bn / relu", "BatchNorm / ReLU", "(B, 28, 28, 32)", 128),
            ("b1_pool / b1_drop", "MaxPool 2×2 / Dropout 0,25", "(B, 14, 14, 32)", 0),
            ("b2c1_conv", "Conv2D 3×3, 64", "(B, 14, 14, 64)", 18432), ("b2c1_bn / relu", "BatchNorm / ReLU", "(B, 14, 14, 64)", 256),
            ("b2c2_conv", "Conv2D 3×3, 64", "(B, 14, 14, 64)", 36864), ("b2c2_bn / relu", "BatchNorm / ReLU", "(B, 14, 14, 64)", 256),
            ("b2_pool / b2_drop", "MaxPool 2×2 / Dropout 0,25", "(B, 7, 7, 64)", 0),
            ("b3c1_conv", "Conv2D 3×3, 128", "(B, 7, 7, 128)", 73728), ("b3c1_bn / relu", "BatchNorm / ReLU", "(B, 7, 7, 128)", 512),
            ("b3_pool / b3_drop", "MaxPool 2×2 / Dropout 0,30", "(B, 3, 3, 128)", 0),
            ("flatten", "Flatten", "(B, 1152)", 0), ("fc_dense", "Dense 256", "(B, 256)", 294912),
            ("fc_bn / relu / drop", "BatchNorm / ReLU / Dropout 0,40", "(B, 256)", 1024),
            ("output_softmax", "Dense 24, softmax", "(B, 24)", 6168)]
    rows = [[a, b, c, num(d)] for a, b, c, d in arch]
    total = num(441912)
rows.append(["Total", "", "", total])
TABLE(["Lapisan", "Tipe", "Ukuran output", "Parameter"], rows,
      "Rincian lapisan CNN akhir (B = ukuran batch)", widths=[4.2, 5.4, 4.0, 2.9], size=8.5, center_cols=[2, 3])
P(f"Sebagai pembanding, {REF('Tabel')} merangkum arsitektur kelima skenario.")
def _par(k, default):
    return default if DRAFT else V(lambda: num(E[k]["params"]))
TABLE(["ID", "Jenis", "Arsitektur ringkas", "Parameter"], [
    ["E1", "DNN", "Flatten → Dense 256 (ReLU) → Dense 128 (ReLU) → Softmax 24", _par("E1", "236.952")],
    ["E2", "DNN", "Flatten → [Dense 512/256/128 + BN + ReLU + Dropout 0,3] → Softmax 24", _par("E2", "571.928")],
    ["E3", "CNN", "Conv32 → Pool → Conv64 → Pool → Dense 128 → Softmax 24", _par("E3", "423.448")],
    ["E4", "CNN", "Arsitektur pada tabel di atas, tanpa augmentasi", _par("E4", "441.912")],
    ["E5", "CNN", "Arsitektur E4 + augmentasi pada tf.data (model akhir)", _par("E5", "441.912")],
], "Ringkasan arsitektur skenario E1–E5", widths=[1.2, 1.6, 11.0, 2.7], size=9, center_cols=[0, 1, 3])
H3("c. Konfigurasi pelatihan")
TABLE(["Konfigurasi", "E1", "E2", "E3", "E4", "E5"], [
    ["Optimizer", "SGD", "Adam", "Adam", "Adam", "Adam"],
    ["Learning rate awal", "0,01", "0,001", "0,001", "0,001", "0,001"],
    ["LR scheduler", "–", "ReduceLR", "–", "ReduceLR", "ReduceLR"],
    ["Batch size", *(str(META["batch_size"]),) * 5],
    ["Epoch maksimum", "30", "40", "30", "40", "50"],
    ["EarlyStopping (patience)", "5", "8", "5", "8", "10"],
    ["Dropout", "–", "0,3", "–", "0,25–0,40", "0,25–0,40"],
    ["BatchNorm", "–", "✓", "–", "✓", "✓"],
    ["Augmentasi", "–", "–", "–", "–", "✓"],
    ["Fungsi loss", *("Sparse CCE",) * 5],
    ["Inisialisasi bobot", *("He normal",) * 5],
], "Konfigurasi pelatihan (ReduceLR = ReduceLROnPlateau: faktor 0,5, patience 3, min_lr 1e-5)",
   widths=[4.3, 2.3, 2.5, 2.3, 2.5, 2.6], size=9, center_cols=[1, 2, 3, 4, 5])
P("Callback yang digunakan adalah EarlyStopping (monitor val_loss, restore_best_weights), ModelCheckpoint (bobot dengan "
  "val_loss terendah), ReduceLROnPlateau, CSVLogger dan TensorBoard untuk monitoring, serta LRLogger kustom yang mencatat "
  "learning rate setiap epoch. Fungsi loss sparse categorical cross-entropy, L = −(1/N) Σ log p(y_i | x_i), dipakai karena "
  "label berupa indeks bilangan bulat dan output berupa distribusi softmax.")
H3("d. Alasan pemilihan model dan arsitektur")
if not DRAFT:
    A_, O_ = R["activation"], R["optimizer"]
    best_act = max(A_, key=lambda k: A_[k]["val_acc"][-1])
    best_opt = min(O_, key=lambda k: O_[k]["val_loss"][-1])
else:
    best_act = best_opt = MARK
LIST([("Bertahap dari konsep dasar. ", "Perceptron dan MLP (DNN) dipakai untuk membuktikan konsep backpropagation dan "
       "menjadi baseline, lalu ditingkatkan menjadi CNN. Setiap peningkatan diukur sehingga kontribusinya terlihat."),
      ("CNN sesuai karakter data. ", "EDA menunjukkan korelasi piksel yang sangat lokal. Konvolusi memanfaatkan "
       "keterhubungan lokal dan berbagi bobot (LeCun et al., 1998), sehingga lebih efisien dan lebih tahan pergeseran "
       "daripada DNN."),
      ("Konvolusi 3×3 bertumpuk. ", "Dua konvolusi 3×3 memberi receptive field 5×5 dengan parameter lebih sedikit dan "
       "nonlinieritas lebih banyak (Simonyan & Zisserman, 2015)."),
      ("BatchNorm dan Dropout. ", "BatchNorm menstabilkan dan mempercepat pelatihan (Ioffe & Szegedy, 2015), sedangkan "
       "Dropout mencegah co-adaptation antarneuron (Srivastava et al., 2014)."),
      ("ReLU dan Adam berbasis eksperimen. ", f"Pada studi pendahuluan, aktivasi dengan akurasi validasi akhir tertinggi "
       f"adalah {best_act}, dan optimizer dengan loss validasi terendah adalah {best_opt}. ReLU dengan inisialisasi He "
       "(He et al., 2015) dan Adam (Kingma & Ba, 2015) dipilih karena konvergen cepat dan stabil."),
      ("Augmentasi untuk generalisasi. ", "Augmentasi mengatasi pergeseran distribusi antara data latih dan uji "
       "(Shorten & Khoshgoftaar, 2019).")])

H3("e. Dua fitur unik")
# ---- Fitur unik 1
H4("Fitur unik 1: Confidence threshold dan deteksi input di luar domain (OOD)")
RICH([("Apa fiturnya. ", "b"), ("Prediksi hanya diterima jika probabilitas softmax tertinggi (maximum softmax probability, "
      "MSP) ≥ ambang τ. Jika di bawah τ, sistem menjawab \"tidak yakin / bukan isyarat huruf ASL\".", "")])
RICH([("Mengapa dibutuhkan. ", "b"), ("Softmax selalu memaksakan salah satu dari 24 huruf, termasuk untuk foto yang bukan "
      "tangan. Pada aplikasi nyata, jawaban salah yang meyakinkan lebih merugikan daripada penolakan. MSP adalah baseline "
      "deteksi input di luar distribusi yang sederhana dan efektif (Hendrycks & Gimpel, 2017).", "")])
RICH([("Bagaimana implementasinya. ", "b"), ("(1) Data validasi diaugmentasi untuk mensimulasikan pergeseran distribusi. "
      f"(2) Dipilih τ terkecil yang membuat akurasi prediksi yang diterima ≥ {V(lambda: pct(F1['target_acc'], 1))}, dengan "
      f"batas bawah τ = 0,5. Hasilnya τ = {V(lambda: num(F1['tau'], 2))}. (3) τ disimpan di config.json dan dipakai "
      "aplikasi, dan pengguna dapat mengubahnya lewat slider. Data uji tidak dipakai dalam penentuan τ.", "")])
RICH([("Bukti dampaknya. ", "b"), (f"{REF('Tabel')} dan {REF('Tabel', 2)} membandingkan kondisi tanpa dan dengan fitur ini.", "")])
if not DRAFT and F1:
    t_ = F1["test"]
    rows = [["Akurasi prediksi yang diterima", pct(t_["acc_tanpa"]), pct(t_["acc_diterima"])],
            ["Coverage (prediksi yang diterima)", "100,00%", pct(t_["coverage"])],
            ["Kesalahan yang lolos ke pengguna", num(t_["n_salah"]), num(t_["n_salah"] - t_["salah_ditolak"])],
            ["Kesalahan yang berhasil ditolak", "–", f"{num(t_['salah_ditolak'])} dari {num(t_['n_salah'])}"],
            ["Prediksi benar yang ikut ditolak", "–", num(t_["benar_ditolak"])]]
else:
    rows = [[k, MARK, MARK] for k in ("Akurasi prediksi yang diterima", "Coverage (prediksi yang diterima)",
                                      "Kesalahan yang lolos ke pengguna", "Kesalahan yang berhasil ditolak",
                                      "Prediksi benar yang ikut ditolak")]
TABLE(["Metrik (data uji resmi)", "Tanpa ambang", f"Dengan ambang τ = {V(lambda: num(F1['tau'], 2))}"], rows,
      "Dampak fitur unik 1 pada data uji", widths=[7.5, 4.0, 5.0], center_cols=[1, 2])
if not DRAFT and F1.get("ood"):
    rows = [[k, num(v["n"]), pct(v["ditolak"], 1), num(v["mean_conf"], 3), num(v["auroc"], 3)] for k, v in F1["ood"].items()]
else:
    rows = [[k, MARK, MARK, MARK, MARK] for k in ("Derau Gaussian", "Citra polos", "Digit MNIST", "Fashion-MNIST")]
TABLE(["Set input di luar domain", "Jumlah", "% ditolak", "Rata-rata confidence", "AUROC"], rows,
      "Kemampuan menolak input di luar domain (tanpa fitur ini, 0% ditolak)", widths=[4.8, 2.2, 2.6, 3.8, 3.1])
if not DRAFT and F1:
    t_ = F1["test"]; ood = F1.get("ood", {})
    lemah = [k for k, v in ood.items() if v["ditolak"] < 0.5]
    kuat = [k for k, v in ood.items() if v["ditolak"] >= 0.5]
    P(f"Dengan τ = {num(F1['tau'], 2)}, akurasi prediksi yang diterima {naik(t_['acc_diterima'] - t_['acc_tanpa'])} "
      f"dari {pct(t_['acc_tanpa'])} menjadi {pct(t_['acc_diterima'])}, dan {pct(t_['coverage'])} prediksi tetap diterima. "
      f"Sebanyak {num(t_['salah_ditolak'])} dari {num(t_['n_salah'])} kesalahan ditahan sebelum sampai ke pengguna."
      + (f" Untuk input di luar domain, sebagian besar {gabung(kuat)} berhasil ditolak." if kuat else "")
      + (f" Namun, model masih cukup yakin pada {gabung(lemah)} sehingga banyak di antaranya tidak tertolak. Ini adalah "
         "keterbatasan MSP: CNN dapat memberi confidence tinggi pada input yang asing (Hendrycks & Gimpel, 2017). "
         "Perbaikan lanjutan dapat memakai temperature scaling atau kelas \"bukan tangan\" khusus." if lemah else ""))
else:
    P(f"Interpretasi dampak (akurasi, coverage, dan persentase input OOD yang ditolak) {MARK}.")
FIG("fig_fitur1_threshold", "Trade-off akurasi–coverage dan distribusi confidence in-domain vs di luar domain")
FIG("fig_fitur1_ood_contoh", "Contoh input di luar domain dan keputusan model dengan ambang")

# ---- Fitur unik 2
H4("Fitur unik 2: Kompresi model ke TensorFlow Lite")
RICH([("Apa fiturnya. ", "b"), ("Model Keras dikonversi ke TensorFlow Lite (LiteRT) dalam empat varian: FP32, FP16, "
      "dynamic range (bobot INT8), dan full integer INT8 yang dikalibrasi dengan representative dataset.", "")])
RICH([("Mengapa dibutuhkan. ", "b"), ("Aplikasi web gratis (Hugging Face Spaces CPU) dan perangkat seluler membutuhkan "
      "model kecil yang cepat dimuat dan cepat diinferensi. Model TFLite cukup dijalankan dengan runtime ai-edge-litert yang "
      "ringan, tanpa TensorFlow penuh. Kuantisasi INT8 mengecilkan model sekitar 4× dengan penurunan akurasi yang kecil "
      "(Jacob et al., 2018).", "")])
RICH([("Bagaimana implementasinya. ", "b"), ("Model diekspor dengan model.export() ke SavedModel, lalu dikonversi dengan "
      "tf.lite.TFLiteConverter (optimizations = DEFAULT, supported_types = float16, atau representative_dataset untuk INT8). "
      "Setiap varian dievaluasi pada seluruh 7.172 citra uji menggunakan Interpreter. Varian terkecil yang akurasinya "
      "turun tidak lebih dari 0,2 poin dipilih untuk deployment.", "")])
if not DRAFT and F2.get("varian"):
    kk = F2["keras"]
    rows = [["Keras .keras (asli)", num(kk["size_kb"], 1), pct(kk["test_acc"]), pct(kk["macro_f1"]), "100,00%", num(kk["latency_ms"], 3)]]
    rows += [[k + (" ★" if k == F2.get("deploy_variant") else ""), num(v["size_kb"], 1), pct(v["test_acc"]), pct(v["macro_f1"]),
              pct(v["agreement"]), num(v["latency_ms"], 3)] for k, v in F2["varian"].items()]
else:
    rows = [[k] + [MARK] * 5 for k in ("Keras .keras (asli)", "FP32", "FP16", "Dynamic range INT8", "Full INT8")]
RICH([("Bukti dampaknya. ", "b"), (f"{REF('Tabel')} dan {REF('Gambar')} membandingkan model asli dengan varian TFLite.", "")])
TABLE(["Model", "Ukuran (KB)", "Akurasi uji", "Macro-F1", "Sama dengan Keras", "Latensi (ms/citra)"], rows,
      "Dampak fitur unik 2: ukuran, akurasi, dan latensi (★ = varian untuk deployment)",
      widths=[4.4, 2.2, 2.4, 2.4, 2.6, 2.5], size=9)
if not DRAFT and F2.get("varian") and F2.get("deploy_variant"):
    dv = F2["deploy_variant"]; v = F2["varian"][dv]; kk = F2["keras"]
    P(f"Varian {dv} berukuran {num(v['size_kb'], 1)} KB, atau {num(kk['size_kb'] / v['size_kb'], 1)}× lebih kecil "
      f"daripada berkas Keras ({num(kk['size_kb'], 1)} KB). Akurasi ujinya {pct(v['test_acc'])} (selisih "
      f"{pp(v['test_acc'] - kk['test_acc'])} dari model asli), dan {pct(v['agreement'])} prediksinya sama dengan model "
      f"Keras. Latensinya {num(v['latency_ms'], 3)} ms per citra di CPU, dibandingkan {num(kk['latency_ms'], 3)} ms pada "
      "pemanggilan model Keras.")
else:
    P(f"Interpretasi pengurangan ukuran, perubahan akurasi, dan latensi {MARK}.")
FIG("fig_fitur2_tflite", "Ukuran berkas dan akurasi uji model Keras vs varian TFLite")

# ---------------------------------------------------------------- 6. Hasil evaluasi
H2("6. Hasil Evaluasi")
H3("a. Metrik evaluasi yang digunakan dan alasannya")
LIST([("Accuracy: ", "proporsi prediksi benar; ukuran utama yang mudah dipahami."),
      ("Precision, recall, dan F1-score per kelas: ", "precision = TP/(TP+FP), recall = TP/(TP+FN), dan F1 = 2PR/(P+R). "
       "Ketiganya mengungkap huruf mana yang sering salah (Sokolova & Lapalme, 2009)."),
      ("Macro-F1: ", "rata-rata F1 dengan bobot sama per kelas; penting karena data uji timpang."),
      ("Weighted-F1 dan loss uji: ", "pelengkap untuk melihat performa berbobot support dan tingkat keyakinan model."),
      ("Confusion matrix 24×24: ", "memperlihatkan pasangan huruf yang tertukar."),
      ("Coverage, AUROC OOD, ukuran model, dan latensi: ", "untuk mengukur dampak dua fitur unik.")], style="bullet")
H3("b. Nilai metrik pada data testing")
TABLE(["Metrik (model E5, 7.172 citra uji)", "Nilai"], [
    ["Accuracy", V(lambda: pct(FIN["acc"]))],
    ["Prediksi benar / total", V(lambda: f"{num(FIN['n_benar'])} / {num(FIN['n_test'])}")],
    ["Macro precision", V(lambda: pct(FIN["macro_p"]))],
    ["Macro recall", V(lambda: pct(FIN["macro_r"]))],
    ["Macro F1-score", V(lambda: pct(FIN["macro_f1"]))],
    ["Weighted F1-score", V(lambda: pct(FIN["weighted_f1"]))],
    ["Loss uji (sparse CCE)", V(lambda: num(FIN["test_loss"], 4))],
    ["Latensi Keras", V(lambda: f"{num(FIN['latency_single_ms'], 2)} ms/citra (tunggal), {num(FIN['latency_batch_ms'], 3)} ms/citra (batch)")],
], "Nilai metrik model akhir pada data uji", widths=[9, 7.5], center_cols=[1])
if not DRAFT:
    rows = [[c, pct(v["precision"]), pct(v["recall"]), pct(v["f1"]), num(v["support"])] for c, v in FIN["per_class"].items()]
else:
    rows = [[c, MARK, MARK, MARK, MARK] for c in "ABCDEFGHIKLMNOPQRSTUVWXY"]
TABLE(["Huruf", "Precision", "Recall", "F1-score", "Support"], rows, "Classification report per huruf (E5)",
      widths=[2.5, 3.5, 3.5, 3.5, 3.5], size=8.5, center_cols=[0, 1, 2, 3, 4])

H3("c. Grafik loss dan akurasi data training dan validation")
if not DRAFT and "E5" in E:
    e5 = E["E5"]; lrs = e5.get("history", {}).get("lr", [])
    n_red = sum(1 for a_, b_ in zip(lrs[:-1], lrs[1:]) if b_ < a_ * 0.99)
    g5 = e5["train_acc"] - e5["val_acc"]
    kal_fit = ("Akurasi latih lebih rendah daripada validasi karena data latih diaugmentasi dan dropout hanya aktif saat "
               "latih, sehingga tidak ada indikasi overfitting." if g5 < 0 else
               "Selisih latih–validasi kecil, sehingga tidak ada indikasi overfitting yang berarti." if g5 < 0.02 else
               "Terdapat selisih latih–validasi yang menandakan sedikit overfitting.")
    P(f"{REF('Gambar')} menampilkan kurva loss, akurasi, dan learning rate model akhir E5. Pelatihan berjalan "
      f"{e5['epochs_run']} epoch dengan epoch terbaik ke-{e5['best_epoch']}. Pada epoch tersebut, akurasi latih "
      f"{pct(e5['train_acc'])} dan akurasi validasi {pct(e5['val_acc'])}. {kal_fit}"
      + (f" Learning rate diturunkan {n_red} kali oleh ReduceLROnPlateau." if lrs else "")
      + f" Sebagai pembanding, {REF('Gambar', 2)} menampilkan kurva DNN baseline E1.")
else:
    P(f"{REF('Gambar')} menampilkan kurva loss, akurasi, dan learning rate model akhir E5, sedangkan {REF('Gambar', 2)} "
      f"menampilkan kurva DNN baseline E1. Interpretasinya {MARK}.")
FIG("fig_hist_E5", "Kurva loss, akurasi, dan learning rate (latih vs validasi) model E5")
FIG("fig_hist_E1", "Kurva loss dan akurasi (latih vs validasi) DNN baseline E1")

H3("d. Perbandingan antarskenario")
P(f"{REF('Tabel')} dan {REF('Gambar')} membandingkan kelima skenario. Gap adalah selisih akurasi validasi dan uji.")
if not DRAFT:
    rows = [[e["id"], e["nama"].replace(" (model akhir)", ""), f"{e['epochs_run']} ({e['best_epoch']})", num(e["train_time_s"], 0),
             pct(e["train_acc"]), pct(e["val_acc"]), pct(e["test_acc"]), pct(e["macro_f1"]),
             _id(f"{(e['val_acc'] - e['test_acc']) * 100:.2f}") + " pp"] for e in R["experiments"]]
else:
    rows = [[i, n] + [MARK] * 7 for i, n in (("E1", "MLP Baseline"), ("E2", "MLP + BN + Dropout + Adam"), ("E3", "CNN Baseline"),
                                             ("E4", "CNN + BN + Dropout"), ("E5", "CNN + BN + Dropout + Augmentasi"))]
TABLE(["ID", "Skenario", "Epoch (terbaik)", "Waktu (s)", "Acc latih", "Acc val.", "Acc uji", "Macro-F1", "Gap val–uji"], rows,
      "Hasil seluruh skenario", widths=[0.9, 4.0, 1.7, 1.3, 1.6, 1.6, 1.6, 1.8, 2.0], size=8)
FIG("fig_perbandingan_eksperimen", "Perbandingan performa skenario E1–E5 dan kurva akurasi validasi")
if not DRAFT and all(k in E for k in ("E1", "E2", "E3", "E4", "E5")):
    dd = lambda a, b: E[b]["test_acc"] - E[a]["test_acc"]
    LIST([f"E1 → E2 (regularisasi DNN) {naik_turun(dd('E1', 'E2'))} akurasi uji {pp(dd('E1', 'E2'))} menjadi {pct(E['E2']['test_acc'])}.",
          f"E2 → E3 (DNN diganti CNN) {naik_turun(dd('E2', 'E3'))} akurasi uji {pp(dd('E2', 'E3'))} menjadi {pct(E['E3']['test_acc'])}, dengan parameter lebih sedikit ({num(E['E3']['params'])} vs {num(E['E2']['params'])}).",
          f"E3 → E4 (blok konvolusi bertumpuk + BN + Dropout + scheduler) {naik_turun(dd('E3', 'E4'))} akurasi uji {pp(dd('E3', 'E4'))} menjadi {pct(E['E4']['test_acc'])}.",
          f"E4 → E5 (augmentasi) {naik_turun(dd('E4', 'E5'))} akurasi uji {pp(dd('E4', 'E5'))} menjadi {pct(E['E5']['test_acc'])}; gap val–uji berubah dari {pp(E['E4']['val_acc'] - E['E4']['test_acc'])} menjadi {pp(E['E5']['val_acc'] - E['E5']['test_acc'])}.",
          f"Secara keseluruhan, E1 → E5 {naik_turun(dd('E1', 'E5'))} akurasi uji sebesar {pp(dd('E1', 'E5'))}."])
else:
    P(f"Analisis kenaikan akurasi per skenario {MARK}.")
H4("Studi fungsi aktivasi dan optimizer")
if not DRAFT:
    rows = [[k, pct(v["val_acc"][0]), pct(v["val_acc"][-1]), pct(v["test_acc"])] for k, v in R["activation"].items()]
else:
    rows = [[k, MARK, MARK, MARK] for k in ("sigmoid", "tanh", "relu")]
TABLE(["Aktivasi (DNN 784–256–128–24)", "Val. acc epoch 1", "Val. acc akhir", "Akurasi uji"], rows,
      "Perbandingan fungsi aktivasi (Adam, 15 epoch)", widths=[6.0, 3.5, 3.5, 3.5])
FIG("fig_aktivasi", "Loss latih dan akurasi validasi per fungsi aktivasi")
if not DRAFT:
    rows = [[k, num(v["train_loss"][-1], 4), num(v["val_loss"][-1], 4), pct(v["val_acc"][-1]), pct(v["test_acc"])]
            for k, v in R["optimizer"].items()]
else:
    rows = [[k, MARK, MARK, MARK, MARK] for k in ("SGD (η=0.01)", "SGD+Momentum (0.9)", "RMSProp (η=0.001)", "Adam (η=0.001)")]
TABLE(["Optimizer (CNN E4, 10 epoch)", "Loss latih", "Loss val.", "Val. acc", "Akurasi uji"], rows,
      "Perbandingan optimizer", widths=[5.5, 2.6, 2.6, 2.8, 3.0])
FIG("fig_optimizer", "Loss latih dan akurasi validasi per optimizer")
H4("Fondasi: perceptron dan MLP dengan backpropagation manual")
if not DRAFT:
    pc = R["perceptron"]; err = pc["errors"]
    conv_ep = next((i + 1 for i, e_ in enumerate(err) if e_ == 0), None)
    P(f"Perceptron Rosenblatt (Rosenblatt, 1958) untuk huruf A vs B menurunkan jumlah kesalahan per epoch dari {err[0]} "
      f"menjadi {err[-1]}" + (f" dan mencapai nol pada epoch ke-{conv_ep}" if conv_ep else "") +
      f", dengan akurasi uji {pct(pc['test_acc'])}. Neuron sigmoid dengan gradient descent mencapai akurasi uji "
      f"{pct(pc['sigmoid']['test_acc'])}. MLP 784–128–24 yang dilatih dengan custom training loop (tf.GradientTape dan "
      f"SGD manual) mencapai akurasi uji {pct(pc['mlp_manual']['test_acc'])} untuk 24 kelas. Hasil ini membuktikan bahwa "
      "mekanisme backpropagation (Rumelhart et al., 1986) bekerja sebelum memakai model.fit().")
else:
    P(f"Hasil perceptron A vs B, neuron sigmoid, dan MLP manual {MARK}.")
FIG("fig_perceptron_mlp_manual", "Kesalahan perceptron, loss neuron sigmoid, dan kurva MLP manual")
H4("Dengan dan tanpa fitur unik, serta perbandingan dengan proyek sebelumnya")
TABLE(["Perbandingan", "Tanpa", "Dengan"], [
    ["Fitur 1: akurasi prediksi yang diterima", V(lambda: pct(F1["test"]["acc_tanpa"])),
     V(lambda: pct(F1["test"]["acc_diterima"]) + f" (coverage {pct(F1['test']['coverage'])})")],
    ["Fitur 1: input OOD yang ditolak (rata-rata)", "0%",
     V(lambda: pct(sum(v["ditolak"] for v in F1["ood"].values()) / len(F1["ood"]), 1))],
    ["Fitur 2: ukuran model", V(lambda: num(F2["keras"]["size_kb"], 1) + " KB"),
     V(lambda: num(F2["varian"][F2["deploy_variant"]]["size_kb"], 1) + f" KB ({F2['deploy_variant']})")],
    ["Fitur 2: akurasi uji", V(lambda: pct(F2["keras"]["test_acc"])), V(lambda: pct(F2["varian"][F2["deploy_variant"]]["test_acc"]))],
], "Ringkasan perbandingan tanpa dan dengan fitur unik", widths=[7.0, 4.0, 5.5])
P(f"{REF('Gambar')} membandingkan hasil proyek sebelumnya (71,67% pada 240 citra) dengan proyek ini (diuji pada 7.172 "
  "citra). Karena set ujinya berbeda, perbandingan ini bersifat indikatif; namun set uji baru 30× lebih besar dan "
  "sepenuhnya terpisah dari pelatihan.")
FIG("fig_perbandingan_lama_baru", "Akurasi proyek sebelumnya (abu-abu) vs proyek ini (biru)")

H3("e. Analisis hasil")
P(f"Confusion matrix pada {REF('Gambar')} dan akurasi per huruf pada {REF('Gambar', 2)} memperlihatkan huruf yang paling "
  "sering salah.")
FIG("fig_confusion_matrix", "Confusion matrix 24×24 model E5 pada data uji", width=15)
FIG("fig_akurasi_per_kelas", "Akurasi (recall) per huruf, diurutkan; merah < 95%")
MIRIP = {
    frozenset("MN"): "M dan N sama-sama kepalan dengan ibu jari terselip; bedanya hanya jumlah jari yang menutupi ibu jari (tiga vs dua)",
    frozenset("UV"): "U dan V sama-sama mengangkat telunjuk dan jari tengah; bedanya hanya rapat atau renggangnya kedua jari",
    frozenset("RU"): "R dan U sama-sama mengangkat dua jari; pada R kedua jari bersilang",
    frozenset("KV"): "K dan V sama-sama mengangkat dua jari; pada K ibu jari berada di antara keduanya",
    frozenset("VW"): "V dan W hanya berbeda satu jari terangkat",
    frozenset("GH"): "G dan H sama-sama menunjuk ke samping; H memakai dua jari",
    frozenset("PQ"): "P dan Q sama-sama mengarah ke bawah",
    frozenset("KP"): "P adalah konfigurasi K yang diarahkan ke bawah",
    frozenset("CO"): "C dan O sama-sama lengkung; pada O ujung jari menyentuh ibu jari",
    frozenset("IY"): "I dan Y sama-sama mengangkat kelingking; Y juga merentangkan ibu jari",
    frozenset("DI"): "D dan I sama-sama mengangkat satu jari (telunjuk vs kelingking)",
    frozenset("DX"): "X adalah telunjuk yang ditekuk, mirip D dengan telunjuk terangkat",
    frozenset("EO"): "E dan O sama-sama melipat jari ke arah ibu jari",
}


def alasan(a, b):
    if frozenset(a + b) in MIRIP:
        return MIRIP[frozenset(a + b)]
    if a in "AEMNST" and b in "AEMNST":
        return f"{a} dan {b} sama-sama huruf berbasis kepalan yang hanya berbeda pada posisi ibu jari, detail yang sangat kecil di 28×28"
    return f"{a} dan {b} memiliki siluet dan konfigurasi jari yang mirip pada resolusi rendah"


if not DRAFT and FIN.get("top_confusions"):
    tcf = FIN["top_confusions"][:10]
    TABLE(["No", "Aktual → Prediksi", "Jumlah", "Kemungkinan penyebab"],
          [[str(i + 1), f"{a} → {b}", num(n), alasan(a, b)] for i, (a, b, n) in enumerate(tcf)],
          "Pasangan huruf yang paling sering tertukar", widths=[1.0, 3.0, 1.8, 10.7], size=8.5, center_cols=[0, 1, 2])
    pcs = FIN["per_class"]
    terendah = sorted(pcs, key=lambda c: pcs[c]["recall"])[:3]
    sempurna = [c for c, v in pcs.items() if v["recall"] >= 0.9999]
    P(f"Sebanyak {len(sempurna)} dari 24 huruf dikenali dengan recall 100%, sedangkan recall terendah dimiliki oleh "
      + ", ".join(f"{c} ({pct(pcs[c]['recall'])})" for c in terendah) + f". Dari {num(FIN['n_test'])} citra uji, "
      f"{num(FIN['n_salah'])} salah diklasifikasikan. Rata-rata confidence prediksi salah adalah "
      f"{pct(FIN['mean_conf_salah']) if FIN.get('mean_conf_salah') is not None else '–'}, dibandingkan "
      f"{pct(FIN['mean_conf_benar'])} pada prediksi benar. Perbedaan inilah yang dimanfaatkan fitur unik 1.")
else:
    TABLE(["No", "Aktual → Prediksi", "Jumlah", "Kemungkinan penyebab"], [[MARK] * 4],
          "Pasangan huruf yang paling sering tertukar", widths=[1.0, 3.0, 1.8, 10.7], size=8.5)
    P(f"Huruf dengan recall terendah dan perbandingan confidence prediksi benar/salah {MARK}.")
FIG("fig_salah_klasifikasi", "Contoh salah klasifikasi dengan confidence tertinggi")
P(f"Grad-CAM (Selvaraju et al., 2017) pada {REF('Gambar')} menunjukkan area yang paling memengaruhi keputusan. Pada "
  "prediksi benar, heatmap idealnya terpusat pada jari dan ibu jari. Pada prediksi salah, heatmap membantu melihat apakah "
  "model terkecoh oleh bagian yang tidak relevan (latar, pergelangan) atau memang gagal membedakan detail jari.")
FIG("fig_gradcam", "Grad-CAM: baris atas prediksi benar, baris bawah prediksi salah", width=15)
if not DRAFT:
    ind = []
    for e in R["experiments"]:
        gl = e["train_acc"] - e["val_acc"]; gu = e["val_acc"] - e["test_acc"]
        if e["train_acc"] < 0.9 and e["val_acc"] < 0.9:
            s = "underfitting (akurasi latih dan validasi masih rendah)"
        elif gu > 0.05:
            s = "generalisasi lemah ke data uji (gap val–uji besar), yaitu overfitting terhadap distribusi data latih"
        elif gl > 0.02:
            s = "sedikit overfitting (akurasi latih > validasi)"
        else:
            s = "tidak ada indikasi overfitting maupun underfitting yang berarti"
        ind.append(f"{e['id']}: latih {pct(e['train_acc'])}, validasi {pct(e['val_acc'])}, uji {pct(e['test_acc'])}, sehingga {s}.")
    P("Indikasi overfitting/underfitting per skenario:")
    LIST(ind, style="bullet")
else:
    P(f"Indikasi overfitting/underfitting per skenario {MARK}.")

# ---------------------------------------------------------------- 7. Teknologi
H2("7. Teknologi yang Digunakan")
H3("a. Bahasa pemrograman")
P(f"Python {V(lambda: R['meta']['python'])}.")
H3("b. Library dan versinya")
TABLE(["Library", "Versi", "Kegunaan"], [
    ["TensorFlow", V(lambda: R["meta"]["tensorflow"]), "Tensor, tf.data, GradientTape, konverter TFLite"],
    ["Keras", V(lambda: R["meta"]["keras"]), "Model DNN/CNN, callback, augmentasi"],
    ["NumPy", V(lambda: R["meta"]["numpy"]), "Komputasi numerik"],
    ["pandas", "bawaan Colab", "Membaca CSV dan analisis tabel"],
    ["scikit-learn", V(lambda: R["meta"]["sklearn"]), "Split stratified, metrik, AUROC"],
    ["OpenCV", V(lambda: R["meta"]["opencv"]), "Resize heatmap Grad-CAM"],
    ["Matplotlib / Seaborn", "bawaan Colab", "Visualisasi"],
    ["Pillow", "bawaan", "Praproses foto pada aplikasi"],
    ["ai-edge-litert (LiteRT)", "terbaru", "Runtime TFLite pada aplikasi"],
    ["Gradio", "5.9.1", "Antarmuka aplikasi web"],
    ["python-docx", "terbaru", "Pembuatan laporan otomatis"],
], "Library yang digunakan", widths=[4.5, 3.5, 8.5], center_cols=[1])
H3("c. Lingkungan pengembangan")
P(f"Google Colaboratory dengan akselerator {V(lambda: R['meta']['gpu'])}. Total waktu menjalankan seluruh notebook adalah "
  f"{V(lambda: num(R['meta']['runtime_total_min'], 1))} menit (seed global {META['seed']}).")
H3("d. Framework aplikasi dan platform hosting")
LIST(["Gradio sebagai framework antarmuka web (Abid et al., 2019).",
      "Hugging Face Spaces (SDK Gradio, CPU gratis) sebagai platform hosting aplikasi.",
      "GitHub sebagai repositori kode yang berisi notebook, folder deploy, hasil, dan README.",
      "TensorFlow Lite (LiteRT) sebagai format dan runtime model pada aplikasi."], style="bullet")

# ---------------------------------------------------------------- 8. Deployment & demo
H2("8. Deployment dan Demo")
H3("a. Link deployment")
TABLE(["Artefak", "Link"], [
    ["Aplikasi (Hugging Face Spaces)", link_or_isi("app", "link Hugging Face Space")],
    ["Repositori GitHub", link_or_isi("github", "link repositori GitHub")],
    ["Video demo & deployment", link_or_isi("video", "link video (YouTube/Google Drive)")],
], "Link deployment", widths=[5.0, 11.5], center_cols=[])
H3("b. Tahapan deployment")
LANGKAH = [
    ("Menyimpan model",
     "Model akhir disimpan sebagai .keras, lalu dikonversi ke model.tflite (fitur unik 2). Daftar label dan ambang τ "
     "disimpan di config.json.",
     'final_model.save("outputs/models/E5_model_akhir.keras")\n'
     'final_model.export(export_dir)\n'
     'conv = tf.lite.TFLiteConverter.from_saved_model(export_dir)\n'
     'conv.optimizations = [tf.lite.Optimize.DEFAULT]   # sesuai varian terpilih\n'
     'open("deploy/model.tflite", "wb").write(conv.convert())'),
    ("Menyiapkan skrip inferensi",
     "inference.py memuat model.tflite dengan Interpreter LiteRT, menyamakan praproses dengan pelatihan (grayscale → "
     "crop tengah → 28×28 → /255), lalu mengembalikan label, confidence, status diterima/ditolak, dan top-3. Skrip ini "
     f"diuji di notebook; prediksinya {V(lambda: pct(DEP['agreement']))} sama dengan model Keras pada "
     f"{V(lambda: num(DEP['n_cek']))} citra uji.",
     'clf = SignClassifier()            # memuat model.tflite + config.json\n'
     'hasil = clf.predict(foto, threshold=0.9)\n'
     '# -> {"label": "A", "confidence": 0.99, "accepted": True, "top3": [...]}'),
    ("Membangun antarmuka",
     "app.py menyediakan unggah foto, kamera, dan clipboard; slider ambang τ; Top-3 prediksi; status diterima/ditolak; "
     "pratinjau 28×28; dan contoh input.",
     'with gr.Blocks() as demo:\n'
     '    inp = gr.Image(type="numpy", sources=["upload", "webcam", "clipboard"])\n'
     '    thr = gr.Slider(0, 1, value=clf.threshold, label="Ambang confidence")\n'
     '    btn.click(prediksi, [inp, thr], [out_label, out_status, out_prev])'),
    ("Mencatat dependensi",
     "requirements.txt hanya berisi ai-edge-litert, numpy, dan pillow, karena Gradio sudah disediakan oleh SDK Space. "
     "README.md berisi konfigurasi Space (sdk: gradio, app_file: app.py).",
     "# requirements.txt\nai-edge-litert\nnumpy\npillow"),
    ("Mengunggah ke platform hosting",
     "Buat Space baru di huggingface.co (SDK Gradio, hardware CPU basic), lalu unggah seluruh isi folder deploy/ melalui "
     "menu Files atau git push. Kode lengkap juga di-push ke repositori GitHub.",
     "git clone https://huggingface.co/spaces/<username>/<nama-space>\ncp -r outputs/deploy/* <nama-space>/\n"
     "cd <nama-space> && git add . && git commit -m \"deploy\" && git push"),
    ("Menguji aplikasi",
     "Setelah build selesai, aplikasi diuji dengan contoh prediksi benar, contoh salah, input di luar domain (derau), dan "
     "foto tangan asli. Diperiksa juga bahwa slider ambang mengubah keputusan diterima/ditolak.",
     None),
]
for i, (judul, desk, kode) in enumerate(LANGKAH, 1):
    H4(f"Langkah {i}: {judul}")
    P(desk)
    if kode:
        CODE(kode)
    SCREENSHOT(i, f"Tangkapan layar langkah {i}: {judul.lower()}")
H3("c. Penjelasan video demo dan deployment")
H4("i. Tampilan user interface")
P("Antarmuka terdiri atas dua kolom. Kolom kiri berisi input citra (unggah, kamera, clipboard), slider ambang τ, dan "
  "tombol Prediksi. Kolom kanan berisi Top-3 prediksi beserta probabilitasnya, status keputusan (✅ diterima atau ⚠️ "
  "ditolak), dan pratinjau citra 28×28 setelah praproses. Di bawahnya tersedia contoh input yang dapat diklik.")
SCREENSHOT(7, "Tampilan antarmuka aplikasi di Hugging Face Spaces")
H4("ii. Contoh input dan output")
P(f"{REF('Gambar')} memperlihatkan output aplikasi (melalui inference.py yang sama) untuk contoh prediksi benar, "
  "prediksi salah, dan input di luar domain.")
FIG("fig_deploy_contoh", "Contoh input dan output aplikasi: prediksi benar, prediksi salah, dan input di luar domain")
if not DRAFT and DEP.get("contoh"):
    TABLE(["Berkas contoh", "Label aktual", "Output aplikasi", "Confidence", "Keputusan"],
          [[c["file"], c["aktual"], c["prediksi"], pct(c["confidence"], 1), "Diterima" if c["diterima"] else "Ditolak"]
           for c in DEP["contoh"]],
          "Ringkasan contoh input dan output aplikasi", widths=[5.6, 2.4, 2.8, 2.6, 3.1], size=9)
H4("iii. Demonstrasi dua fitur unik")
LIST([f"Fitur 1: unggah contoh derau atau foto yang bukan tangan; aplikasi menampilkan status \"tidak yakin\" karena "
      f"confidence < τ = {V(lambda: num(F1['tau'], 2))}. Geser slider untuk memperlihatkan keputusan diterima/ditolak berubah.",
      f"Fitur 2: tunjukkan berkas model.tflite ({V(lambda: num(DEP['model_kb'], 1))} KB) di tab Files Space, bandingkan dengan "
      f"model Keras ({V(lambda: num(F2['keras']['size_kb'], 1))} KB), dan tunjukkan bahwa aplikasi berjalan tanpa TensorFlow penuh."],
     style="bullet")
H4("iv. Penjelasan singkat tahapan deployment")
P("Video menjelaskan alur singkatnya: model disimpan dan dikonversi ke TFLite, skrip inferensi dan antarmuka Gradio "
  "disiapkan, dependensi dicatat, berkas diunggah ke Hugging Face Spaces, lalu aplikasi diuji. Pembagian segmen video "
  f"ada pada {REF('Tabel')}.")
segmen = ["Pembuka, masalah, dataset, dan arsitektur CNN", "EDA, preprocessing, dan pipeline tf.data",
          "DNN, perceptron, studi aktivasi dan optimizer", "Hasil evaluasi, analisis error, Grad-CAM, dan dua fitur unik",
          "Demo aplikasi dan tahapan deployment"]
TABLE(["Segmen", "Materi", "Pembawa"],
      [[str(i + 1), segmen[i], a["nama"]] for i, a in enumerate(META["anggota"][:5])],
      "Pembagian segmen video (setiap anggota menjelaskan bagiannya)", widths=[1.6, 9.4, 5.5], center_cols=[0])

# ---------------------------------------------------------------- 9. Lampiran
H2("9. Lampiran")
H3("a. Cuplikan kode penting")
H4("Preprocessing dan pipeline tf.data")
CODE("""
X_full = (X_all_raw.astype("float32") / 255.0)[..., np.newaxis]     # (N,28,28,1)
X_train, X_val, y_train, y_val = train_test_split(
    X_full, y_full, test_size=0.10, stratify=y_full, random_state=SEED)

geo_augment = keras.Sequential([layers.RandomRotation(0.03, fill_mode="nearest"),
                                layers.RandomZoom(0.10, fill_mode="nearest"),
                                layers.RandomTranslation(0.10, 0.10, fill_mode="nearest")])
def photometric_augment(x, y):
    x = tf.image.random_brightness(x, max_delta=0.10)
    x = tf.image.random_contrast(x, lower=0.8, upper=1.2)
    return tf.clip_by_value(x, 0.0, 1.0), y
""")
H4("Arsitektur CNN akhir")
CODE("""
def conv_bn(filters, name):
    return [layers.Conv2D(filters, 3, padding="same", use_bias=False, kernel_initializer="he_normal", name=f"{name}_conv"),
            layers.BatchNormalization(name=f"{name}_bn"), layers.Activation("relu", name=f"{name}_relu")]

model = keras.Sequential(
    [keras.Input((28, 28, 1))]
    + conv_bn(32, "b1c1") + conv_bn(32, "b1c2") + [layers.MaxPooling2D(2), layers.Dropout(0.25)]
    + conv_bn(64, "b2c1") + conv_bn(64, "b2c2") + [layers.MaxPooling2D(2), layers.Dropout(0.25)]
    + conv_bn(128, "b3c1") + [layers.MaxPooling2D(2), layers.Dropout(0.30)]
    + [layers.Flatten(), layers.Dense(256, use_bias=False), layers.BatchNormalization(),
       layers.Activation("relu"), layers.Dropout(0.4), layers.Dense(24, activation="softmax")])
model.compile(optimizer=keras.optimizers.Adam(1e-3), loss="sparse_categorical_crossentropy", metrics=["accuracy"])
callbacks = [EarlyStopping(monitor="val_loss", patience=10, restore_best_weights=True),
             ModelCheckpoint("E5_best.weights.h5", save_best_only=True, save_weights_only=True),
             ReduceLROnPlateau(monitor="val_loss", factor=0.5, patience=3, min_lr=1e-5), CSVLogger(...), TensorBoard(...)]
""")
H4("Fitur unik 1: penentuan ambang τ")
CODE("""
val_aug_ds = make_dataset(X_val, y_val, augment=True)   # validasi teraugmentasi (bukan uji)
prob_va = final_model.predict(val_aug_ds, verbose=0)
conf_va, pred_va = prob_va.max(1), prob_va.argmax(1)
for t in np.arange(0.0, 1.0, 0.01):          # tau terkecil yang memenuhi target
    m = conf_va >= t
    if (pred_va[m] == y_val[m]).mean() >= TARGET_ACC_DITERIMA:
        TAU = float(t); break
TAU = max(TAU, 0.50)
""")
H4("Fitur unik 2: konversi ke TensorFlow Lite")
CODE("""
final_model.export(export_dir)
def convert(varian):
    conv = tf.lite.TFLiteConverter.from_saved_model(export_dir)
    if varian == "FP16":
        conv.optimizations = [tf.lite.Optimize.DEFAULT]; conv.target_spec.supported_types = [tf.float16]
    elif varian == "Dynamic range INT8":
        conv.optimizations = [tf.lite.Optimize.DEFAULT]
    elif varian == "Full INT8":
        conv.optimizations = [tf.lite.Optimize.DEFAULT]; conv.representative_dataset = representative_data
    return conv.convert()
""")
H4("Inferensi pada aplikasi (inference.py)")
CODE("""
def preprocess(image):
    img = Image.fromarray(np.asarray(image).astype(np.uint8)).convert("L")
    w, h = img.size; s = min(w, h)
    img = img.crop(((w - s) // 2, (h - s) // 2, (w - s) // 2 + s, (h - s) // 2 + s)).resize((28, 28), Image.BOX)
    return (np.asarray(img, dtype=np.float32) / 255.0).reshape(1, 28, 28, 1)

def predict(self, image, threshold=None):
    self.itp.set_tensor(self.inp["index"], self.preprocess(image)); self.itp.invoke()
    p = self.itp.get_tensor(self.out["index"])[0]
    tau = self.threshold if threshold is None else threshold
    return {"label": self.labels[p.argmax()], "confidence": float(p.max()), "accepted": bool(p.max() >= tau)}
""")
H3("b. Grafik evaluasi")
P("Grafik evaluasi tambahan yang melengkapi Bagian 6:")
for eid, judul in (("E2", "DNN teregularisasi E2"), ("E3", "CNN baseline E3"), ("E4", "CNN teregularisasi E4")):
    FIG(f"fig_hist_{eid}", f"Kurva loss dan akurasi (latih vs validasi) {judul}")
FIG("fig_confidence", "Distribusi confidence prediksi benar vs salah (E5)", width=12)
FIG("fig_lr_schedule", "Strategi learning rate scheduling (konsep)", width=12)
FIG("fig_tensor", "Representasi tensor: citra (28×28×1), reduce_mean, dan reshape ke (B, 784)", width=13)

# ---------------------------------------------------------------- 10. Daftar pustaka
H2("10. Daftar Pustaka")
REF_LIST = [
    "Abadi, M., Barham, P., Chen, J., Chen, Z., Davis, A., Dean, J., … Zheng, X. (2016). TensorFlow: A system for large-scale machine learning. In 12th USENIX Symposium on Operating Systems Design and Implementation (OSDI 16) (pp. 265–283).",
    "Abid, A., Abdalla, A., Abid, A., Khan, D., Alfozan, A., & Zou, J. (2019). Gradio: Hassle-free sharing and testing of ML models in the wild. arXiv:1906.02569.",
    "Chollet, F., et al. (2015). Keras. https://keras.io",
    "Datamunge. (2017). Sign Language MNIST: Drop-in replacement for MNIST for hand gesture recognition tasks [Dataset]. Kaggle. https://www.kaggle.com/datasets/datamunge/sign-language-mnist",
    "Gad, A. F., & Jarmouni, F. E. (2020). Introduction to deep learning and neural networks with Python™: A practical guide. Academic Press.",
    "Goodfellow, I., Bengio, Y., & Courville, A. (2016). Deep learning. MIT Press.",
    "He, K., Zhang, X., Ren, S., & Sun, J. (2015). Delving deep into rectifiers: Surpassing human-level performance on ImageNet classification. In Proceedings of the IEEE International Conference on Computer Vision (pp. 1026–1034).",
    "Hendrycks, D., & Gimpel, K. (2017). A baseline for detecting misclassified and out-of-distribution examples in neural networks. In International Conference on Learning Representations (ICLR).",
    "Ioffe, S., & Szegedy, C. (2015). Batch normalization: Accelerating deep network training by reducing internal covariate shift. In Proceedings of the 32nd International Conference on Machine Learning (pp. 448–456).",
    "Jacob, B., Kligys, S., Chen, B., Zhu, M., Tang, M., Howard, A., Adam, H., & Kalenichenko, D. (2018). Quantization and training of neural networks for efficient integer-arithmetic-only inference. In Proceedings of the IEEE Conference on Computer Vision and Pattern Recognition (pp. 2704–2713).",
    "Kingma, D. P., & Ba, J. (2015). Adam: A method for stochastic optimization. In International Conference on Learning Representations (ICLR).",
    "Kneusel, R. T. (2021). Practical deep learning: A Python-based introduction. No Starch Press.",
    "LeCun, Y., Bottou, L., Bengio, Y., & Haffner, P. (1998). Gradient-based learning applied to document recognition. Proceedings of the IEEE, 86(11), 2278–2324.",
    "Prechelt, L. (1998). Early stopping – but when? In G. B. Orr & K.-R. Müller (Eds.), Neural networks: Tricks of the trade (pp. 55–69). Springer.",
    "Rosenblatt, F. (1958). The perceptron: A probabilistic model for information storage and organization in the brain. Psychological Review, 65(6), 386–408.",
    "Rumelhart, D. E., Hinton, G. E., & Williams, R. J. (1986). Learning representations by back-propagating errors. Nature, 323(6088), 533–536.",
    "Selvaraju, R. R., Cogswell, M., Das, A., Vedantam, R., Parikh, D., & Batra, D. (2017). Grad-CAM: Visual explanations from deep networks via gradient-based localization. In Proceedings of the IEEE International Conference on Computer Vision (pp. 618–626).",
    "Shorten, C., & Khoshgoftaar, T. M. (2019). A survey on image data augmentation for deep learning. Journal of Big Data, 6(1), 60.",
    "Simonyan, K., & Zisserman, A. (2015). Very deep convolutional networks for large-scale image recognition. In International Conference on Learning Representations (ICLR).",
    "Sokolova, M., & Lapalme, G. (2009). A systematic analysis of performance measures for classification tasks. Information Processing & Management, 45(4), 427–437.",
    "Srivastava, N., Hinton, G., Krizhevsky, A., Sutskever, I., & Salakhutdinov, R. (2014). Dropout: A simple way to prevent neural networks from overfitting. Journal of Machine Learning Research, 15(56), 1929–1958.",
    "World Health Organization. (2021). World report on hearing. WHO.",
    "Zhang, A., Lipton, Z. C., Li, M., & Smola, A. J. (2023). Dive into deep learning. Cambridge University Press. https://d2l.ai",
]
for ref in REF_LIST:
    p = doc.add_paragraph(ref); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Cm(1.0); p.paragraph_format.first_line_indent = Cm(-1.0)
    p.paragraph_format.line_spacing = 1.0; p.paragraph_format.space_after = Pt(6)

# ============================================================================= BAGIAN B
H1("BAGIAN B. TAUTAN")
P("Semua tautan berikut dapat dibuka tanpa meminta akses. Laporan, source code, dan video membahas implementasi yang "
  "sama, dengan dataset, arsitektur, dan angka hasil evaluasi yang konsisten (seluruh angka laporan dibuat otomatis dari "
  "results.json hasil running notebook).")
LIST([("Link source code (GitHub): ", link_or_isi("github", "link repositori GitHub")),
      ("Link video demo beserta tahapan deployment: ", link_or_isi("video", "link YouTube/Google Drive (akses publik)")),
      ("Link aplikasi yang sudah di-deploy: ", link_or_isi("app", "link Hugging Face Space"))])

doc.save(OUT)
print(f"Laporan tersimpan: {OUT}  (mode: {'DRAF' if DRAFT else ('QUICK' if QUICK else 'FINAL')})")
