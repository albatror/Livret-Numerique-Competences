import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, colorchooser
from tkinter import font as tkfont
from PIL import Image, ImageTk
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
from pptx.enum.shapes import MSO_SHAPE
import json
import os
from collections import OrderedDict

# ===================== Configuration =====================

APP_TITLE = "Compétences Pro Ultimate"
MAX_LINES_PER_SLIDE = 20   # lignes (en-têtes + compétences) par diapositive
LEFT_PANEL_WIDTH = 420     # largeur colonne "Compétences disponibles"
COMP_LISTBOX_WIDTH = 62    # largeur listbox (caractères)
PREVIEW_WIDTH = 900
PREVIEW_HEIGHT = 520
HEADER_HEIGHT = 48         # hauteur bandeau domaine dans l'aperçu
SUBHEADER_SPACING = 8      # espacement après un sous-domaine
LINE_SPACING = 6           # espacement entre lignes de compétences
COVER_HEADER_HEIGHT = 64   # bandeau gris couverture (mini-apercu)
TEXT_MARGIN_X = 24
TEXT_MARGIN_Y = 18
DEFAULT_BODY_FONT = ("Arial", 12)
DEFAULT_TITLE_SIZE_PT = 20
DEFAULT_BODY_SIZE_PT = 12
DEFAULT_SUBHEADER_BOLD = True
DEFAULT_SUBHEADER_UNDERLINE = True
DEFAULT_TITLE_FG = "white"
COVER_HEADER_COLOR = "#6e6e6e"  # gris bandeau couverture

# Palette de couleurs pour domaines (assignation automatique)
DOMAIN_COLORS = [
    "#2E86C1", "#AF7AC5", "#48C9B0", "#F5B041", "#EC7063",
    "#16A085", "#5D6D7E", "#CA6F1E", "#7D3C98", "#1F618D"
]

# Sections (maternelle)
SECTION_KEYS = ["TPS", "PS", "MS", "GS"]
SECTION_LABELS = {
    "TPS": "TOUTE PETITE SECTION",
    "PS": "PETITE SECTION",
    "MS": "MOYENNE SECTION",
    "GS": "GRANDE SECTION",
}
# Champs par section
SECTION_FIELDS = OrderedDict([
    ("annee", "Année scolaire"),
    ("ecole", "École"),
    ("enseignants", "Enseignant(s)"),
])

# ===================== Structures de données =====================

class DomainState:
    def __init__(self, name, color):
        self.name = name
        self.color = color  # hex
        # images: list of dicts {path, pil, tk, pos[x,y], size[w,h]}
        self.images = []
        self.font_body = DEFAULT_BODY_FONT


class CompetenceItem:
    def __init__(self, domain, subdomain, text):
        self.domain = domain
        self.subdomain = subdomain
        self.text = text

    def key(self):
        return (self.domain, self.subdomain or "", self.text)


# ===================== Application =====================

class CompetenceApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)

        # Etat
        self.available = OrderedDict()  # domain -> OrderedDict{subdomain -> [competences]}
        self.domain_order = []
        self.domain_states = {}         # domain -> DomainState
        self.selected_items = []        # list[CompetenceItem]
        self.added_set = set()          # keys pour anti-doublon

        # Aperçu global (toutes pages de tous domaines)
        self.domain_page_map = {}       # domain -> list[page] ; page = list[(is_header, subdomain, item|None)]
        self.item_page_index = {}       # item.key() -> (domain, page_index)
        self.flat_pages = []            # list of (domain, page_index) pour navigation globale
        self.current_flat_index = 0     # index dans flat_pages
        self.current_domain = None      # domaine actif (déduit de flat_pages)

        # Infos couverture (uniquement personnelles)
        self.nom_var = tk.StringVar()
        self.prenom_var = tk.StringVar()
        self.naissance_var = tk.StringVar()
        self.photo_path = None
        self.personal_completed = False

        # Sections: données et widgets (3 champs par section)
        self.sections_data = {
            key: {
                "completed": False,
                "fields": {fname: "" for fname in SECTION_FIELDS.keys()}
            } for key in SECTION_KEYS
        }
        # key -> { 'entries': {fname:(Entry,Var)}, 'completed_var': tk.BooleanVar }
        self.sections_widgets = {}

        # Pour mesure du texte
        self.measure_font = tkfont.Font(family="Arial", size=12)

        # Drag/Resize images
        self.drag_data = {"x": 0, "y": 0, "image_index": None}
        self.resize_data = {"image_index": None, "start_x": 0, "start_y": 0}

        # UI
        self._build_ui()
        self.update_cover_preview()
        self.rebuild_pages_and_refresh()

    # -------------------- UI --------------------

    def _build_ui(self):
        self.root.geometry("1560x980")

        # Ligne haute: informations personnelles + Sections onglets
        top = ttk.Frame(self.root)
        top.pack(fill="x", padx=8, pady=6)

        # Section: Informations personnelles
        pers = ttk.LabelFrame(top, text="Informations personnelles")
        pers.pack(side="left", padx=6, pady=4, fill="x", expand=True)

        ttk.Label(pers, text="Nom:").grid(row=0, column=0, sticky="w")
        ttk.Entry(pers, textvariable=self.nom_var, width=18).grid(row=0, column=1, sticky="w", padx=4)

        ttk.Label(pers, text="Prénom:").grid(row=0, column=2, sticky="w")
        ttk.Entry(pers, textvariable=self.prenom_var, width=18).grid(row=0, column=3, sticky="w", padx=4)

        ttk.Label(pers, text="Date de naissance:").grid(row=0, column=4, sticky="w")
        ttk.Entry(pers, textvariable=self.naissance_var, width=16).grid(row=0, column=5, sticky="w", padx=4)

        ttk.Button(pers, text="Importer photo", command=self._import_photo).grid(row=0, column=6, padx=6)
        ttk.Button(pers, text="Marquer comme complétée", command=self._mark_personal_completed).grid(row=0, column=7, padx=6)

        # Bloc Sections (TPS/PS/MS/GS) via Notebook
        sections_block = ttk.LabelFrame(self.root, text="Sections de cycle (mémorisées indépendamment)")
        sections_block.pack(fill="x", padx=8, pady=4)

        self.sections_nb = ttk.Notebook(sections_block)
        self.sections_nb.pack(fill="x", padx=6, pady=6)

        for key in SECTION_KEYS:
            self._build_section_tab(key)

        # Zone principale
        main = ttk.Frame(self.root)
        main.pack(fill="both", expand=True, padx=8, pady=6)

        # Gauche: Disponibles (arbre + liste)
        left = ttk.LabelFrame(main, text="Compétences disponibles")
        left.pack(side="left", fill="y", padx=6, pady=4)

        left_tree_frame = ttk.Frame(left)
        left_tree_frame.pack(side="left", fill="y", padx=4, pady=4)
        self.tree = ttk.Treeview(left_tree_frame, show="tree", height=22)
        self.tree.pack(side="left", fill="y")
        self.tree.column("#0", width=LEFT_PANEL_WIDTH, minwidth=300, stretch=True)
        yscroll_tree = ttk.Scrollbar(left_tree_frame, orient="vertical", command=self.tree.yview)
        yscroll_tree.pack(side="left", fill="y")
        self.tree.configure(yscrollcommand=yscroll_tree.set)

        mid_left = ttk.Frame(left)
        mid_left.pack(side="left", fill="both", padx=8, pady=4)
        ttk.Label(mid_left, text="Compétences du sous-domaine").pack(anchor="w")
        self.comps_list = tk.Listbox(mid_left, width=COMP_LISTBOX_WIDTH, height=20, selectmode=tk.EXTENDED)
        self.comps_list.pack(fill="both", expand=True)
        yscroll_comp = ttk.Scrollbar(mid_left, orient="vertical", command=self.comps_list.yview)
        yscroll_comp.pack(side="right", fill="y")
        self.comps_list.configure(yscrollcommand=yscroll_comp.set)

        btns_left = ttk.Frame(mid_left)
        btns_left.pack(fill="x", pady=6)
        ttk.Button(btns_left, text="Charger COMPETENCES.txt", command=self.load_competences_file).pack(side="left", padx=2)
        ttk.Button(btns_left, text="Ajouter ->", command=self.add_selected_competences).pack(side="left", padx=2)

        # Centre: Sélectionnées
        center = ttk.LabelFrame(main, text="Sélectionnées (dans le PPT)")
        center.pack(side="left", fill="y", padx=6, pady=4)

        self.selected_tree = ttk.Treeview(center, columns=("subdomain", "text"), show="headings", height=22)
        self.selected_tree.heading("subdomain", text="Sous-domaine")
        self.selected_tree.heading("text", text="Compétence")
        self.selected_tree.column("subdomain", width=200, stretch=True)
        self.selected_tree.column("text", width=460, stretch=True)
        self.selected_tree.pack(side="left", fill="both", expand=True, padx=4, pady=4)

        yscroll_sel = ttk.Scrollbar(center, orient="vertical", command=self.selected_tree.yview)
        yscroll_sel.pack(side="left", fill="y")
        self.selected_tree.configure(yscrollcommand=yscroll_sel.set)

        btns_center = ttk.Frame(center)
        btns_center.pack(side="left", fill="y", padx=6)
        ttk.Button(btns_center, text="Retirer <-", command=self.remove_selected_from_ppt).pack(pady=4, fill="x")
        ttk.Button(btns_center, text="Aller à la page", command=self.goto_selected_page).pack(pady=4, fill="x")
        ttk.Button(btns_center, text="Ajouter image (domaine)", command=self.add_image_domain).pack(pady=12, fill="x")
        ttk.Button(btns_center, text="Police/Couleur (domaine)", command=self.change_font_color).pack(pady=4, fill="x")
        ttk.Button(btns_center, text="Sauvegarder projet", command=self.save_project).pack(pady=12, fill="x")
        ttk.Button(btns_center, text="Charger projet", command=self.load_project).pack(pady=4, fill="x")
        ttk.Button(btns_center, text="Exporter PowerPoint", command=self.export_ppt).pack(pady=16, fill="x")

        # Droite: Mini couverture + Pagination globale + Aperçu
        right = ttk.Frame(main)
        right.pack(side="left", fill="both", expand=True, padx=6, pady=4)

        cover_frame = ttk.LabelFrame(right, text="Mini-apercu Page de garde")
        cover_frame.pack(fill="x")
        self.cover_canvas = tk.Canvas(cover_frame, width=PREVIEW_WIDTH, height=150, bg="white")
        self.cover_canvas.pack(fill="x")

        pager = ttk.Frame(right)
        pager.pack(fill="x", pady=6)
        self.btn_prev = ttk.Button(pager, text="◀ Page précédente", command=self.prev_page)
        self.btn_prev.pack(side="left", padx=4)
        self.page_var = tk.StringVar(value="Page 0/0")
        ttk.Label(pager, textvariable=self.page_var).pack(side="left", padx=10)
        self.btn_next = ttk.Button(pager, text="Page suivante ▶", command=self.next_page)
        self.btn_next.pack(side="left", padx=4)

        self.preview_frame = ttk.LabelFrame(right, text="Aperçu des pages (tous domaines)")
        self.preview_frame.pack(fill="both", expand=True)
        self.preview_canvas = tk.Canvas(self.preview_frame, width=PREVIEW_WIDTH, height=PREVIEW_HEIGHT, bg="white")
        self.preview_canvas.pack(fill="both", expand=True)

        # drag/resize images sur domaine courant
        self.preview_canvas.bind("<Button-1>", self.start_drag)
        self.preview_canvas.bind("<B1-Motion>", self.drag_image)
        self.preview_canvas.bind("<Button-3>", self.start_resize)
        self.preview_canvas.bind("<B3-Motion>", self.resize_image)

        # Events
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)

    def _build_section_tab(self, key):
        title = SECTION_LABELS[key]
        frame = ttk.Frame(self.sections_nb)
        self.sections_nb.add(frame, text=title)

        # Case à cocher "Section complétée"
        completed_var = tk.BooleanVar(value=False)
        chk = ttk.Checkbutton(frame, text="Marquer cette section comme complétée",
                              variable=completed_var, command=lambda k=key: self._toggle_section_completed(k))
        chk.grid(row=0, column=0, columnspan=2, sticky="w", pady=(6, 2))

        entries = {}
        r = 1
        for fname, flabel in SECTION_FIELDS.items():
            ttk.Label(frame, text=flabel + " :").grid(row=r, column=0, sticky="e", padx=(0, 6), pady=(6, 2))
            var = tk.StringVar()
            ent = ttk.Entry(frame, width=60, textvariable=var)
            ent.grid(row=r, column=1, sticky="we", pady=(6, 2))
            # Bind mise à jour
            def on_change(var=var, k=key, fn=fname):
                self.sections_data[k]["fields"][fn] = var.get().strip()
                # auto: tous les champs remplis -> completed True
                all_filled = all(self.sections_data[k]["fields"][f].strip() for f in SECTION_FIELDS.keys())
                self.sections_widgets[k]["completed_var"].set(all_filled)
                self.sections_data[k]["completed"] = all_filled
                self.update_cover_preview()
            var.trace_add("write", lambda *args, cb=on_change: cb())
            entries[fname] = (ent, var)
            r += 1

        # Boutons
        btns = ttk.Frame(frame)
        btns.grid(row=r, column=0, columnspan=2, sticky="w", pady=8)
        ttk.Button(btns, text="Marquer comme complétée",
                   command=lambda k=key: self._set_section_completed(k, True)).pack(side="left", padx=2)
        ttk.Button(btns, text="Décocher",
                   command=lambda k=key: self._set_section_completed(k, False)).pack(side="left", padx=2)
        ttk.Button(btns, text="Effacer le contenu",
                   command=lambda k=key: self._clear_section(k)).pack(side="left", padx=8)

        frame.grid_columnconfigure(1, weight=1)
        self.sections_widgets[key] = {
            "completed_var": completed_var,
            "entries": entries
        }

    # -------------------- Chargement / parsing --------------------

    def load_competences_file(self):
        path = filedialog.askopenfilename(
            title="Sélectionnez le fichier COMPETENCES.txt",
            filetypes=[("Fichiers texte", "*.txt")]
        )
        if not path:
            return

        self.available.clear()
        self.domain_order.clear()
        self.domain_states.clear()
        self.selected_items.clear()
        self.added_set.clear()
        self.domain_page_map.clear()
        self.item_page_index.clear()
        self.flat_pages.clear()
        self.current_flat_index = 0
        self.current_domain = None

        current_domain = None
        current_subdomain = None

        with open(path, "r", encoding="utf-8-sig") as f:
            for raw in f:
                line = raw.strip()
                if not line:
                    continue
                # Nettoyage unicode
                line = line.replace("\u202f", " ").replace("\u00a0", " ").replace("\u2019", "'")

                if line.startswith("##-Domaine"):
                    current_domain = line.replace("##-Domaine", "").strip()
                    if current_domain not in self.available:
                        self.available[current_domain] = OrderedDict()
                        self.domain_order.append(current_domain)
                        color = DOMAIN_COLORS[(len(self.domain_order)-1) % len(DOMAIN_COLORS)]
                        self.domain_states[current_domain] = DomainState(current_domain, color)
                    current_subdomain = None

                elif line.startswith("#-"):
                    sub = line.replace("#-", "").strip()
                    if sub.lower().startswith("sous-domaine:"):
                        sub = sub.split(":", 1)[1].strip()
                    current_subdomain = sub

                elif line.startswith("XX"):
                    comp = line.replace("XX", "").strip()
                    if current_domain is None:
                        current_domain = "Domaine"
                        if current_domain not in self.available:
                            self.available[current_domain] = OrderedDict()
                            self.domain_order.append(current_domain)
                            color = DOMAIN_COLORS[(len(self.domain_order)-1) % len(DOMAIN_COLORS)]
                            self.domain_states[current_domain] = DomainState(current_domain, color)
                    sd = current_subdomain if current_subdomain else current_domain
                    self.available[current_domain].setdefault(sd, [])
                    self.available[current_domain][sd].append(comp)

        self.build_available_tree()
        self.rebuild_pages_and_refresh()

    def build_available_tree(self):
        self.tree.delete(*self.tree.get_children())
        for domain in self.domain_order:
            d_id = self.tree.insert("", "end", text=domain, open=True)
            submap = self.available[domain]
            if not submap:
                self.tree.insert(d_id, "end", text=domain)
            else:
                for sub in submap.keys():
                    self.tree.insert(d_id, "end", text=sub)

    def on_tree_select(self, event):
        # Affiche les compétences du sous-domaine choisi (retire celles déjà ajoutées)
        self.comps_list.delete(0, tk.END)
        sel = self.tree.selection()
        if not sel:
            return
        item_id = sel[0]
        text = self.tree.item(item_id, "text")
        parent = self.tree.parent(item_id)

        if parent:
            domain = self.tree.item(parent, "text")
            sub = text
            comps = self.available.get(domain, {}).get(sub, [])
            for c in comps:
                key = (domain, sub or "", c)
                if key not in self.added_set:
                    self.comps_list.insert(tk.END, c)
        else:
            pass

    # -------------------- Ajout / retrait --------------------

    def add_selected_competences(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Info", "Sélectionnez un sous-domaine.")
            return
        item_id = sel[0]
        parent = self.tree.parent(item_id)
        if not parent:
            messagebox.showinfo("Info", "Sélectionnez un sous-domaine, pas un domaine.")
            return

        domain = self.tree.item(parent, "text")
        sub = self.tree.item(item_id, "text")
        chosen = [self.comps_list.get(i) for i in self.comps_list.curselection()]
        if not chosen:
            messagebox.showinfo("Info", "Sélectionnez au moins une compétence.")
            return

        added_any = False
        for comp in chosen:
            item = CompetenceItem(domain, sub, comp)
            if item.key() in self.added_set:
                continue
            self.selected_items.append(item)
            self.added_set.add(item.key())
            added_any = True

        if not added_any:
            return

        self.refresh_selected_tree()
        self.rebuild_pages_and_refresh()
        self.on_tree_select(None)

    def refresh_selected_tree(self):
        self.selected_tree.delete(*self.selected_tree.get_children())
        for it in self.selected_items:
            self.selected_tree.insert("", "end", values=(it.subdomain, it.text))

    def remove_selected_from_ppt(self):
        sel = self.selected_tree.selection()
        if not sel:
            return
        indices = []
        for iid in sel:
            sd, txt = self.selected_tree.item(iid, "values")
            for idx, it in enumerate(self.selected_items):
                if it.subdomain == sd and it.text == txt:
                    indices.append(idx)
                    break
        if not indices:
            return

        for idx in sorted(indices, reverse=True):
            k = self.selected_items[idx].key()
            if k in self.added_set:
                self.added_set.remove(k)
            self.selected_items.pop(idx)

        self.refresh_selected_tree()
        self.rebuild_pages_and_refresh()
        self.on_tree_select(None)

    def goto_selected_page(self):
        sel = self.selected_tree.selection()
        if not sel:
            return
        iid = sel[0]
        sd, txt = self.selected_tree.item(iid, "values")
        for it in self.selected_items:
            if it.subdomain == sd and it.text == txt:
                loc = self.item_page_index.get(it.key())
                if loc:
                    domain, page_idx = loc
                    for i, (d, p) in enumerate(self.flat_pages):
                        if d == domain and p == page_idx:
                            self.current_flat_index = i
                            self.current_domain = d
                            break
                    self.update_preview()
                return

    # -------------------- Images (domaine) --------------------

    def add_image_domain(self):
        # Ajoute des images au domaine actuellement visible dans l'aperçu
        if not self.flat_pages:
            messagebox.showinfo("Info", "Aucune page domaine n'est disponible.")
            return
        domain, _ = self.flat_pages[self.current_flat_index]
        ds = self.domain_states.get(domain)
        if not ds:
            return
        paths = filedialog.askopenfilenames(
            title="Sélectionnez une ou plusieurs images",
            filetypes=[("Images", "*.png;*.jpg;*.jpeg;*.bmp")]
        )
        if not paths:
            return
        for p in paths:
            try:
                pil = Image.open(p)
                pil.thumbnail((360, 360), Image.LANCZOS)
                tkimg = ImageTk.PhotoImage(pil)
                ds.images.append({
                    "path": p,
                    "pil": pil,
                    "tk": tkimg,
                    "pos": [60, HEADER_HEIGHT + 30],
                    "size": list(pil.size)
                })
            except Exception as e:
                messagebox.showerror("Image", f"Erreur avec {p}: {e}")
        self.update_preview()

    def change_font_color(self):
        if not self.flat_pages:
            return
        domain, _ = self.flat_pages[self.current_flat_index]
        ds = self.domain_states.get(domain)
        if not ds:
            return
        size = simpledialog.askinteger("Taille police corps", "Taille de la police (ex: 12):",
                                       initialvalue=ds.font_body[1])
        color = colorchooser.askcolor(title="Couleur du domaine (bandeau/titres)")[1]
        if size:
            ds.font_body = (ds.font_body[0], size)
        if color:
            ds.color = color
        self.update_preview()

    # Drag & drop / resize
    def start_drag(self, event):
        if not self.flat_pages:
            return
        domain, _ = self.flat_pages[self.current_flat_index]
        ds = self.domain_states.get(domain)
        if not ds:
            return
        for i, img in enumerate(ds.images):
            x, y = img["pos"]
            w, h = img["size"]
            if x <= event.x <= x + w and y <= event.y <= y + h:
                self.drag_data = {"image_index": i, "x": event.x, "y": event.y}
                return

    def drag_image(self, event):
        idx = self.drag_data.get("image_index")
        if idx is None or not self.flat_pages:
            return
        domain, _ = self.flat_pages[self.current_flat_index]
        ds = self.domain_states.get(domain)
        if not ds:
            return
        img = ds.images[idx]
        dx = event.x - self.drag_data["x"]
        dy = event.y - self.drag_data["y"]
        img["pos"][0] += dx
        img["pos"][1] += dy
        self.drag_data["x"] = event.x
        self.drag_data["y"] = event.y
        self.update_preview()

    def start_resize(self, event):
        if not self.flat_pages:
            return
        domain, _ = self.flat_pages[self.current_flat_index]
        ds = self.domain_states.get(domain)
        if not ds:
            return
        for i, img in enumerate(ds.images):
            x, y = img["pos"]
            w, h = img["size"]
            if x <= event.x <= x + w and y <= event.y <= y + h:
                self.resize_data = {"image_index": i, "start_x": event.x, "start_y": event.y}
                return

    def resize_image(self, event):
        idx = self.resize_data.get("image_index")
        if idx is None or not self.flat_pages:
            return
        domain, _ = self.flat_pages[self.current_flat_index]
        ds = self.domain_states.get(domain)
        if not ds:
            return
        img = ds.images[idx]
        dx = event.x - self.resize_data["start_x"]
        dy = event.y - self.resize_data["start_y"]
        new_w = max(30, img["size"][0] + dx)
        new_h = max(30, img["size"][1] + dy)
        try:
            pil = Image.open(img["path"]).resize((int(new_w), int(new_h)), Image.LANCZOS)
            img["pil"] = pil
            img["tk"] = ImageTk.PhotoImage(pil)
            img["size"] = [int(new_w), int(new_h)]
            self.resize_data["start_x"] = event.x
            self.resize_data["start_y"] = event.y
            self.update_preview()
        except Exception as e:
            messagebox.showerror("Redimensionner", str(e))

    # -------------------- Sections - logique --------------------

    def _toggle_section_completed(self, key):
        val = self.sections_widgets[key]["completed_var"].get()
        self.sections_data[key]["completed"] = bool(val)
        self.update_cover_preview()

    def _set_section_completed(self, key, value):
        self.sections_widgets[key]["completed_var"].set(bool(value))
        self.sections_data[key]["completed"] = bool(value)
        self.update_cover_preview()

    def _clear_section(self, key):
        for fname in SECTION_FIELDS.keys():
            ent, var = self.sections_widgets[key]["entries"][fname]
            var.set("")
        # recalcul auto completed
        self._recalc_section_completed(key)
        self.update_cover_preview()

    def _recalc_section_completed(self, key):
        all_filled = all(self.sections_data[key]["fields"][f].strip() for f in SECTION_FIELDS.keys())
        self.sections_widgets[key]["completed_var"].set(all_filled)
        self.sections_data[key]["completed"] = all_filled

    # -------------------- Couverture (aperçu mini) --------------------

    def _import_photo(self):
        p = filedialog.askopenfilename(
            title="Importer photo de l'élève",
            filetypes=[("Images", "*.png;*.jpg;*.jpeg;*.bmp")]
        )
        if p:
            self.photo_path = p
            self.update_cover_preview()

    def _mark_personal_completed(self):
        self.personal_completed = bool(
            self.nom_var.get().strip() and self.prenom_var.get().strip() and self.naissance_var.get().strip()
        )
        self.update_cover_preview()

    def update_cover_preview(self):
        c = self.cover_canvas
        c.delete("all")
        # Bandeau gris
        c.create_rectangle(0, 0, PREVIEW_WIDTH, COVER_HEADER_HEIGHT, fill=COVER_HEADER_COLOR, outline=COVER_HEADER_COLOR)
        c.create_text(PREVIEW_WIDTH // 2, COVER_HEADER_HEIGHT // 2, text="PROFIL DE L'ELEVE", fill="white",
                      font=("Arial", 16, "bold"))

        # Zone texte à gauche (infos personnelles)
        x = TEXT_MARGIN_X
        y = COVER_HEADER_HEIGHT + 10
        lines = []
        if self.nom_var.get().strip():
            lines.append(f"Nom: {self.nom_var.get().strip()}")
        if self.prenom_var.get().strip():
            lines.append(f"Prénom: {self.prenom_var.get().strip()}")
        if self.naissance_var.get().strip():
            lines.append(f"Date de naissance: {self.naissance_var.get().strip()}")

        for line in lines:
            c.create_text(x, y, anchor="nw", text=line, font=("Arial", 11), fill="black")
            y += 18

        # Sections complétées (italique)
        sec_y = COVER_HEADER_HEIGHT + 10
        sec_x = PREVIEW_WIDTH - 520
        c.create_text(sec_x, sec_y, anchor="nw", text="Sections complétées:", font=("Arial", 11, "bold"))
        sec_y += 18
        if self.personal_completed:
            c.create_text(sec_x, sec_y, anchor="nw", text="Informations personnelles (Section complétée)",
                          font=("Arial", 11, "italic"))
            sec_y += 16
        for key in SECTION_KEYS:
            if self.sections_data[key]["completed"]:
                c.create_text(sec_x, sec_y, anchor="nw",
                              text=f"{SECTION_LABELS[key]} (Section complétée)", font=("Arial", 11, "italic"))
                sec_y += 16

        # Photo (si fournie) - mini-aperçu
        if self.photo_path and os.path.exists(self.photo_path):
            try:
                pil = Image.open(self.photo_path)
                pil.thumbnail((120, 120), Image.LANCZOS)
                tkimg = ImageTk.PhotoImage(pil)
                c.image = tkimg  # éviter le GC
                c.create_image(PREVIEW_WIDTH - 140, COVER_HEADER_HEIGHT + 8, anchor="nw", image=tkimg)
            except Exception:
                pass

    # -------------------- Pagination & Aperçu --------------------

    def rebuild_pages_and_refresh(self):
        # Regroupe par domaine/sous-domaine et découpe en pages
        self.domain_page_map.clear()
        self.item_page_index.clear()
        # ordre des domaines fixe
        by_domain = OrderedDict((d, OrderedDict()) for d in self.domain_order)
        for it in self.selected_items:
            d = it.domain
            sd = it.subdomain or d
            by_domain.setdefault(d, OrderedDict())
            by_domain[d].setdefault(sd, [])
            by_domain[d][sd].append(it)

        for d in self.domain_order:
            pages = []
            current_page = []
            line_count = 0
            submap = by_domain.get(d, {})

            for sd, items in submap.items():
                # header sd
                if line_count + 1 > MAX_LINES_PER_SLIDE and current_page:
                    pages.append(current_page)
                    current_page = []
                    line_count = 0
                current_page.append((True, sd, None))
                line_count += 1

                # items
                for it in items:
                    if line_count + 1 > MAX_LINES_PER_SLIDE and current_page:
                        pages.append(current_page)
                        current_page = []
                        line_count = 0
                        # Répéter le header du sd sur la nouvelle page
                        current_page.append((True, sd, None))
                        line_count += 1
                    current_page.append((False, sd, it))
                    line_count += 1

            if current_page:
                pages.append(current_page)

            self.domain_page_map[d] = pages

            # indexer items -> page
            for pi, page in enumerate(pages):
                for is_header, sd, payload in page:
                    if not is_header and payload is not None:
                        self.item_page_index[payload.key()] = (d, pi)

        # Construire la liste plate de navigation (tous domaines)
        self.flat_pages = []
        for d in self.domain_order:
            pages = self.domain_page_map.get(d, [])
            if not pages:
                continue
            for pi in range(len(pages)):
                self.flat_pages.append((d, pi))

        # Ajuster current_flat_index
        if not self.flat_pages:
            self.current_flat_index = 0
            self.current_domain = None
        else:
            if self.current_flat_index >= len(self.flat_pages):
                self.current_flat_index = len(self.flat_pages) - 1
            self.current_domain = self.flat_pages[self.current_flat_index][0]

        self.update_preview()

    def prev_page(self):
        if not self.flat_pages:
            return
        if self.current_flat_index > 0:
            self.current_flat_index -= 1
            self.current_domain = self.flat_pages[self.current_flat_index][0]
            self.update_preview()

    def next_page(self):
        if not self.flat_pages:
            return
        if self.current_flat_index < len(self.flat_pages) - 1:
            self.current_flat_index += 1
            self.current_domain = self.flat_pages[self.current_flat_index][0]
            self.update_preview()

    def update_preview(self):
        # Met à jour la mini-couverture + la page domaine courante
        self.update_cover_preview()
        c = self.preview_canvas
        c.delete("all")

        if not self.flat_pages:
            self.page_var.set("Page 0/0 — Aucune page (ajoutez des compétences)")
            c.create_text(PREVIEW_WIDTH//2, PREVIEW_HEIGHT//2, text="Aucune page à afficher",
                          font=("Arial", 14, "italic"), fill="#666")
            return

        d, pi = self.flat_pages[self.current_flat_index]
        pages = self.domain_page_map.get(d, [])
        total_pages = len(self.flat_pages)
        self.page_var.set(f"Page {self.current_flat_index + 1}/{total_pages} — Domaine: {d} — p.{pi + 1}/{len(pages)}")

        ds = self.domain_states[d]
        # Bandeau de domaine
        c.create_rectangle(0, 0, PREVIEW_WIDTH, HEADER_HEIGHT, fill=ds.color, outline=ds.color)
        c.create_text(TEXT_MARGIN_X, HEADER_HEIGHT // 2, anchor="w",
                      text=d, fill=DEFAULT_TITLE_FG, font=("Arial", 16, "bold"))

        # Contenu
        y = HEADER_HEIGHT + 14
        x = TEXT_MARGIN_X
        body_font = ("Arial", ds.font_body[1])
        max_text_width = PREVIEW_WIDTH - 2 * TEXT_MARGIN_X - 10

        page = pages[pi] if 0 <= pi < len(pages) else []
        first_para_drawn = False
        for is_header, sub, payload in page:
            if is_header:
                c.create_text(x, y, anchor="nw", text=sub, fill=ds.color,
                              font=("Arial", 13, "bold", "underline"))
                y += 22
            else:
                wrapped = self.wrap_text(payload.text, max_text_width, body_font)
                for li, line in enumerate(wrapped):
                    c.create_text(x + 16, y, anchor="nw",
                                  text=("• " + line if li == 0 else "  " + line),
                                  fill="black", font=body_font)
                    y += (ds.font_body[1] + LINE_SPACING)
                y += SUBHEADER_SPACING

        # Images du domaine
        for img in ds.images:
            c.create_image(img["pos"][0], img["pos"][1], image=img["tk"], anchor="nw")

    def wrap_text(self, text, max_width_px, font_tuple):
        f = tkfont.Font(family=font_tuple[0], size=font_tuple[1])
        words = text.split()
        lines = []
        cur = ""
        for w in words:
            test = w if not cur else cur + " " + w
            if f.measure(test) <= max_width_px:
                cur = test
            else:
                if cur:
                    lines.append(cur)
                cur = w
        if cur:
            lines.append(cur)
        return lines

    # -------------------- Sauvegarde / Chargement --------------------

    def save_project(self):
        path = filedialog.asksaveasfilename(title="Sauvegarder projet",
                                            defaultextension=".json",
                                            filetypes=[("JSON", "*.json")])
        if not path:
            return
        data = {
            "available": self.available,
            "domain_order": self.domain_order,
            "selected": [(it.domain, it.subdomain, it.text) for it in self.selected_items],
            "domains": {
                d: {
                    "color": self.domain_states[d].color,
                    "font_body": self.domain_states[d].font_body,
                    "images": [
                        {"path": im["path"], "pos": im["pos"], "size": im["size"]}
                        for im in self.domain_states[d].images
                    ]
                } for d in self.domain_order
            },
            "infos": {
                "nom": self.nom_var.get(),
                "prenom": self.prenom_var.get(),
                "naissance": self.naissance_var.get(),
                "photo": self.photo_path,
                "personal_completed": self.personal_completed
            },
            "sections": {
                key: {
                    "completed": self.sections_data[key]["completed"],
                    "fields": self.sections_data[key]["fields"]
                } for key in SECTION_KEYS
            }
        }
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            messagebox.showinfo("Sauvegarde", f"Projet sauvegardé : {path}")
        except Exception as e:
            messagebox.showerror("Sauvegarde", str(e))

    def load_project(self):
        path = filedialog.askopenfilename(title="Charger projet", filetypes=[("JSON", "*.json")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            messagebox.showerror("Chargement", str(e))
            return

        # Restaurer domaines/compétences
        self.available = OrderedDict()
        for d, submap in data.get("available", {}).items():
            self.available[d] = OrderedDict()
            for sd, lst in submap.items():
                self.available[d][sd] = list(lst)

        self.domain_order = data.get("domain_order", [])
        self.domain_states.clear()
        domdata = data.get("domains", {})
        for idx, d in enumerate(self.domain_order):
            color = domdata.get(d, {}).get("color", DOMAIN_COLORS[idx % len(DOMAIN_COLORS)])
            ds = DomainState(d, color)
            ds.font_body = tuple(domdata.get(d, {}).get("font_body", DEFAULT_BODY_FONT))
            for im in domdata.get(d, {}).get("images", []):
                p = im["path"]
                try:
                    pil = Image.open(p).resize((int(im["size"][0]), int(im["size"][1])), Image.LANCZOS)
                    tkimg = ImageTk.PhotoImage(pil)
                    ds.images.append({
                        "path": p, "pil": pil, "tk": tkimg,
                        "pos": [int(im["pos"][0]), int(im["pos"][1])],
                        "size": [int(im["size"][0]), int(im["size"][1])]
                    })
                except Exception:
                    pass
            self.domain_states[d] = ds

        self.selected_items = []
        self.added_set = set()
        for d, sd, txt in data.get("selected", []):
            it = CompetenceItem(d, sd, txt)
            self.selected_items.append(it)
            self.added_set.add(it.key())

        # Infos
        infos = data.get("infos", {})
        self.nom_var.set(infos.get("nom", ""))
        self.prenom_var.set(infos.get("prenom", ""))
        self.naissance_var.set(infos.get("naissance", ""))
        self.photo_path = infos.get("photo", None)
        self.personal_completed = bool(infos.get("personal_completed", False))

        # Sections
        secdata = data.get("sections", {})
        for key in SECTION_KEYS:
            sd = secdata.get(key, {})
            comp = bool(sd.get("completed", False))
            self.sections_data[key]["completed"] = comp
            if key in self.sections_widgets:
                self.sections_widgets[key]["completed_var"].set(comp)
            # fields
            fields = sd.get("fields", {})
            for fname in SECTION_FIELDS.keys():
                val = fields.get(fname, "")
                self.sections_data[key]["fields"][fname] = val
                if key in self.sections_widgets:
                    ent, var = self.sections_widgets[key]["entries"][fname]
                    var.set(val)

        self.build_available_tree()
        self.refresh_selected_tree()
        self.rebuild_pages_and_refresh()
        messagebox.showinfo("Chargement", "Projet chargé avec succès")

    # -------------------- Export PowerPoint --------------------

    def export_ppt(self):
        path = filedialog.asksaveasfilename(title="Créer PowerPoint",
                                            defaultextension=".pptx",
                                            filetypes=[("PowerPoint", "*.pptx")])
        if not path:
            return
        try:
            prs = Presentation()
            # Page 1: Couverture
            self.build_cover_slide(prs)

            # Pages domaines
            self.rebuild_pages_and_refresh()
            any_domain_page = False
            for d in self.domain_order:
                pages = self.domain_page_map.get(d, [])
                if not pages:
                    continue
                any_domain_page = True
                ds = self.domain_states[d]
                for page in pages:
                    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blanc
                    # Bandeau domaine
                    self.add_domain_banner(slide, prs, d, ds.color)
                    # Zone de texte
                    left = Inches(0.6)
                    top = Inches(1.2)
                    width = prs.slide_width - Inches(1.2)
                    height = prs.slide_height - Inches(1.7)
                    tf_shape = slide.shapes.add_textbox(left, top, width, height)
                    tf = tf_shape.text_frame
                    tf.clear()

                    first_written = False
                    # Rendu contenu
                    for is_header, sub, payload in page:
                        if is_header:
                            p = tf.paragraphs[0] if not first_written else tf.add_paragraph()
                            first_written = True
                            p.text = sub
                            p.font.size = Pt(DEFAULT_BODY_SIZE_PT + 1)
                            p.font.bold = DEFAULT_SUBHEADER_BOLD
                            p.font.underline = DEFAULT_SUBHEADER_UNDERLINE
                            r, g, b = self.hex_to_rgb(self.domain_states[d].color)
                            p.font.color.rgb = RGBColor(r, g, b)
                            p.space_after = Pt(2)
                        else:
                            p = tf.paragraphs[0] if not first_written else tf.add_paragraph()
                            first_written = True
                            p.text = f"• {payload.text}"
                            p.level = 1
                            p.font.size = Pt(DEFAULT_BODY_SIZE_PT)
                            p.font.bold = False
                            p.space_after = Pt(2)

                    # Exporter images du domaine (même placement relatif)
                    self.export_domain_images(slide, ds, prs)

            if not any_domain_page:
                # Pas de pages de domaines: on conserve la couverture seule
                pass

            prs.save(path)
            messagebox.showinfo("Succès", f"PowerPoint sauvegardé : {path}")
        except Exception as e:
            messagebox.showerror("Export", f"Échec de l'export : {e}")

    def build_cover_slide(self, prs):
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # blanc

        sw = prs.slide_width
        sh = prs.slide_height
        margin = Inches(0.6)
        header_h = Inches(0.8)
        content_top = header_h + Inches(0.15)

        # Bandeau gris haut avec titre UPPERCASE
        rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), sw, header_h)
        rect.fill.solid()
        r, g, b = self.hex_to_rgb(COVER_HEADER_COLOR)
        rect.fill.fore_color.rgb = RGBColor(r, g, b)
        rect.line.fill.background()

        tb = slide.shapes.add_textbox(Inches(0), Inches(0), sw, header_h)
        tf = tb.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.text = "PROFIL DE L'ELEVE"
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER

        # Mise en page: colonne gauche (infos personnelles) + colonne droite (photo + sections complétées)
        gap = Inches(0.2)
        right_w = Inches(3.6)  # largeur cible colonne droite
        right_w = min(right_w, sw - 2 * margin)  # sécurité
        right_left = sw - margin - right_w

        left_left = margin
        left_w = right_left - left_left - gap
        if left_w < Inches(3.0):
            # Si trop étroit, bascule en mono-colonne: tout en largeur
            left_left = margin
            left_w = sw - 2 * margin
            right_w = 0

        # Colonne gauche: infos personnelles
        left_top = content_top
        left_h = Inches(2.6)
        tb2 = slide.shapes.add_textbox(left_left, left_top, left_w, left_h)
        tf2 = tb2.text_frame
        tf2.clear()
        lines = []
        if self.nom_var.get().strip():
            lines.append(f"Nom: {self.nom_var.get().strip()}")
        if self.prenom_var.get().strip():
            lines.append(f"Prénom: {self.prenom_var.get().strip()}")
        if self.naissance_var.get().strip():
            lines.append(f"Date de naissance: {self.naissance_var.get().strip()}")

        first = True
        for line in lines:
            if first:
                p = tf2.paragraphs[0]
                p.text = line
                first = False
            else:
                p = tf2.add_paragraph()
                p.text = line
            p.font.size = Pt(14)

        # Colonne droite: photo + "Sections complétées"
        sec_box_top = content_top  # sera recalé sous la photo si elle existe
        sec_box_h = Inches(2.6)

        if right_w > 0:
            # Photo en haut à droite, dans la page
            if self.photo_path and os.path.exists(self.photo_path):
                try:
                    max_photo_h = Inches(1.8)
                    pic = slide.shapes.add_picture(self.photo_path, Inches(0), Inches(0), height=max_photo_h)
                    # Recalage top/right
                    pic.left = right_left + right_w - pic.width
                    pic.top = content_top
                    sec_box_top = pic.top + pic.height + Inches(0.15)
                except Exception:
                    pass

            # Bloc "Sections complétées" sous la photo, à droite
            tb3 = slide.shapes.add_textbox(right_left, sec_box_top, right_w, sec_box_h)
            tf3 = tb3.text_frame
            tf3.clear()
            p = tf3.paragraphs[0]
            p.text = "Sections complétées:"
            p.font.size = Pt(14)
            p.font.bold = True

            if self.personal_completed:
                p = tf3.add_paragraph()
                p.text = "Informations personnelles (Section complétée)"
                p.font.size = Pt(12)
                p.font.italic = True

            for key in SECTION_KEYS:
                if self.sections_data[key]["completed"]:
                    p = tf3.add_paragraph()
                    p.text = f"{SECTION_LABELS[key]} (Section complétée)"
                    p.font.size = Pt(12)
                    p.font.italic = True

        # Détails des sections complétées: sous les colonnes, pleine largeur
        details_top = max(left_top + left_h, sec_box_top + sec_box_h if right_w > 0 else left_top + left_h) + Inches(0.25)
        details_h = max(Inches(1.5), sh - details_top - Inches(0.6))

        tb4 = slide.shapes.add_textbox(Inches(0.6), details_top, sw - Inches(1.2), details_h)
        tf4 = tb4.text_frame
        tf4.clear()
        any_detail = False
        for key in SECTION_KEYS:
            if not self.sections_data[key]["completed"]:
                continue
            any_detail = True
            # Titre de section
            p = tf4.add_paragraph()
            p.text = SECTION_LABELS[key]
            p.font.size = Pt(14)
            p.font.bold = True
            # Champs (les 3 spécifiques)
            for fname, flabel in SECTION_FIELDS.items():
                val = self.sections_data[key]["fields"].get(fname, "").strip()
                if not val:
                    continue
                p = tf4.add_paragraph()
                p.text = f"{flabel}: {val}"
                p.level = 1
                p.font.size = Pt(11)
        if not any_detail:
            p = tf4.paragraphs[0]
            p.text = " "
            p.font.size = Pt(1)

    def add_domain_banner(self, slide, prs, domain_name, color_hex):
        # Utilise les dimensions de la présentation (prs), pas slide.part.presentation
        sw = prs.slide_width

        rect = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0), Inches(0),
            sw, Inches(0.7)
        )
        rect.fill.solid()
        r, g, b = self.hex_to_rgb(color_hex)
        rect.fill.fore_color.rgb = RGBColor(r, g, b)
        rect.line.fill.background()

        tb = slide.shapes.add_textbox(
            Inches(0.4), Inches(0.05),
            sw - Inches(0.8), Inches(0.6)
        )
        tf = tb.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.text = domain_name
        p.font.size = Pt(DEFAULT_TITLE_SIZE_PT)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)

    def export_domain_images(self, slide, ds, prs):
        # Map coordonnées apercu -> slide
        sw = prs.slide_width
        sh = prs.slide_height
        for img in ds.images:
            lx = int(img["pos"][0] / PREVIEW_WIDTH * sw)
            ly = int(img["pos"][1] / PREVIEW_HEIGHT * sh)
            w = int(img["size"][0] / PREVIEW_WIDTH * sw)
            h = int(img["size"][1] / PREVIEW_HEIGHT * sh)
            try:
                if w > 0 and h > 0:
                    slide.shapes.add_picture(img["path"], lx, ly, width=w, height=h)
            except Exception:
                pass

    # -------------------- Utils --------------------

    @staticmethod
    def hex_to_rgb(hx):
        hx = hx.lstrip("#")
        return tuple(int(hx[i:i+2], 16) for i in (0, 2, 4))


# ===================== Lancement =====================

if __name__ == "__main__":
    root = tk.Tk()
    app = CompetenceApp(root)
    root.mainloop()