"""
Ejercicio 4: Diagrama de Dispersión y Regresión Lineal
Curso: Algoritmia y Programación 2026-1

Permite pegar datos de Excel o cargar un archivo .xlsx/.csv,
genera el diagrama de dispersión y calcula la regresión lineal
mostrando la fórmula y el valor R².
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from scipy import stats
import io


# ─── Paleta de colores ───────────────────────────────────────────────────────
BG_DARK      = "#0f1117"
BG_CARD      = "#1a1d2e"
BG_INPUT     = "#22253a"
ACCENT       = "#6c63ff"
ACCENT2      = "#ff6584"
TEXT_PRIMARY = "#e8eaf6"
TEXT_MUTED   = "#7c83a0"
BORDER       = "#2e3250"
SUCCESS      = "#43e97b"
WARNING      = "#f7971e"


class RegresionApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("📊 Diagrama de Dispersión & Regresión")
        self.geometry("1200x800")
        self.minsize(900, 650)
        self.configure(bg=BG_DARK)
        self.resizable(True, True)

        self.color_puntos = ACCENT
        self.color_linea = ACCENT2

        self._setup_styles()
        self._build_ui()

    # ── Estilos ──────────────────────────────────────────────────────────────
    def _setup_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background=BG_DARK)
        style.configure("Card.TFrame", background=BG_CARD, relief="flat")
        style.configure(
            "TButton",
            background=ACCENT, foreground="white",
            font=("Segoe UI", 10, "bold"),
            borderwidth=0, relief="flat", padding=(12, 7),
        )
        style.map("TButton",
                  background=[("active", "#8078ff"), ("pressed", "#5048cc")],
                  relief=[("pressed", "flat")])
        style.configure(
            "Secondary.TButton",
            background=BG_INPUT, foreground=TEXT_PRIMARY,
            font=("Segoe UI", 10),
            borderwidth=0, relief="flat", padding=(12, 7),
        )
        style.map("Secondary.TButton",
                  background=[("active", BORDER)])
        style.configure("TLabel", background=BG_DARK, foreground=TEXT_PRIMARY,
                        font=("Segoe UI", 10))
        style.configure("Title.TLabel", background=BG_DARK, foreground=TEXT_PRIMARY,
                        font=("Segoe UI", 15, "bold"))
        style.configure("Card.TLabel", background=BG_CARD, foreground=TEXT_PRIMARY,
                        font=("Segoe UI", 10))
        style.configure("Muted.TLabel", background=BG_CARD, foreground=TEXT_MUTED,
                        font=("Segoe UI", 9))
        style.configure("Stat.TLabel", background=BG_CARD, foreground=SUCCESS,
                        font=("Consolas", 12, "bold"))
        style.configure("Formula.TLabel", background=BG_CARD, foreground=ACCENT,
                        font=("Consolas", 13, "bold"))
        style.configure(
            "TCombobox",
            fieldbackground=BG_INPUT, background=BG_INPUT,
            foreground=TEXT_PRIMARY, selectbackground=ACCENT,
        )
        style.configure("TScrollbar", background=BG_INPUT, troughcolor=BG_DARK,
                        borderwidth=0, relief="flat")

    # ── Layout principal ──────────────────────────────────────────────────────
    def _build_ui(self):
        # Header
        header = tk.Frame(self, bg=BG_DARK, pady=12)
        header.pack(fill="x", padx=20)
        tk.Label(header, text="📊  Diagrama de Dispersión & Regresión",
                 bg=BG_DARK, fg=TEXT_PRIMARY,
                 font=("Segoe UI", 18, "bold")).pack(side="left")
        tk.Label(header, text="Algoritmia y Programación 2026-1",
                 bg=BG_DARK, fg=TEXT_MUTED,
                 font=("Segoe UI", 10)).pack(side="right", padx=4)

        # Panel principal (dos columnas)
        main = tk.Frame(self, bg=BG_DARK)
        main.pack(fill="both", expand=True, padx=20, pady=(0, 16))

        self._build_left_panel(main)
        self._build_right_panel(main)

    # ── Panel izquierdo: entrada de datos ────────────────────────────────────
    def _build_left_panel(self, parent):
        left_container = tk.Frame(parent, bg=BG_DARK, width=360)
        left_container.pack(side="left", fill="y", padx=(0, 12))
        left_container.pack_propagate(False)

        canvas = tk.Canvas(left_container, bg=BG_DARK, highlightthickness=0)
        scrollbar = ttk.Scrollbar(left_container, orient="vertical", command=canvas.yview)
        
        left = tk.Frame(canvas, bg=BG_DARK)
        
        left.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        frame_id = canvas.create_window((0, 0), window=left, anchor="nw")
        
        def _configure_canvas(e):
            canvas.itemconfig(frame_id, width=e.width)

        canvas.bind("<Configure>", _configure_canvas)
        canvas.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        def _bind_mouse(e):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            
        def _unbind_mouse(e):
            canvas.unbind_all("<MouseWheel>")
            
        canvas.bind("<Enter>", _bind_mouse)
        canvas.bind("<Leave>", _unbind_mouse)

        # Tarjeta: pegar datos
        self._section_label(left, "1  Pegar datos de Excel")
        paste_card = tk.Frame(left, bg=BG_CARD, bd=0,
                              highlightthickness=1, highlightbackground=BORDER)
        paste_card.pack(fill="x", pady=(4, 12))

        hint = tk.Label(paste_card,
                        text="Copia las celdas en Excel (Ctrl+C)\ny pégalas aquí (Ctrl+V).\nLa primera fila puede ser encabezado.",
                        bg=BG_CARD, fg=TEXT_MUTED,
                        font=("Segoe UI", 9), justify="left")
        hint.pack(anchor="w", padx=10, pady=(10, 4))

        self.txt_paste = tk.Text(
            paste_card, height=9, bg=BG_INPUT, fg=TEXT_PRIMARY,
            insertbackground=ACCENT, relief="flat", bd=0,
            font=("Consolas", 10), selectbackground=ACCENT,
            wrap="none",
        )
        self.txt_paste.pack(fill="both", padx=8, pady=(0, 8))

        btn_frame = tk.Frame(paste_card, bg=BG_CARD)
        btn_frame.pack(fill="x", padx=8, pady=(0, 10))

        btn_parse = tk.Button(
            btn_frame, text="⬆  Cargar datos",
            bg=ACCENT, fg="white", activebackground="#8078ff",
            font=("Segoe UI", 10, "bold"), relief="flat",
            cursor="hand2", command=self._load_from_paste, bd=0,
            padx=12, pady=7,
        )
        btn_parse.pack(side="left", expand=True, fill="x", padx=(0, 4))

        btn_clear = tk.Button(
            btn_frame, text="🗑  Limpiar",
            bg=BG_INPUT, fg=TEXT_PRIMARY, activebackground=BORDER,
            font=("Segoe UI", 10, "bold"), relief="flat",
            cursor="hand2", command=self._clear_paste, bd=0,
            padx=12, pady=7,
        )
        btn_clear.pack(side="right", expand=True, fill="x", padx=(4, 0))

        # Tarjeta: cargar archivo
        self._section_label(left, "2  Cargar archivo Excel / CSV")
        file_card = tk.Frame(left, bg=BG_CARD, bd=0,
                             highlightthickness=1, highlightbackground=BORDER)
        file_card.pack(fill="x", pady=(4, 12))

        btn_file = tk.Button(
            file_card, text="📂  Seleccionar archivo (.xlsx / .csv)",
            bg=BG_INPUT, fg=TEXT_PRIMARY, activebackground=BORDER,
            font=("Segoe UI", 10), relief="flat",
            cursor="hand2", command=self._load_from_file, bd=0,
            padx=12, pady=7,
        )
        btn_file.pack(fill="x", padx=8, pady=10)

        self.lbl_file = tk.Label(file_card, text="Sin archivo cargado",
                                 bg=BG_CARD, fg=TEXT_MUTED,
                                 font=("Segoe UI", 9))
        self.lbl_file.pack(padx=10, pady=(0, 8))

        # Tarjeta: selección de columnas
        self._section_label(left, "3  Seleccionar columnas")
        col_card = tk.Frame(left, bg=BG_CARD, bd=0,
                            highlightthickness=1, highlightbackground=BORDER)
        col_card.pack(fill="x", pady=(4, 12))

        tk.Label(col_card, text="Eje X (variable independiente):",
                 bg=BG_CARD, fg=TEXT_MUTED,
                 font=("Segoe UI", 9)).pack(anchor="w", padx=10, pady=(10, 2))
        self.combo_x = ttk.Combobox(col_card, state="readonly", font=("Segoe UI", 10))
        self.combo_x.pack(fill="x", padx=10, pady=(0, 6))

        tk.Label(col_card, text="Eje Y (variable dependiente):",
                 bg=BG_CARD, fg=TEXT_MUTED,
                 font=("Segoe UI", 9)).pack(anchor="w", padx=10, pady=(4, 2))
        self.combo_y = ttk.Combobox(col_card, state="readonly", font=("Segoe UI", 10))
        self.combo_y.pack(fill="x", padx=10, pady=(0, 6))

        # Tipo de regresión
        tk.Label(col_card, text="Tipo de regresión:",
                 bg=BG_CARD, fg=TEXT_MUTED,
                 font=("Segoe UI", 9)).pack(anchor="w", padx=10, pady=(4, 2))
        self.combo_reg = ttk.Combobox(
            col_card, state="readonly", font=("Segoe UI", 10),
            values=["Automática (Mejor ajuste)", "Lineal", "Cuadrática (polinómica grado 2)",
                    "Cúbica (polinómica grado 3)", "Exponencial", "Logarítmica"],
        )
        self.combo_reg.current(0)
        self.combo_reg.pack(fill="x", padx=10, pady=(0, 10))

        # Tarjeta: personalización
        self._section_label(left, "4  Personalización")
        cust_card = tk.Frame(left, bg=BG_CARD, bd=0,
                             highlightthickness=1, highlightbackground=BORDER)
        cust_card.pack(fill="x", pady=(4, 12))

        tk.Label(cust_card, text="Título (acepta $LaTeX$):",
                 bg=BG_CARD, fg=TEXT_MUTED,
                 font=("Segoe UI", 9)).pack(anchor="w", padx=10, pady=(10, 2))
        self.entry_title = ttk.Entry(cust_card, font=("Segoe UI", 10))
        self.entry_title.insert(0, "Diagrama de Dispersión")
        self.entry_title.pack(fill="x", padx=10, pady=(0, 6))

        tk.Label(cust_card, text="Posición Leyenda:",
                 bg=BG_CARD, fg=TEXT_MUTED,
                 font=("Segoe UI", 9)).pack(anchor="w", padx=10, pady=(4, 2))
        self.combo_legend = ttk.Combobox(
            cust_card, state="readonly", font=("Segoe UI", 10),
            values=["Automático (Mejor)", "Arriba Izquierda", "Arriba Derecha", 
                    "Abajo Izquierda", "Abajo Derecha", "Centro Derecha", "Centro Izquierda"]
        )
        self.combo_legend.current(0)
        self.combo_legend.pack(fill="x", padx=10, pady=(0, 6))

        tk.Label(cust_card, text="Tamaño Leyenda (pts):",
                 bg=BG_CARD, fg=TEXT_MUTED,
                 font=("Segoe UI", 9)).pack(anchor="w", padx=10, pady=(4, 2))
        self.combo_legend_size = ttk.Combobox(
            cust_card, state="readonly", font=("Segoe UI", 10),
            values=["8", "9", "10", "11", "12", "14", "16", "18", "20", "24"]
        )
        self.combo_legend_size.current(3) # Default 11
        self.combo_legend_size.pack(fill="x", padx=10, pady=(0, 6))

        color_frame = tk.Frame(cust_card, bg=BG_CARD)
        color_frame.pack(fill="x", padx=10, pady=(4, 10))

        self.btn_color_pts = tk.Button(
            color_frame, text="Color Puntos",
            bg=self.color_puntos, fg="white", font=("Segoe UI", 9, "bold"),
            relief="flat", cursor="hand2", command=self._choose_color_pts, bd=0, pady=4
        )
        self.btn_color_pts.pack(side="left", expand=True, fill="x", padx=(0, 4))

        self.btn_color_lin = tk.Button(
            color_frame, text="Color Línea",
            bg=self.color_linea, fg="white", font=("Segoe UI", 9, "bold"),
            relief="flat", cursor="hand2", command=self._choose_color_lin, bd=0, pady=4
        )
        self.btn_color_lin.pack(side="right", expand=True, fill="x", padx=(4, 0))

        # Botón generar
        btn_gen = tk.Button(
            left, text="✨  Generar Diagrama & Regresión",
            bg=ACCENT2, fg="white", activebackground="#ff8aa0",
            font=("Segoe UI", 11, "bold"), relief="flat",
            cursor="hand2", command=self._generate, bd=0,
            padx=12, pady=10,
        )
        btn_gen.pack(fill="x", pady=(4, 0))

        # Resultados estadísticos
        self._section_label(left, "Resultados")
        self.result_card = tk.Frame(left, bg=BG_CARD, bd=0,
                                    highlightthickness=1, highlightbackground=BORDER)
        self.result_card.pack(fill="x", pady=(4, 0))

        self.lbl_formula = tk.Label(self.result_card, text="—",
                                    bg=BG_CARD, fg=ACCENT,
                                    font=("Consolas", 11, "bold"),
                                    wraplength=290, justify="left")
        self.lbl_formula.pack(anchor="w", padx=10, pady=(10, 2))

        self.lbl_r2 = tk.Label(self.result_card, text="R² = —",
                               bg=BG_CARD, fg=SUCCESS,
                               font=("Consolas", 12, "bold"))
        self.lbl_r2.pack(anchor="w", padx=10, pady=(0, 4))

        self.lbl_r = tk.Label(self.result_card, text="r  = —",
                              bg=BG_CARD, fg=TEXT_MUTED,
                              font=("Consolas", 11))
        self.lbl_r.pack(anchor="w", padx=10)

        self.lbl_interp = tk.Label(self.result_card, text="",
                                   bg=BG_CARD, fg=TEXT_MUTED,
                                   font=("Segoe UI", 9), wraplength=290,
                                   justify="left")
        self.lbl_interp.pack(anchor="w", padx=10, pady=(4, 10))

        self.df = None

    def _section_label(self, parent, text):
        tk.Label(parent, text=text.upper(),
                 bg=BG_DARK, fg=TEXT_MUTED,
                 font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=(10, 2))

    def _choose_color_pts(self):
        color = colorchooser.askcolor(title="Elegir color de puntos", color=self.color_puntos)[1]
        if color:
            self.color_puntos = color
            self.btn_color_pts.config(bg=color)

    def _choose_color_lin(self):
        color = colorchooser.askcolor(title="Elegir color de línea", color=self.color_linea)[1]
        if color:
            self.color_linea = color
            self.btn_color_lin.config(bg=color)

    # ── Panel derecho: gráfica ────────────────────────────────────────────────
    def _build_right_panel(self, parent):
        right = tk.Frame(parent, bg=BG_CARD, bd=0,
                         highlightthickness=1, highlightbackground=BORDER)
        right.pack(side="left", fill="both", expand=True)

        # Placeholder
        self.placeholder = tk.Label(
            right,
            text="📉\n\nAquí aparecerá\nel diagrama de dispersión\ny la línea de regresión",
            bg=BG_CARD, fg=TEXT_MUTED,
            font=("Segoe UI", 14),
        )
        self.placeholder.pack(expand=True)

        self.graph_frame = right

        # Barra de herramientas placeholder
        self.toolbar_frame = tk.Frame(right, bg=BG_CARD)
        self.toolbar_frame.pack(side="bottom", fill="x")

        self.canvas_widget = None
        self.toolbar = None

    # ── Carga de datos ────────────────────────────────────────────────────────
    def _clear_paste(self):
        self.txt_paste.delete("1.0", "end")

    def _load_from_paste(self):
        raw = self.txt_paste.get("1.0", "end").strip()
        if not raw:
            messagebox.showwarning("Sin datos", "Pega los datos de Excel primero.")
            return
        try:
            self.df = pd.read_csv(io.StringIO(raw), sep="\t", engine="python")
            # Intentar con coma si solo da una columna
            if len(self.df.columns) == 1:
                self.df = pd.read_csv(io.StringIO(raw), sep=",", engine="python")
            self._update_combos()
            messagebox.showinfo("Datos cargados",
                                f"✅ {len(self.df)} filas × {len(self.df.columns)} columnas.")
        except Exception as e:
            messagebox.showerror("Error al parsear", str(e))

    def _load_from_file(self):
        path = filedialog.askopenfilename(
            title="Seleccionar archivo",
            filetypes=[("Excel / CSV", "*.xlsx *.xls *.csv"), ("Todos", "*.*")],
        )
        if not path:
            return
        try:
            if path.lower().endswith(".csv"):
                self.df = pd.read_csv(path)
            else:
                self.df = pd.read_excel(path)
            self._update_combos()
            self.lbl_file.config(text=f"📄 {path.split('/')[-1].split(chr(92))[-1]}",
                                 fg=SUCCESS)
            messagebox.showinfo("Archivo cargado",
                                f"✅ {len(self.df)} filas × {len(self.df.columns)} columnas.")
        except Exception as e:
            messagebox.showerror("Error al leer archivo", str(e))

    def _update_combos(self):
        # Solo columnas numéricas
        numeric_cols = list(self.df.select_dtypes(include=[np.number]).columns)
        all_cols = list(self.df.columns)
        opts = all_cols  # permitir todas; filtraremos al generar

        self.combo_x["values"] = opts
        self.combo_y["values"] = opts

        if len(opts) >= 2:
            self.combo_x.current(0)
            self.combo_y.current(1)
        elif len(opts) == 1:
            self.combo_x.current(0)

    # ── Generación del gráfico ────────────────────────────────────────────────
    def _generate(self):
        if self.df is None:
            messagebox.showwarning("Sin datos", "Primero carga o pega datos.")
            return

        col_x = self.combo_x.get()
        col_y = self.combo_y.get()

        if not col_x or not col_y:
            messagebox.showwarning("Columnas", "Selecciona las columnas X e Y.")
            return
        if col_x == col_y:
            messagebox.showwarning("Columnas", "Elige columnas distintas para X e Y.")
            return

        try:
            # Reemplazar comas por puntos en caso de usar formato español
            x_series = self.df[col_x].astype(str).str.replace(',', '.', regex=False)
            y_series = self.df[col_y].astype(str).str.replace(',', '.', regex=False)
            x = pd.to_numeric(x_series, errors="coerce")
            y = pd.to_numeric(y_series, errors="coerce")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        mask = x.notna() & y.notna()
        x, y = x[mask].values, y[mask].values

        if len(x) < 2:
            messagebox.showerror("Datos insuficientes",
                                 "Se necesitan al menos 2 puntos numéricos válidos.\n"
                                 "Verifica que las columnas contengan números.")
            return

        reg_type = self.combo_reg.get()
        self._plot(x, y, col_x, col_y, reg_type)

    def _plot(self, x, y, label_x, label_y, reg_type):
        # ── Matplotlib white style for reports ─────────────────────────────
        plt.rcParams.update({
            "figure.facecolor": "white",
            "axes.facecolor": "white",
            "axes.edgecolor": "black",
            "axes.labelcolor": "black",
            "xtick.color": "black",
            "ytick.color": "black",
            "grid.color": "#e0e0e0",
            "text.color": "black",
            "legend.facecolor": "white",
            "legend.edgecolor": "black",
        })

        fig, ax = plt.subplots(figsize=(8, 5.5), dpi=100)

        # Scatter
        ax.scatter(x, y, color=self.color_puntos, edgecolors="black",
                   linewidths=0.6, s=70, zorder=3, alpha=0.85, label="Datos")

        # ── Calcular regresión ─────────────────────────────────────────────
        x_fit = np.linspace(x.min(), x.max(), 300)
        
        if reg_type == "Automática (Mejor ajuste)":
            best_r2 = -1.0
            best_data = None
            for t in ["Lineal", "Cuadrática (polinómica grado 2)", "Cúbica (polinómica grado 3)", "Exponencial", "Logarítmica"]:
                data = self._get_regression_data(x, y, x_fit, t)
                if data[2] > best_r2:
                    best_r2 = data[2]
                    best_data = data
            if best_data is None or best_data[0] is None:
                messagebox.showerror("Error", "No se pudo calcular ninguna regresión.")
                return
            formula_str, latex_str, r2, r, y_fit, label, color = best_data
            label += " (Mejor automática)"
        else:
            formula_str, latex_str, r2, r, y_fit, label, color = self._get_regression_data(x, y, x_fit, reg_type)
            if formula_str is None:
                messagebox.showerror("Error", f"No se pudo calcular la regresión {reg_type}. Revisa que los datos sean válidos (ej. no valores <= 0 para log/exp).")
                return

        ax.plot(x_fit, y_fit, color=self.color_linea, lw=2.5, label=label)

        # Anotaciones en la gráfica
        ax.set_xlabel(label_x, fontsize=11)
        ax.set_ylabel(label_y, fontsize=11)
        
        title_text = self.entry_title.get()
        if not title_text.strip():
            title_text = f"Diagrama de Dispersión: {label_x} vs {label_y}"
        ax.set_title(title_text, fontsize=13, fontweight="bold", pad=14)
        
        ax.grid(True, linestyle="--", alpha=0.4)

        # Incorporar la fórmula y R² a la leyenda para que matplotlib use loc="best"
        # y ubique la caja en el lugar más despejado, evitando chocar con la línea o los puntos.
        import matplotlib.patches as mpatches
        text_content = f"{latex_str}\n$R^2 = {r2:.4f}$"
        dummy_patch = mpatches.Rectangle((0,0), 1, 1, fill=False, edgecolor='none', visible=False)
        
        handles, labels = ax.get_legend_handles_labels()
        handles.append(dummy_patch)
        labels.append(text_content)
        
        legend_locs = {
            "Automático (Mejor)": "best",
            "Arriba Izquierda": "upper left",
            "Arriba Derecha": "upper right",
            "Abajo Izquierda": "lower left",
            "Abajo Derecha": "lower right",
            "Centro Derecha": "center right",
            "Centro Izquierda": "center left"
        }
        loc_str = legend_locs.get(self.combo_legend.get(), "best")
        
        try:
            leg_size = int(self.combo_legend_size.get())
        except Exception:
            leg_size = 11
        
        leg = ax.legend(handles, labels, loc=loc_str, fontsize=leg_size, framealpha=0.9, edgecolor="black")
        leg.set_draggable(True)

        fig.tight_layout()

        # ── Mostrar en la UI ───────────────────────────────────────────────
        if self.canvas_widget:
            self.canvas_widget.get_tk_widget().destroy()
        if self.toolbar:
            self.toolbar.destroy()
        self.placeholder.pack_forget()

        canvas = FigureCanvasTkAgg(fig, master=self.graph_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

        toolbar = NavigationToolbar2Tk(canvas, self.toolbar_frame)
        toolbar.update()
        toolbar.configure(background=BG_CARD)

        self.canvas_widget = canvas
        self.toolbar = toolbar

        # ── Actualizar panel de resultados ─────────────────────────────────
        self.lbl_formula.config(text=f"f(x) = {formula_str.split('= ', 1)[-1]}")
        self.lbl_r2.config(text=f"R²  =  {r2:.6f}")
        self.lbl_r.config(text=f"r    =  {r:.6f}")
        self.lbl_interp.config(text=_interpret_r2(r2))

    def _get_regression_data(self, x, y, x_fit, reg_type):
        formula_str, latex_str, r2, r = "", "", -1.0, 0.0
        y_fit = None
        color = ACCENT2
        label = "Regresión"
        
        try:
            if reg_type == "Lineal":
                slope, intercept, r_val, p_val, std_err = stats.linregress(x, y)
                r, r2 = r_val, r_val**2
                y_fit = slope * x_fit + intercept
                sign = "+" if intercept >= 0 else "-"
                formula_str = (f"y = {slope:.4f}x {sign} {abs(intercept):.4f}")
                latex_str = f"$y = {slope:.4f}x {sign} {abs(intercept):.4f}$"
                color = ACCENT2
                label = "Regresión lineal"

            elif "Cuadrática" in reg_type:
                coeffs = np.polyfit(x, y, 2)
                r2 = _poly_r2(x, y, coeffs)
                r = r2**0.5
                y_fit = np.polyval(coeffs, x_fit)
                formula_str = _poly_str(coeffs, degree=2)
                latex_str = f"${formula_str}$"
                color = SUCCESS
                label = "Regresión cuadrática"

            elif "Cúbica" in reg_type:
                coeffs = np.polyfit(x, y, 3)
                r2 = _poly_r2(x, y, coeffs)
                r = r2**0.5
                y_fit = np.polyval(coeffs, x_fit)
                formula_str = _poly_str(coeffs, degree=3)
                latex_str = f"${formula_str}$"
                color = WARNING
                label = "Regresión cúbica"

            elif reg_type == "Exponencial":
                is_negative = False
                if np.all(y < 0):
                    is_negative = True
                    y_calc = -y
                elif np.any(y <= 0):
                    raise ValueError("Y contiene valores <= 0 mezclados (o ceros).")
                else:
                    y_calc = y

                log_y = np.log(y_calc)
                slope, intercept, r_val, _, _ = stats.linregress(x, log_y)
                r = r_val
                r2 = r_val**2
                a = np.exp(intercept)
                b = slope

                if is_negative:
                    a = -a

                y_fit = a * np.exp(b * x_fit)
                formula_str = f"y = {a:.4f} * e^({b:.4f}x)"
                latex_str = f"$y = {a:.4f} \\cdot e^{{{b:.4f}x}}$"
                color = "#f7c59f"
                label = "Regresión exponencial"

            elif reg_type == "Logarítmica":
                if np.any(x <= 0):
                    raise ValueError("X <= 0")
                log_x = np.log(x)
                slope, intercept, r_val, _, _ = stats.linregress(log_x, y)
                r = r_val
                r2 = r_val**2
                y_fit = slope * np.log(x_fit)  + intercept
                sign = "+" if intercept >= 0 else "-"
                formula_str = f"y = {slope:.4f}*ln(x) {sign} {abs(intercept):.4f}"
                latex_str = f"$y = {slope:.4f}\\ln(x) {sign} {abs(intercept):.4f}$"
                color = "#a8edea"
                label = "Regresión logarítmica"

            return formula_str, latex_str, r2, r, y_fit, label, color
        except Exception:
            return None, None, -1.0, 0.0, None, None, None


# ── Helpers ────────────────────────────────────────────────────────────────────
def _poly_r2(x, y, coeffs):
    y_pred = np.polyval(coeffs, x)
    ss_res = np.sum((y - y_pred) ** 2)
    ss_tot = np.sum((y - np.mean(y)) ** 2)
    return 1 - ss_res / ss_tot if ss_tot != 0 else 0.0


def _poly_str(coeffs, degree):
    terms = []
    for i, c in enumerate(coeffs):
        exp = degree - i
        sign = "+" if c >= 0 and i > 0 else ("-" if c < 0 and i > 0 else "")
        val = abs(c)
        if exp == 0:
            terms.append(f"{sign} {val:.4f}" if i > 0 else f"{c:.4f}")
        elif exp == 1:
            terms.append(f"{sign} {val:.4f}x" if i > 0 else f"{c:.4f}x")
        else:
            terms.append(f"{sign} {val:.4f}x^{exp}" if i > 0 else f"{c:.4f}x^{exp}")
    return "y = " + " ".join(terms)


def _interpret_r2(r2):
    if r2 >= 0.95:
        return "⭐ Ajuste excelente: el modelo explica ≥95% de la variación."
    elif r2 >= 0.80:
        return "✅ Buen ajuste: el modelo explica entre 80% y 95% de la variación."
    elif r2 >= 0.60:
        return "⚠️ Ajuste moderado: entre 60% y 80% de variación explicada."
    elif r2 >= 0.40:
        return "🔸 Ajuste débil: entre 40% y 60% de variación explicada."
    else:
        return "❌ Ajuste pobre: el modelo explica menos del 40% de la variación."


# ── Main ───────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = RegresionApp()
    app.mainloop()
