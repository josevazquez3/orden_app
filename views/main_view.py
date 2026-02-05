"""
Vista Principal con Tkinter
Sistema de Órdenes del Día
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
from PIL import Image, ImageTk
import os


class VentanaPrincipal(tk.Tk):
    """Ventana principal de la aplicación"""
    
    def __init__(self):
        try:
            print("[DEBUG] Iniciando VentanaPrincipal.__init__")
            super().__init__()
            
            # Configuración de la ventana
            self.title("Sistema de Órdenes del Día - Colegio de Médicos Prov. Buenos Aires")
            self.geometry("1200x800")
            
            # Colores corporativos
            self.color_verde = "#2E7D32"
            self.color_verde_claro = "#66BB6A"
            self.color_fondo = "#F5F5F5"
            self.color_blanco = "#FFFFFF"
            
            self.configure(bg=self.color_fondo)
            
            # Variables para imagen y texto editable
            self.imagen_logo = None
            self.texto_encabezado = tk.StringVar(value="REUNIÓN ORDINARIA")
            self.subtitulo_encabezado = tk.StringVar(value="COLEGIO DE MÉDICOS DE LA PROV. DE BUENOS AIRES")
            self.imagen_path = "assets/logo.png"
            self.tamaño_logo_ancho = tk.DoubleVar(value=3.5)  # cm para PDF
            self.tamaño_logo_alto = tk.DoubleVar(value=2.0)   # cm para PDF
            self.tamaño_logo_docx = tk.DoubleVar(value=1.2)   # pulgadas para DOCX
            self.tipo_reunion = tk.StringVar(value="presencial")  # Tipo de reunión
            self.plataforma = tk.StringVar(value="")  # Plataforma para reuniones virtuales
            self.tamaño_titulo = tk.IntVar(value=12)  # Tamaño de la letra del título
            self.fuente_titulo = tk.StringVar(value="Helvetica")  # Fuente del título
            self.negrita_titulo = tk.BooleanVar(value=True)  # Negrita del título
            self.negrita_subtitulo = tk.BooleanVar(value=True)  # Negrita del subtítulo
            
            print("[DEBUG] Variables inicializadas")
            
            # Cargar iconos para las pestañas
            self._cargar_iconos_pestañas()
            print("[DEBUG] Iconos de pestañas cargados")
            
            # Crear encabezado con canvas
            self._crear_encabezado()
            print("[DEBUG] Encabezado creado")
            
            # Crear el notebook (pestañas) con estilo personalizado
            self.notebook = ttk.Notebook(self)
            self._configurar_estilo_notebook()
            self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
            
            print("[DEBUG] Notebook creado")
            
            # Crear las pestañas
            self.crear_tab_reunion()
            print("[DEBUG] Tab reunión creado")
            self.crear_tab_temas()
            print("[DEBUG] Tab temas creado")
            self.crear_tab_delegados()
            print("[DEBUG] Tab delegados creado")
            self.crear_tab_historial()
            print("[DEBUG] Tab historial creado")
            print("[DEBUG] VentanaPrincipal.__init__ completado exitosamente")
            
        except Exception as e:
            print(f"[ERROR] Error en VentanaPrincipal.__init__: {str(e)}")
            import traceback
            traceback.print_exc()
            raise
    
    def _cargar_iconos_pestañas(self):
        """Carga los iconos para las pestañas del notebook"""
        try:
            # Tamaño más grande para los iconos de las pestañas (40x40 píxeles)
            tamaño_icono = (40, 40)
            ruta_base = "assets"
            
            # Icono Nueva Reunión
            if os.path.exists(os.path.join(ruta_base, "icono_nueva_reunion.png")):
                img = Image.open(os.path.join(ruta_base, "icono_nueva_reunion.png"))
                img = img.resize(tamaño_icono, Image.Resampling.LANCZOS)
                self.icono_pestaña_reunion = ImageTk.PhotoImage(img)
            else:
                self.icono_pestaña_reunion = None
            
            # Icono Gestión de Temas
            if os.path.exists(os.path.join(ruta_base, "icono_gestion_temas.png")):
                img = Image.open(os.path.join(ruta_base, "icono_gestion_temas.png"))
                img = img.resize(tamaño_icono, Image.Resampling.LANCZOS)
                self.icono_pestaña_temas = ImageTk.PhotoImage(img)
            else:
                self.icono_pestaña_temas = None
            
            # Icono Gestión de Delegados
            if os.path.exists(os.path.join(ruta_base, "icono_gestion_delegados.png")):
                img = Image.open(os.path.join(ruta_base, "icono_gestion_delegados.png"))
                img = img.resize(tamaño_icono, Image.Resampling.LANCZOS)
                self.icono_pestaña_delegados = ImageTk.PhotoImage(img)
            else:
                self.icono_pestaña_delegados = None
            
            # Icono Historial
            if os.path.exists(os.path.join(ruta_base, "icono_historial.png")):
                img = Image.open(os.path.join(ruta_base, "icono_historial.png"))
                img = img.resize(tamaño_icono, Image.Resampling.LANCZOS)
                self.icono_pestaña_historial = ImageTk.PhotoImage(img)
            else:
                self.icono_pestaña_historial = None
                
        except Exception as e:
            print(f"[ERROR] Error al cargar iconos de pestañas: {e}")
            self.icono_pestaña_reunion = None
            self.icono_pestaña_temas = None
            self.icono_pestaña_delegados = None
            self.icono_pestaña_historial = None
    
    def _configurar_estilo_notebook(self):
        """Configura el estilo del notebook para hacer las pestañas más grandes"""
        style = ttk.Style()
        # Configurar estilo para hacer las pestañas más grandes
        style.configure('TNotebook.Tab', 
                       padding=[25, 15],  # Padding más grande: horizontal y vertical
                       font=('Arial', 12, 'bold'),  # Fuente más grande y en negrita
                       height=50)  # Altura mínima de las pestañas
        
    def _crear_encabezado(self):
        """Crea encabezado decorativo con canvas"""
        header_frame = tk.Frame(self, bg=self.color_verde, height=80)
        header_frame.pack(fill='x', side='top')
        
        # Canvas para encabezado
        canvas_header = tk.Canvas(
            header_frame,
            bg=self.color_verde,
            height=80,
            highlightthickness=0
        )
        canvas_header.pack(fill='x')
        
        # Dibujar elementos decorativos
        canvas_header.create_line(0, 75, 1200, 75, fill=self.color_verde_claro, width=3)
        
        # Texto del encabezado
        canvas_header.create_text(
            600, 20,
            text="COLEGIO DE MÉDICOS DE LA PROVINCIA DE BUENOS AIRES",
            font=('Arial', 12, 'bold'),
            fill=self.color_blanco
        )
        
        canvas_header.create_text(
            600, 45,
            text="SISTEMA DE ÓRDENES DEL DÍA",
            font=('Arial', 10),
            fill=self.color_verde_claro
        )
    
    def crear_tab_reunion(self):
        """Crea el tab para nueva reunión"""
        frame = tk.Frame(self.notebook, bg=self.color_fondo)
        # Agregar pestaña con icono si está disponible
        if hasattr(self, 'icono_pestaña_reunion') and self.icono_pestaña_reunion:
            self.notebook.add(frame, text=" Nueva Reunión", image=self.icono_pestaña_reunion, compound='left')
        else:
            self.notebook.add(frame, text="📋 Nueva Reunión")
        
        # Crear canvas con scrollbar
        canvas = tk.Canvas(frame, bg=self.color_fondo, highlightthickness=0)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=self.color_fondo)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Frame para encabezado editable con logo
        self.titulo_frame = tk.Frame(scrollable_frame, bg=self.color_blanco, relief='flat', borderwidth=1)
        self.titulo_frame.pack(fill='x', padx=20, pady=15)
        
        # Frame para logo y controles
        frame_logo_controles = tk.Frame(self.titulo_frame, bg=self.color_blanco)
        frame_logo_controles.pack(fill='x', padx=10, pady=10)
        
        # Canvas para mostrar logo
        self.canvas_logo = tk.Canvas(
            frame_logo_controles,
            bg=self.color_blanco,
            width=80,
            height=80,
            highlightthickness=0
        )
        self.canvas_logo.pack(side='left', padx=10)
        
        # Frame para texto y botones
        frame_texto_botones = tk.Frame(frame_logo_controles, bg=self.color_blanco)
        frame_texto_botones.pack(side='left', fill='both', expand=True, padx=10)
        
        # Entry para editar texto
        tk.Label(frame_texto_botones, text="Texto del Encabezado:", bg=self.color_blanco, 
                font=('Arial', 10)).pack(anchor='w')
        self.entry_texto_encabezado = tk.Entry(frame_texto_botones, textvariable=self.texto_encabezado,
                                               font=('Arial', 10), width=50)
        self.entry_texto_encabezado.pack(anchor='w', fill='x', pady=5)
        
        # Entry para editar subtítulo
        tk.Label(frame_texto_botones, text="Subtítulo (opcional):", bg=self.color_blanco, 
                font=('Arial', 10)).pack(anchor='w')
        self.entry_subtitulo_encabezado = tk.Entry(frame_texto_botones, textvariable=self.subtitulo_encabezado,
                                                    font=('Arial', 10), width=50)
        self.entry_subtitulo_encabezado.pack(anchor='w', fill='x', pady=5)
        
        # Frame para formato del título
        frame_formato_titulo = tk.Frame(frame_texto_botones, bg=self.color_blanco)
        frame_formato_titulo.pack(anchor='w', fill='x', pady=5)
        
        tk.Label(frame_formato_titulo, text="Formato Título - Fuente:", bg=self.color_blanco, 
                font=('Arial', 9)).pack(side='left', padx=5)
        self.combo_fuente_titulo = ttk.Combobox(frame_formato_titulo, textvariable=self.fuente_titulo,
                                               values=["Helvetica", "Arial", "Times New Roman", "Courier", "Georgia"],
                                               state='readonly', width=15, font=('Arial', 8))
        self.combo_fuente_titulo.pack(side='left', padx=2)
        
        tk.Label(frame_formato_titulo, text="Tamaño:", bg=self.color_blanco, 
                font=('Arial', 9)).pack(side='left', padx=5)
        self.spinbox_tamaño_titulo = tk.Spinbox(frame_formato_titulo, from_=8, to=28, width=4, 
                                               textvariable=self.tamaño_titulo, font=('Arial', 8))
        self.spinbox_tamaño_titulo.pack(side='left', padx=2)
        
        # Checkbox para negrita del título
        self.check_negrita_titulo = tk.Checkbutton(frame_formato_titulo, text="Negrita Título",
                                                   variable=self.negrita_titulo, bg=self.color_blanco,
                                                   font=('Arial', 8))
        self.check_negrita_titulo.pack(side='left', padx=10)
        
        # Checkbox para negrita del subtítulo
        self.check_negrita_subtitulo = tk.Checkbutton(frame_formato_titulo, text="Negrita Subtítulo",
                                                      variable=self.negrita_subtitulo, bg=self.color_blanco,
                                                      font=('Arial', 8))
        self.check_negrita_subtitulo.pack(side='left', padx=2)
        
        # Frame para tamaño del logo
        frame_tamaño_logo = tk.Frame(frame_texto_botones, bg=self.color_blanco)
        frame_tamaño_logo.pack(anchor='w', fill='x', pady=5)
        
        tk.Label(frame_tamaño_logo, text="Tamaño Logo (PDF):", bg=self.color_blanco, 
                font=('Arial', 9)).pack(anchor='w')
        
        frame_tamaño_inputs = tk.Frame(frame_tamaño_logo, bg=self.color_blanco)
        frame_tamaño_inputs.pack(anchor='w', fill='x')
        
        tk.Label(frame_tamaño_inputs, text="Ancho (cm):", bg=self.color_blanco, 
                font=('Arial', 8)).pack(side='left', padx=5)
        self.spinbox_ancho_logo = tk.Spinbox(frame_tamaño_inputs, from_=0.5, to=10, width=6, 
                                            textvariable=self.tamaño_logo_ancho, font=('Arial', 8))
        self.spinbox_ancho_logo.pack(side='left', padx=2)
        
        tk.Label(frame_tamaño_inputs, text="Alto (cm):", bg=self.color_blanco, 
                font=('Arial', 8)).pack(side='left', padx=5)
        self.spinbox_alto_logo = tk.Spinbox(frame_tamaño_inputs, from_=0.5, to=10, width=6, 
                                           textvariable=self.tamaño_logo_alto, font=('Arial', 8))
        self.spinbox_alto_logo.pack(side='left', padx=2)
        
        tk.Label(frame_tamaño_inputs, text="Ancho (DOCX - pulgadas):", bg=self.color_blanco, 
                font=('Arial', 8)).pack(side='left', padx=5)
        self.spinbox_docx_logo = tk.Spinbox(frame_tamaño_inputs, from_=0.2, to=3, width=6, 
                                           textvariable=self.tamaño_logo_docx, font=('Arial', 8))
        self.spinbox_docx_logo.pack(side='left', padx=2)
        
        # Frame para botones
        frame_botones_logo = tk.Frame(frame_texto_botones, bg=self.color_blanco)
        frame_botones_logo.pack(anchor='w', fill='x')
        
        btn_cargar_logo = tk.Button(
            frame_botones_logo,
            text="📷 Cargar Logo",
            bg=self.color_verde,
            fg='white',
            font=('Arial', 9),
            padx=10,
            pady=5,
            cursor='hand2',
            command=self._cargar_logo
        )
        btn_cargar_logo.pack(side='left', padx=5)
        
        btn_actualizar_encabezado = tk.Button(
            frame_botones_logo,
            text="✔️ Actualizar",
            bg=self.color_verde_claro,
            fg='white',
            font=('Arial', 9),
            padx=10,
            pady=5,
            cursor='hand2',
            command=self._actualizar_encabezado_reunion
        )
        btn_actualizar_encabezado.pack(side='left', padx=5)
        
        # Canvas para mostrar el encabezado con líneas
        self.titulo_canvas = tk.Canvas(
            self.titulo_frame,
            bg=self.color_blanco,
            height=70,
            highlightthickness=0
        )
        self.titulo_canvas.pack(fill='x', expand=True)
        
        # Vinculación de eventos (después de usar after_idle para evitar loops)
        self.after_idle(lambda: self.titulo_canvas.bind("<Configure>", self._on_titulo_canvas_configure))
        self.after_idle(lambda: self.texto_encabezado.trace('w', self._on_texto_encabezado_cambio))
        
        # Cargar logo si existe (después de la inicialización completa)
        self.after_idle(self._cargar_logo_inicial)
        
        # Frame para datos de la reunión
        frame_datos = tk.LabelFrame(
            scrollable_frame,
            text="Datos de la Reunión",
            font=('Arial', 12, 'bold'),
            bg=self.color_blanco,
            fg=self.color_verde,
            padx=20,
            pady=15,
            relief='flat',
            borderwidth=2
        )
        frame_datos.pack(fill='x', padx=20, pady=10)
        
        # Fecha
        tk.Label(frame_datos, text="Fecha:", bg=self.color_blanco, 
                font=('Arial', 10)).grid(row=0, column=0, sticky='e', padx=5, pady=5)
        self.entry_fecha = tk.Entry(frame_datos, width=40, font=('Arial', 10))
        self.entry_fecha.grid(row=0, column=1, padx=5, pady=5, sticky='w')
        self.entry_fecha.insert(0, "viernes 23 de enero de 2026")
        
        # Hora
        tk.Label(frame_datos, text="Hora:", bg=self.color_blanco,
                font=('Arial', 10)).grid(row=1, column=0, sticky='e', padx=5, pady=5)
        self.entry_hora = tk.Entry(frame_datos, width=20, font=('Arial', 10))
        self.entry_hora.grid(row=1, column=1, padx=5, pady=5, sticky='w')
        self.entry_hora.insert(0, "17 hs.")
        
        # Lugar
        tk.Label(frame_datos, text="Lugar:", bg=self.color_blanco,
                font=('Arial', 10)).grid(row=2, column=0, sticky='e', padx=5, pady=5)
        self.entry_lugar = tk.Entry(frame_datos, width=50, font=('Arial', 10))
        self.entry_lugar.grid(row=2, column=1, padx=5, pady=5, sticky='w')
        self.entry_lugar.insert(0, "Calle 8 Nº 486 – La Plata")
        
        # Sede
        tk.Label(frame_datos, text="Sede (opcional):", bg=self.color_blanco,
                font=('Arial', 10)).grid(row=3, column=0, sticky='e', padx=5, pady=5)
        self.entry_sede = tk.Entry(frame_datos, width=50, font=('Arial', 10))
        self.entry_sede.grid(row=3, column=1, padx=5, pady=5, sticky='w')
        
        # Tipo
        tk.Label(frame_datos, text="Tipo:", bg=self.color_fondo,
                font=('Arial', 10)).grid(row=4, column=0, sticky='e', padx=5, pady=5)
        self.combo_tipo = ttk.Combobox(
            frame_datos, 
            values=["presencial", "virtual"],
            textvariable=self.tipo_reunion,
            state='readonly',
            width=37,
            font=('Arial', 10)
        )
        self.combo_tipo.grid(row=4, column=1, padx=5, pady=5, sticky='w')
        self.combo_tipo.current(0)
        self.combo_tipo.bind('<<ComboboxSelected>>', self._actualizar_visibilidad_plataforma)
        
        # Plataforma (solo para reuniones virtuales)
        self.label_plataforma = tk.Label(frame_datos, text="Plataforma:", bg=self.color_blanco,
                font=('Arial', 10))
        self.entry_plataforma = tk.Entry(frame_datos, width=50, font=('Arial', 10), 
                                        textvariable=self.plataforma)
        # Inicialmente ocultos
        self.label_plataforma.grid_remove()
        self.entry_plataforma.grid_remove()
        
        # Frame para delegados
        frame_delegados = tk.LabelFrame(
            scrollable_frame,
            text="Delegados Titulares",
            font=('Arial', 12, 'bold'),
            bg=self.color_fondo,
            fg=self.color_verde,
            padx=20,
            pady=15
        )
        frame_delegados.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Tabla de delegados
        columns = ('Título', 'Nombre', 'Apellido', 'Distrito')
        self.tree_delegados = ttk.Treeview(
            frame_delegados,
            columns=columns,
            show='headings',
            height=8
        )
        
        for col in columns:
            self.tree_delegados.heading(col, text=col)
            self.tree_delegados.column(col, width=150)
        
        scrollbar_del = ttk.Scrollbar(frame_delegados, orient="vertical", 
                                     command=self.tree_delegados.yview)
        self.tree_delegados.configure(yscrollcommand=scrollbar_del.set)
        
        self.tree_delegados.pack(side='left', fill='both', expand=True)
        scrollbar_del.pack(side='right', fill='y')
        
        # Botón editar delegado
        frame_btn_del = tk.Frame(scrollable_frame, bg=self.color_fondo)
        frame_btn_del.pack(fill='x', padx=20, pady=5)
        
        self.btn_editar_delegado = tk.Button(
            frame_btn_del,
            text="✏️ Editar Delegado Seleccionado",
            bg=self.color_verde_claro,
            fg='white',
            font=('Arial', 10),
            padx=15,
            pady=5,
            cursor='hand2'
        )
        self.btn_editar_delegado.pack(side='left', padx=5)
        
        self.btn_subir_delegado = tk.Button(
            frame_btn_del,
            text="⬆️ Subir Delegado",
            bg=self.color_verde_claro,
            fg='white',
            font=('Arial', 10),
            padx=15,
            pady=5,
            cursor='hand2'
        )
        self.btn_subir_delegado.pack(side='left', padx=5)
        
        self.btn_bajar_delegado = tk.Button(
            frame_btn_del,
            text="⬇️ Bajar Delegado",
            bg=self.color_verde_claro,
            fg='white',
            font=('Arial', 10),
            padx=15,
            pady=5,
            cursor='hand2'
        )
        self.btn_bajar_delegado.pack(side='left', padx=5)
        
        # Frame para orden del día
        frame_orden = tk.LabelFrame(
            scrollable_frame,
            text="Orden del Día",
            font=('Arial', 12, 'bold'),
            bg=self.color_fondo,
            fg=self.color_verde,
            padx=20,
            pady=15
        )
        frame_orden.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Botones para gestionar orden
        frame_botones_orden = tk.Frame(frame_orden, bg=self.color_fondo)
        frame_botones_orden.pack(fill='x', pady=5)
        
        self.btn_agregar_tema = tk.Button(
            frame_botones_orden,
            text="➕ Agregar Tema",
            bg=self.color_verde,
            fg='white',
            font=('Arial', 10, 'bold'),
            padx=15,
            pady=5,
            cursor='hand2'
        )
        self.btn_agregar_tema.pack(side='left', padx=5)
        
        self.btn_subir = tk.Button(
            frame_botones_orden,
            text="⬆️ Subir",
            bg=self.color_verde_claro,
            fg='white',
            font=('Arial', 10),
            padx=10,
            pady=5,
            cursor='hand2'
        )
        self.btn_subir.pack(side='left', padx=5)
        
        self.btn_bajar = tk.Button(
            frame_botones_orden,
            text="⬇️ Bajar",
            bg=self.color_verde_claro,
            fg='white',
            font=('Arial', 10),
            padx=10,
            pady=5,
            cursor='hand2'
        )
        self.btn_bajar.pack(side='left', padx=5)
        
        self.btn_eliminar = tk.Button(
            frame_botones_orden,
            text="🗑️ Eliminar",
            bg='#D32F2F',
            fg='white',
            font=('Arial', 10),
            padx=10,
            pady=5,
            cursor='hand2'
        )
        self.btn_eliminar.pack(side='left', padx=5)
        
        # Listbox para orden del día
        frame_listbox = tk.Frame(frame_orden, bg=self.color_fondo)
        frame_listbox.pack(fill='both', expand=True)
        
        scrollbar_orden = tk.Scrollbar(frame_listbox)
        scrollbar_orden.pack(side='right', fill='y')
        
        self.listbox_orden = tk.Listbox(
            frame_listbox,
            font=('Arial', 10),
            height=10,
            yscrollcommand=scrollbar_orden.set
        )
        self.listbox_orden.pack(side='left', fill='both', expand=True)
        scrollbar_orden.config(command=self.listbox_orden.yview)
        
        # Frame para firmas
        frame_firmas = tk.LabelFrame(
            scrollable_frame,
            text="Firmas",
            font=('Arial', 12, 'bold'),
            bg=self.color_fondo,
            fg=self.color_verde,
            padx=20,
            pady=15
        )
        frame_firmas.pack(fill='x', padx=20, pady=10)
        
        tk.Label(frame_firmas, text="Presidente:", bg=self.color_fondo,
                font=('Arial', 10)).grid(row=0, column=0, sticky='e', padx=5, pady=5)
        self.combo_presidente = ttk.Combobox(frame_firmas, state='readonly',
                                            width=40, font=('Arial', 10))
        self.combo_presidente.grid(row=0, column=1, padx=5, pady=5, sticky='w')
        
        tk.Label(frame_firmas, text="Secretario General:", bg=self.color_fondo,
                font=('Arial', 10)).grid(row=1, column=0, sticky='e', padx=5, pady=5)
        self.combo_secretario = ttk.Combobox(frame_firmas, state='readonly',
                                            width=40, font=('Arial', 10))
        self.combo_secretario.grid(row=1, column=1, padx=5, pady=5, sticky='w')
        
        # Botones finales
        frame_botones_final = tk.Frame(scrollable_frame, bg=self.color_fondo)
        frame_botones_final.pack(fill='x', padx=20, pady=20)
        
        self.btn_vista_previa = tk.Button(
            frame_botones_final,
            text="👁️ Vista Previa",
            bg='#1976D2',
            fg='white',
            font=('Arial', 12, 'bold'),
            padx=30,
            pady=10,
            cursor='hand2'
        )
        self.btn_vista_previa.pack(side='left', padx=10)
        
        self.btn_generar_pdf = tk.Button(
            frame_botones_final,
            text="📄 Generar PDF",
            bg=self.color_verde,
            fg='white',
            font=('Arial', 12, 'bold'),
            padx=30,
            pady=10,
            cursor='hand2'
        )
        self.btn_generar_pdf.pack(side='left', padx=10)
        
        self.btn_generar_doc = tk.Button(
            frame_botones_final,
            text="📝 Generar DOC",
            bg=self.color_verde,
            fg='white',
            font=('Arial', 12, 'bold'),
            padx=30,
            pady=10,
            cursor='hand2'
        )
        self.btn_generar_doc.pack(side='left', padx=10)
        
        # Pack canvas y scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
    def crear_tab_temas(self):
        """Crea el tab para gestión de temas"""
        frame = tk.Frame(self.notebook, bg=self.color_fondo)
        # Agregar pestaña con icono si está disponible
        if hasattr(self, 'icono_pestaña_temas') and self.icono_pestaña_temas:
            self.notebook.add(frame, text=" Gestión de Temas", image=self.icono_pestaña_temas, compound='left')
        else:
            self.notebook.add(frame, text="📝 Gestión de Temas")
        
        # Encabezado con canvas
        header_canvas = tk.Canvas(
            frame,
            bg=self.color_fondo,
            height=60,
            highlightthickness=0
        )
        header_canvas.pack(fill='x', padx=20, pady=(10, 5))
        
        header_canvas.create_line(0, 5, 1100, 5, fill=self.color_verde, width=2)
        header_canvas.create_text(
            50, 30,
            text="GESTIÓN DE TEMAS",
            font=('Arial', 13, 'bold'),
            fill=self.color_verde,
            anchor='w'
        )
        header_canvas.create_line(0, 55, 1100, 55, fill=self.color_verde_claro, width=1)
        
        # Canvas con scrollbar para todo el contenido
        canvas = tk.Canvas(frame, bg=self.color_fondo, highlightthickness=0)
        scrollbar_main = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=self.color_fondo)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar_main.set)
        
        # Título y botón
        frame_top = tk.Frame(scrollable_frame, bg=self.color_fondo)
        frame_top.pack(fill='x', padx=20, pady=10)
        
        self.btn_nuevo_tema = tk.Button(
            frame_top,
            text="➕ Nuevo Tema",
            bg=self.color_verde,
            fg='white',
            font=('Arial', 11, 'bold'),
            padx=20,
            pady=8,
            cursor='hand2'
        )
        self.btn_nuevo_tema.pack(side='right', padx=5)
        
        self.btn_cargar_masivo = tk.Button(
            frame_top,
            text="📥 Cargar desde Archivo",
            bg='#FF9800',
            fg='white',
            font=('Arial', 11, 'bold'),
            padx=20,
            pady=8,
            cursor='hand2'
        )
        self.btn_cargar_masivo.pack(side='right', padx=5)
        
        # Tabla de temas
        frame_tabla = tk.Frame(scrollable_frame, bg=self.color_fondo)
        frame_tabla.pack(fill='both', expand=True, padx=20, pady=10)
        
        columns = ('ID', 'Descripción', 'Categoría', 'Usos', 'Estado')
        self.tree_temas = ttk.Treeview(
            frame_tabla,
            columns=columns,
            show='headings',
            height=10,
            selectmode='extended'
        )
        
        self.tree_temas.heading('ID', text='ID')
        self.tree_temas.heading('Descripción', text='Descripción')
        self.tree_temas.heading('Categoría', text='Categoría')
        self.tree_temas.heading('Usos', text='Veces Usado')
        self.tree_temas.heading('Estado', text='Estado')
        
        self.tree_temas.column('ID', width=50)
        self.tree_temas.column('Descripción', width=500)
        self.tree_temas.column('Categoría', width=150)
        self.tree_temas.column('Usos', width=100)
        self.tree_temas.column('Estado', width=100)
        
        scrollbar = ttk.Scrollbar(frame_tabla, orient="vertical",
                                 command=self.tree_temas.yview)
        self.tree_temas.configure(yscrollcommand=scrollbar.set)
        
        self.tree_temas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Botones de acción
        frame_botones = tk.Frame(scrollable_frame, bg=self.color_fondo)
        frame_botones.pack(fill='x', padx=20, pady=10)
        
        self.btn_modificar_tema = tk.Button(
            frame_botones,
            text="✏️ Modificar",
            bg=self.color_verde_claro,
            fg='white',
            font=('Arial', 10, 'bold'),
            padx=15,
            pady=8,
            cursor='hand2'
        )
        self.btn_modificar_tema.pack(side='left', padx=5)
        
        self.btn_eliminar_tema = tk.Button(
            frame_botones,
            text="🗑️ Eliminar",
            bg='#D32F2F',
            fg='white',
            font=('Arial', 10, 'bold'),
            padx=15,
            pady=8,
            cursor='hand2'
        )
        self.btn_eliminar_tema.pack(side='left', padx=5)
        
        self.btn_historial_tema = tk.Button(
            frame_botones,
            text="📊 Ver Historial",
            bg='#1976D2',
            fg='white',
            font=('Arial', 10, 'bold'),
            padx=15,
            pady=8,
            cursor='hand2'
        )
        self.btn_historial_tema.pack(side='left', padx=5)
        
        self.btn_exportar_excel = tk.Button(
            frame_botones,
            text="📊 Exportar a Excel",
            bg='#2E7D32',
            fg='white',
            font=('Arial', 10, 'bold'),
            padx=15,
            pady=8,
            cursor='hand2'
        )
        self.btn_exportar_excel.pack(side='left', padx=5)
        
        self.btn_exportar_pdf = tk.Button(
            frame_botones,
            text="📄 Exportar a PDF",
            bg='#D32F2F',
            fg='white',
            font=('Arial', 10, 'bold'),
            padx=15,
            pady=8,
            cursor='hand2'
        )
        self.btn_exportar_pdf.pack(side='left', padx=5)
        
        self.btn_actualizar_temas = tk.Button(
            frame_botones,
            text="🔄 Actualizar",
            bg='#757575',
            fg='white',
            font=('Arial', 10),
            padx=15,
            pady=8,
            cursor='hand2'
        )
        self.btn_actualizar_temas.pack(side='right', padx=5)
        
        # Empaquetar canvas y scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar_main.pack(side="right", fill="y")
    
    def crear_tab_delegados(self):
        """Crea el tab para gestión de delegados"""
        frame = tk.Frame(self.notebook, bg=self.color_fondo)
        # Agregar pestaña con icono si está disponible
        if hasattr(self, 'icono_pestaña_delegados') and self.icono_pestaña_delegados:
            self.notebook.add(frame, text=" Gestión de Delegados", image=self.icono_pestaña_delegados, compound='left')
        else:
            self.notebook.add(frame, text="👥 Gestión de Delegados")
        
        # Encabezado con canvas
        header_canvas = tk.Canvas(
            frame,
            bg=self.color_fondo,
            height=60,
            highlightthickness=0
        )
        header_canvas.pack(fill='x', padx=20, pady=(10, 5))
        
        header_canvas.create_line(0, 5, 1100, 5, fill=self.color_verde, width=2)
        header_canvas.create_text(
            50, 30,
            text="GESTIÓN DE DELEGADOS",
            font=('Arial', 13, 'bold'),
            fill=self.color_verde,
            anchor='w'
        )
        header_canvas.create_line(0, 55, 1100, 55, fill=self.color_verde_claro, width=1)
        
        # Título y botón
        frame_top = tk.Frame(frame, bg=self.color_fondo)
        frame_top.pack(fill='x', padx=20, pady=10)
        
        self.btn_nuevo_delegado = tk.Button(
            frame_top,
            text="➕ Nuevo Delegado",
            bg=self.color_verde,
            fg='white',
            font=('Arial', 11, 'bold'),
            padx=20,
            pady=8,
            cursor='hand2'
        )
        self.btn_nuevo_delegado.pack(side='right')
        
        # Frame contenedor para tabla y botones
        frame_contenedor = tk.Frame(frame, bg=self.color_fondo)
        frame_contenedor.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Tabla de delegados
        frame_tabla = tk.Frame(frame_contenedor, bg=self.color_fondo)
        frame_tabla.pack(fill='both', expand=True, pady=(0, 10))
        
        columns = ('ID', 'Título', 'Nombre', 'Apellido', 'Distrito', 'Tipo')
        self.tree_delegados_lista = ttk.Treeview(
            frame_tabla,
            columns=columns,
            show='headings',
            height=15  # Reducir altura para dejar espacio a los botones
        )
        
        for col in columns:
            self.tree_delegados_lista.heading(col, text=col)
        
        self.tree_delegados_lista.column('ID', width=50)
        self.tree_delegados_lista.column('Título', width=80)
        self.tree_delegados_lista.column('Nombre', width=200)
        self.tree_delegados_lista.column('Apellido', width=200)
        self.tree_delegados_lista.column('Distrito', width=120)
        self.tree_delegados_lista.column('Tipo', width=100)
        
        scrollbar = ttk.Scrollbar(frame_tabla, orient="vertical",
                                 command=self.tree_delegados_lista.yview)
        self.tree_delegados_lista.configure(yscrollcommand=scrollbar.set)
        
        self.tree_delegados_lista.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Botones de acción - siempre visibles en la parte inferior
        frame_botones = tk.Frame(frame_contenedor, bg=self.color_fondo)
        frame_botones.pack(fill='x', pady=(10, 0))
        
        self.btn_modificar_delegado = tk.Button(
            frame_botones,
            text="✏️ Modificar",
            bg=self.color_verde_claro,
            fg='white',
            font=('Arial', 12, 'bold'),
            padx=25,
            pady=12,
            cursor='hand2',
            relief='raised',
            borderwidth=2
        )
        self.btn_modificar_delegado.pack(side='left', padx=10)
        
        self.btn_eliminar_delegado = tk.Button(
            frame_botones,
            text="🗑️ Eliminar",
            bg='#D32F2F',
            fg='white',
            font=('Arial', 12, 'bold'),
            padx=25,
            pady=12,
            cursor='hand2',
            relief='raised',
            borderwidth=2
        )
        self.btn_eliminar_delegado.pack(side='left', padx=10)
        
        self.btn_actualizar_delegados = tk.Button(
            frame_botones,
            text="🔄 Actualizar",
            bg='#757575',
            fg='white',
            font=('Arial', 11, 'bold'),
            padx=20,
            pady=12,
            cursor='hand2',
            relief='raised',
            borderwidth=2
        )
        self.btn_actualizar_delegados.pack(side='right', padx=10)
    
    def crear_tab_historial(self):
        """Crea el tab para historial de reuniones"""
        frame = tk.Frame(self.notebook, bg=self.color_fondo)
        # Agregar pestaña con icono si está disponible
        if hasattr(self, 'icono_pestaña_historial') and self.icono_pestaña_historial:
            self.notebook.add(frame, text=" Historial", image=self.icono_pestaña_historial, compound='left')
        else:
            self.notebook.add(frame, text="📚 Historial")
        
        # Encabezado con canvas
        header_canvas = tk.Canvas(
            frame,
            bg=self.color_fondo,
            height=60,
            highlightthickness=0
        )
        header_canvas.pack(fill='x', padx=20, pady=(10, 5))
        
        header_canvas.create_line(0, 5, 1100, 5, fill=self.color_verde, width=2)
        header_canvas.create_text(
            50, 30,
            text="HISTORIAL DE REUNIONES",
            font=('Arial', 13, 'bold'),
            fill=self.color_verde,
            anchor='w'
        )
        header_canvas.create_line(0, 55, 1100, 55, fill=self.color_verde_claro, width=1)
        
        # Frame de búsqueda
        frame_busqueda = tk.Frame(frame, bg=self.color_fondo)
        frame_busqueda.pack(fill='x', padx=20, pady=10)
        
        tk.Label(frame_busqueda, text="Buscar por Tema o Fecha:", bg=self.color_fondo, 
                font=('Arial', 9)).pack(side='left', padx=5)
        
        self.entry_buscar_historial = tk.Entry(frame_busqueda, width=30, font=('Arial', 9))
        self.entry_buscar_historial.pack(side='left', padx=5)
        
        self.btn_buscar_historial = tk.Button(
            frame_busqueda,
            text="🔍 Buscar",
            bg=self.color_verde,
            fg='white',
            font=('Arial', 9, 'bold'),
            padx=15,
            pady=5,
            cursor='hand2'
        )
        self.btn_buscar_historial.pack(side='left', padx=5)
        
        self.btn_limpiar_busqueda = tk.Button(
            frame_busqueda,
            text="🔄 Limpiar",
            bg='#757575',
            fg='white',
            font=('Arial', 9),
            padx=15,
            pady=5,
            cursor='hand2'
        )
        self.btn_limpiar_busqueda.pack(side='left', padx=2)
        
        # Frame de botones de acción
        frame_top = tk.Frame(frame, bg=self.color_fondo)
        frame_top.pack(fill='x', padx=20, pady=10)
        
        self.btn_exportar_historial_excel = tk.Button(
            frame_top,
            text="📊 Exportar a Excel",
            bg='#2E7D32',
            fg='white',
            font=('Arial', 10, 'bold'),
            padx=15,
            pady=8,
            cursor='hand2'
        )
        self.btn_exportar_historial_excel.pack(side='left', padx=5)
        
        self.btn_exportar_historial_pdf = tk.Button(
            frame_top,
            text="📄 Exportar a PDF",
            bg='#D32F2F',
            fg='white',
            font=('Arial', 10, 'bold'),
            padx=15,
            pady=8,
            cursor='hand2'
        )
        self.btn_exportar_historial_pdf.pack(side='left', padx=5)
        
        self.btn_actualizar_historial = tk.Button(
            frame_top,
            text="🔄 Actualizar",
            bg='#757575',
            fg='white',
            font=('Arial', 10),
            padx=15,
            pady=8,
            cursor='hand2'
        )
        self.btn_actualizar_historial.pack(side='right', padx=5)
        
        self.btn_borrar_historial = tk.Button(
            frame_top,
            text="🗑️ Borrar Seleccionadas",
            bg='#D32F2F',
            fg='white',
            font=('Arial', 10, 'bold'),
            padx=15,
            pady=8,
            cursor='hand2'
        )
        self.btn_borrar_historial.pack(side='right', padx=5)
        
        # Tabla de reuniones
        frame_tabla = tk.Frame(frame, bg=self.color_fondo)
        frame_tabla.pack(fill='both', expand=True, padx=20, pady=10)
        
        columns = ('ID', 'Fecha', 'Hora', 'Lugar', 'Tipo', 'Temas')
        self.tree_historial = ttk.Treeview(
            frame_tabla,
            columns=columns,
            show='headings',
            height=20,
            selectmode='extended'
        )
        
        for col in columns:
            self.tree_historial.heading(col, text=col)
        
        self.tree_historial.column('ID', width=50)
        self.tree_historial.column('Fecha', width=200)
        self.tree_historial.column('Hora', width=100)
        self.tree_historial.column('Lugar', width=300)
        self.tree_historial.column('Tipo', width=100)
        self.tree_historial.column('Temas', width=100)
        
        scrollbar = ttk.Scrollbar(frame_tabla, orient="vertical",
                                 command=self.tree_historial.yview)
        self.tree_historial.configure(yscrollcommand=scrollbar.set)
        
        self.tree_historial.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
    
    def _actualizar_visibilidad_plataforma(self, event=None):
        """Muestra u oculta el campo de plataforma según el tipo de reunión"""
        if self.combo_tipo.get() == "virtual":
            # Mostrar plataforma
            self.label_plataforma.grid(row=5, column=0, sticky='e', padx=5, pady=5)
            self.entry_plataforma.grid(row=5, column=1, padx=5, pady=5, sticky='w')
        else:
            # Ocultar plataforma
            self.label_plataforma.grid_remove()
            self.entry_plataforma.grid_remove()
    
    def _cargar_logo(self):
        """Abre diálogo para cargar imagen de logo"""
        ruta = filedialog.askopenfilename(
            title="Seleccionar imagen para logo",
            filetypes=[
                ("Imágenes", "*.png *.jpg *.jpeg *.gif *.bmp"),
                ("PNG", "*.png"),
                ("JPEG", "*.jpg *.jpeg"),
                ("Todos", "*.*")
            ]
        )
        
        if ruta:
            # Crear directorio assets si no existe
            os.makedirs("assets", exist_ok=True)
            
            try:
                # Copiar imagen a assets
                import shutil
                dest_path = os.path.join("assets", os.path.basename(ruta))
                shutil.copy(ruta, dest_path)
                self.imagen_path = dest_path
                
                # Mostrar en canvas
                self._mostrar_logo()
                messagebox.showinfo("Éxito", "Logo cargado correctamente")
            except Exception as e:
                messagebox.showerror("Error", f"Error al cargar logo: {str(e)}")
    
    def _cargar_logo_inicial(self):
        """Carga el logo si existe en la ruta por defecto"""
        if os.path.exists(self.imagen_path):
            self._mostrar_logo()
    
    def _mostrar_logo(self):
        """Muestra el logo en el canvas"""
        try:
            img = Image.open(self.imagen_path)
            # Redimensionar a 80x80
            img.thumbnail((80, 80), Image.Resampling.LANCZOS)
            self.imagen_logo = ImageTk.PhotoImage(img)
            
            # Limpiar y mostrar en canvas
            self.canvas_logo.delete("all")
            self.canvas_logo.create_image(40, 40, image=self.imagen_logo)
        except Exception as e:
            print(f"Error al mostrar logo: {e}")
    
    def _actualizar_encabezado_reunion(self):
        """Actualiza el encabezado de la reunión"""
        # El texto ya se actualiza automáticamente con la traza
        # Solo confirmamos con un mensaje
        messagebox.showinfo("Éxito", "Encabezado actualizado correctamente")
    
    def _on_titulo_canvas_configure(self, event):
        """Maneja el evento Configure del canvas del título"""
        try:
            ancho = event.width
            if ancho <= 0:
                return
            
            # Línea decorativa arriba
            self.titulo_canvas.delete("all")
            self.titulo_canvas.create_line(0, 5, ancho, 5, fill=self.color_verde, width=2)
            
            # Texto centrado
            self.titulo_canvas.create_text(
                ancho // 2, 38,
                text=self.texto_encabezado.get(),
                font=('Arial', 14, 'bold'),
                fill=self.color_verde,
                anchor='center'
            )
            
            # Línea decorativa abajo
            self.titulo_canvas.create_line(0, 65, ancho, 65, fill=self.color_verde_claro, width=1)
        except Exception as e:
            print(f"Error en _on_titulo_canvas_configure: {e}")
    
    def _on_texto_encabezado_cambio(self, *args):
        """Maneja cambios en el StringVar del texto del encabezado"""
        try:
            if hasattr(self, 'titulo_canvas') and self.titulo_canvas.winfo_exists():
                self.titulo_canvas.event_generate("<Configure>")
        except Exception as e:
            print(f"Error en _on_texto_encabezado_cambio: {e}")