"""
Diálogos para agregar/modificar temas y delegados
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext


class DialogoTema(tk.Toplevel):
    """Diálogo para agregar o modificar un tema"""
    
    def __init__(self, parent, titulo="Nuevo Tema", tema=None):
        super().__init__(parent)
        
        self.title(titulo)
        self.geometry("600x300")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        
        self.resultado = None
        
        # Frame principal
        frame = tk.Frame(self, bg='#F5F5F5', padx=20, pady=20)
        frame.pack(fill='both', expand=True)
        
        # Descripción
        tk.Label(
            frame,
            text="Descripción del Tema:",
            bg='#F5F5F5',
            font=('Arial', 10, 'bold')
        ).grid(row=0, column=0, sticky='w', pady=5)
        
        self.text_descripcion = scrolledtext.ScrolledText(
            frame,
            width=60,
            height=6,
            font=('Arial', 10),
            wrap=tk.WORD
        )
        self.text_descripcion.grid(row=1, column=0, columnspan=2, pady=5)
        
        if tema:
            self.text_descripcion.insert('1.0', tema.get('descripcion', ''))
        
        # Categoría
        tk.Label(
            frame,
            text="Categoría (opcional):",
            bg='#F5F5F5',
            font=('Arial', 10, 'bold')
        ).grid(row=2, column=0, sticky='w', pady=5)
        
        self.entry_categoria = tk.Entry(frame, width=40, font=('Arial', 10))
        self.entry_categoria.grid(row=3, column=0, sticky='w', pady=5)
        
        if tema:
            self.entry_categoria.insert(0, tema.get('categoria', ''))
        
        # Botones
        frame_botones = tk.Frame(frame, bg='#F5F5F5')
        frame_botones.grid(row=4, column=0, columnspan=2, pady=20)
        
        tk.Button(
            frame_botones,
            text="Guardar",
            command=self._guardar,
            bg='#2E7D32',
            fg='white',
            font=('Arial', 10, 'bold'),
            padx=20,
            pady=8,
            cursor='hand2'
        ).pack(side='left', padx=10)
        
        tk.Button(
            frame_botones,
            text="Cancelar",
            command=self.destroy,
            bg='#757575',
            fg='white',
            font=('Arial', 10),
            padx=20,
            pady=8,
            cursor='hand2'
        ).pack(side='left', padx=10)
    
    def _guardar(self):
        descripcion = self.text_descripcion.get('1.0', 'end-1c').strip()
        
        if not descripcion:
            messagebox.showwarning("Advertencia", "La descripción no puede estar vacía")
            return
        
        self.resultado = {
            'descripcion': descripcion,
            'categoria': self.entry_categoria.get().strip()
        }
        self.destroy()


class DialogoDelegado(tk.Toplevel):
    """Diálogo para agregar o modificar un delegado"""
    
    def __init__(self, parent, titulo="Nuevo Delegado", delegado=None):
        super().__init__(parent)
        
        self.title(titulo)
        self.geometry("500x350")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        
        self.resultado = None
        
        # Frame principal
        frame = tk.Frame(self, bg='#F5F5F5', padx=20, pady=20)
        frame.pack(fill='both', expand=True)
        
        # Título
        tk.Label(
            frame,
            text="Título:",
            bg='#F5F5F5',
            font=('Arial', 10, 'bold')
        ).grid(row=0, column=0, sticky='w', pady=5)
        
        self.combo_titulo = ttk.Combobox(
            frame,
            values=["Dr.", "Dra."],
            state='readonly',
            width=10
        )
        self.combo_titulo.grid(row=0, column=1, sticky='w', pady=5)
        self.combo_titulo.current(0)
        
        # Nombre
        tk.Label(
            frame,
            text="Nombre:",
            bg='#F5F5F5',
            font=('Arial', 10, 'bold')
        ).grid(row=1, column=0, sticky='w', pady=5)
        
        self.entry_nombre = tk.Entry(frame, width=40, font=('Arial', 10))
        self.entry_nombre.grid(row=1, column=1, pady=5, sticky='w')
        
        # Apellido
        tk.Label(
            frame,
            text="Apellido:",
            bg='#F5F5F5',
            font=('Arial', 10, 'bold')
        ).grid(row=2, column=0, sticky='w', pady=5)
        
        self.entry_apellido = tk.Entry(frame, width=40, font=('Arial', 10))
        self.entry_apellido.grid(row=2, column=1, pady=5, sticky='w')
        
        # Distrito
        tk.Label(
            frame,
            text="Distrito:",
            bg='#F5F5F5',
            font=('Arial', 10, 'bold')
        ).grid(row=3, column=0, sticky='w', pady=5)
        
        distritos = [f"Dist. {n}" for n in ['I', 'II', 'III', 'IV', 'V', 
                                             'VI', 'VII', 'VIII', 'IX', 'X']]
        self.combo_distrito = ttk.Combobox(
            frame,
            values=distritos,
            state='readonly',
            width=15
        )
        self.combo_distrito.grid(row=3, column=1, sticky='w', pady=5)
        self.combo_distrito.current(0)
        
        # Tipo
        tk.Label(
            frame,
            text="Tipo:",
            bg='#F5F5F5',
            font=('Arial', 10, 'bold')
        ).grid(row=4, column=0, sticky='w', pady=5)
        
        self.var_titular = tk.BooleanVar(value=True)
        tk.Checkbutton(
            frame,
            text="Titular",
            variable=self.var_titular,
            bg='#F5F5F5',
            font=('Arial', 10)
        ).grid(row=4, column=1, sticky='w', pady=5)
        
        # Cargar datos si es modificación
        if delegado:
            self.combo_titulo.set(delegado.get('titulo', 'Dr.'))
            self.entry_nombre.insert(0, delegado.get('nombre', ''))
            self.entry_apellido.insert(0, delegado.get('apellido', ''))
            self.combo_distrito.set(delegado.get('distrito', 'Dist. I'))
            self.var_titular.set(delegado.get('titular', 1) == 1)
        
        # Botones
        frame_botones = tk.Frame(frame, bg='#F5F5F5')
        frame_botones.grid(row=5, column=0, columnspan=2, pady=20)
        
        tk.Button(
            frame_botones,
            text="Guardar",
            command=self._guardar,
            bg='#2E7D32',
            fg='white',
            font=('Arial', 10, 'bold'),
            padx=20,
            pady=8,
            cursor='hand2'
        ).pack(side='left', padx=10)
        
        tk.Button(
            frame_botones,
            text="Cancelar",
            command=self.destroy,
            bg='#757575',
            fg='white',
            font=('Arial', 10),
            padx=20,
            pady=8,
            cursor='hand2'
        ).pack(side='left', padx=10)
    
    def _guardar(self):
        nombre = self.entry_nombre.get().strip()
        apellido = self.entry_apellido.get().strip()
        
        if not nombre or not apellido:
            messagebox.showwarning("Advertencia", "El nombre y apellido son obligatorios")
            return
        
        self.resultado = {
            'titulo': self.combo_titulo.get(),
            'nombre': nombre,
            'apellido': apellido,
            'distrito': self.combo_distrito.get(),
            'titular': self.var_titular.get()
        }
        self.destroy()


class DialogoSeleccionTema(tk.Toplevel):
    """Diálogo para seleccionar un tema de la lista"""
    
    def __init__(self, parent, temas):
        super().__init__(parent)
        
        self.title("Seleccionar Tema")
        self.geometry("700x500")
        self.transient(parent)
        self.grab_set()
        
        self.resultado = None
        self.temas = temas
        
        # Frame principal
        frame = tk.Frame(self, bg='#F5F5F5', padx=20, pady=20)
        frame.pack(fill='both', expand=True)
        
        tk.Label(
            frame,
            text="Seleccione un tema para agregar al orden del día:",
            bg='#F5F5F5',
            font=('Arial', 12, 'bold')
        ).pack(pady=10)
        
        # Tabla de temas
        columns = ('ID', 'Descripción', 'Categoría')
        self.tree = ttk.Treeview(
            frame,
            columns=columns,
            show='headings',
            height=15
        )
        
        self.tree.heading('ID', text='ID')
        self.tree.heading('Descripción', text='Descripción')
        self.tree.heading('Categoría', text='Categoría')
        
        self.tree.column('ID', width=50)
        self.tree.column('Descripción', width=500)
        self.tree.column('Categoría', width=100)
        
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Cargar temas
        for tema in temas:
            self.tree.insert('', 'end', values=(
                tema['id'],
                tema['descripcion'],
                tema['categoria'] or '-'
            ), tags=(tema['id'],))
        
        # Botones
        frame_botones = tk.Frame(self, bg='#F5F5F5')
        frame_botones.pack(fill='x', pady=10)
        
        tk.Button(
            frame_botones,
            text="Seleccionar",
            command=self._seleccionar,
            bg='#2E7D32',
            fg='white',
            font=('Arial', 10, 'bold'),
            padx=20,
            pady=8,
            cursor='hand2'
        ).pack(side='left', padx=10)
        
        tk.Button(
            frame_botones,
            text="Cancelar",
            command=self.destroy,
            bg='#757575',
            fg='white',
            font=('Arial', 10),
            padx=20,
            pady=8,
            cursor='hand2'
        ).pack(side='left', padx=10)
        
        # Doble click para seleccionar
        self.tree.bind('<Double-1>', lambda e: self._seleccionar())
    
    def _seleccionar(self):
        seleccion = self.tree.selection()
        
        if not seleccion:
            messagebox.showwarning("Advertencia", "Seleccione un tema")
            return
        
        item = seleccion[0]
        self.resultado = self.tree.item(item, 'tags')[0]
        self.destroy()


class VentanaVistaPrevia(tk.Toplevel):
    """Ventana de vista previa del documento"""
    
    def __init__(self, parent, contenido):
        super().__init__(parent)
        
        self.title("Vista Previa - Orden del Día")
        self.geometry("900x700")
        self.transient(parent)
        
        # Frame con scroll
        frame = tk.Frame(self, bg='white')
        frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Texto con scroll
        self.text = scrolledtext.ScrolledText(
            frame,
            width=100,
            height=40,
            font=('Courier', 10),
            wrap=tk.WORD,
            bg='white'
        )
        self.text.pack(fill='both', expand=True)
        self.text.insert('1.0', contenido)
        self.text.config(state='disabled')
        
        # Botón cerrar
        tk.Button(
            self,
            text="Cerrar",
            command=self.destroy,
            bg='#757575',
            fg='white',
            font=('Arial', 10),
            padx=20,
            pady=8,
            cursor='hand2'
        ).pack(pady=10)


class VentanaHistorialTema(tk.Toplevel):
    """Ventana para mostrar el historial de un tema"""
    
    def __init__(self, parent, tema, historial, stats):
        super().__init__(parent)
        
        self.title(f"Historial: {tema['descripcion'][:50]}...")
        self.geometry("800x600")
        self.transient(parent)
        
        # Frame superior con estadísticas
        frame_stats = tk.Frame(self, bg='#E8F5E9', padx=20, pady=15)
        frame_stats.pack(fill='x')
        
        tk.Label(
            frame_stats,
            text=f"Tema: {tema['descripcion']}",
            bg='#E8F5E9',
            font=('Arial', 12, 'bold'),
            fg='#2E7D32'
        ).pack(anchor='w', pady=5)
        
        tk.Label(
            frame_stats,
            text=f"Usado {stats['cantidad_usos']} veces",
            bg='#E8F5E9',
            font=('Arial', 10)
        ).pack(anchor='w')
        
        if stats['primera_fecha']:
            tk.Label(
                frame_stats,
                text=f"Primera vez: {stats['primera_fecha']}",
                bg='#E8F5E9',
                font=('Arial', 10)
            ).pack(anchor='w')
        
        if stats['ultima_fecha']:
            tk.Label(
                frame_stats,
                text=f"Última vez: {stats['ultima_fecha']}",
                bg='#E8F5E9',
                font=('Arial', 10)
            ).pack(anchor='w')
        
        # Frame con tabla de historial
        frame_middle = tk.Frame(self, bg='white')
        frame_middle.pack(fill='both', expand=True, padx=10, pady=10)
        
        columns = ('Fecha', 'Lugar', 'Sede', 'Tipo', 'Nº Orden')
        tree = ttk.Treeview(
            frame_middle,
            columns=columns,
            show='headings',
            height=15
        )
        
        for col in columns:
            tree.heading(col, text=col)
        
        tree.column('Fecha', width=200)
        tree.column('Lugar', width=250)
        tree.column('Sede', width=150)
        tree.column('Tipo', width=100)
        tree.column('Nº Orden', width=80)
        
        scrollbar = ttk.Scrollbar(frame_middle, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Cargar historial
        for h in historial:
            tree.insert('', 'end', values=(
                h['fecha'],
                h['lugar'],
                h['sede'] or '-',
                h['tipo'],
                h['numero_orden']
            ))
        
        # Botón cerrar
        tk.Button(
            self,
            text="Cerrar",
            command=self.destroy,
            bg='#757575',
            fg='white',
            font=('Arial', 10),
            padx=20,
            pady=8,
            cursor='hand2'
        ).pack(pady=10)