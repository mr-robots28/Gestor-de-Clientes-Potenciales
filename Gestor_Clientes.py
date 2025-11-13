# Creado por Wayner Castillo
# github: https://github.com/cybersecrd

import pandas as pd
import os
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import seaborn as sns
import numpy as np

# Configuración para gráficos
plt.style.use('default')
sns.set_palette("husl")

class GestorClientesApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Gestión de Clientes Potenciales")
        self.root.geometry("1300x800")
        self.root.configure(bg='#f0f0f0')
        
        self.archivo_excel = "clientes_potenciales.xlsx"
        self.columnas = [
            'ID', 'Nombre_Empresa', 'Sector', 'Localidad', 'Telefono', 
            'Correo_Electronico', 'Estado_Contacto', 'Fecha_Contacto', 
            'Observaciones', 'Sitio_Web_Actual', 'Interes', 'Fecha_Proximo_Contacto',
            'Es_Cliente', 'Solicito_Propuesta', 'Se_Le_Envio_Propuesta', 'Fecha_Envio_Propuesta'
        ]
        
        self.inicializar_archivo()
        self.crear_interfaz()
        self.actualizar_lista_clientes()
    
    def inicializar_archivo(self):
        """Crea el archivo Excel si no existe"""
        if not os.path.exists(self.archivo_excel):
            df = pd.DataFrame(columns=self.columnas)
            df.to_excel(self.archivo_excel, index=False)
    
    def leer_clientes(self):
        """Lee todos los clientes del archivo Excel"""
        try:
            df = pd.read_excel(self.archivo_excel)
            # Asegurar que las columnas existan
            for col in self.columnas:
                if col not in df.columns:
                    df[col] = None
            return df
        except Exception as e:
            print(f"Error al leer archivo: {e}")
            return pd.DataFrame(columns=self.columnas)
    
    def guardar_clientes(self, df):
        """Guarda el DataFrame en el archivo Excel"""
        try:
            df.to_excel(self.archivo_excel, index=False)
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar: {e}")
            return False
    
    def crear_interfaz(self):
        """Crea la interfaz gráfica principal"""
        # Frame principal con paned window para mejor control
        main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Frame izquierdo para controles
        left_frame = ttk.Frame(main_paned, width=300)
        main_paned.add(left_frame, weight=0)
        
        # Frame derecho para la tabla
        right_frame = ttk.Frame(main_paned)
        main_paned.add(right_frame, weight=1)
        
        # Configurar el frame izquierdo
        self.crear_panel_controles(left_frame)
        
        # Configurar el frame derecho
        self.crear_panel_tabla(right_frame)
        
        # Hacer que el frame izquierdo mantenga su tamaño
        left_frame.pack_propagate(False)
    
    def crear_panel_controles(self, parent):
        """Crea el panel de controles a la izquierda"""
        # Título
        titulo = ttk.Label(parent, text="Gestión de Clientes", 
                          font=('Arial', 14, 'bold'))
        titulo.pack(pady=15, padx=10)
        
        # Frame para estadísticas rápidas
        frame_stats = ttk.LabelFrame(parent, text="Estadísticas Rápidas", padding=10)
        frame_stats.pack(fill=tk.X, padx=10, pady=10)
        
        self.stats_labels = {}
        stats_info = [
            ('Total:', 'total_clientes'),
            ('Sin Web:', 'sin_web'),
            ('Interés Alto:', 'interes_alto'),
            ('Clientes:', 'es_cliente'),
            ('Propuestas Env:', 'propuestas_env')
        ]
        
        for text, key in stats_info:
            frame_stat = ttk.Frame(frame_stats)
            frame_stat.pack(fill=tk.X, pady=2)
            ttk.Label(frame_stat, text=text, width=12, anchor='w').pack(side=tk.LEFT)
            self.stats_labels[key] = ttk.Label(frame_stat, text="0", font=('Arial', 10, 'bold'))
            self.stats_labels[key].pack(side=tk.LEFT)
        
        # Frame para acciones principales
        frame_acciones = ttk.LabelFrame(parent, text="Acciones", padding=10)
        frame_acciones.pack(fill=tk.X, padx=10, pady=10)
        
        botones_principales = [
            ("➕ Agregar Cliente", self.mostrar_formulario_agregar),
            ("✏️ Modificar Seleccionado", self.mostrar_formulario_modificar),
            ("🗑️ Eliminar Seleccionado", self.eliminar_cliente),
            ("📊 Ver Gráficos", self.mostrar_graficos),
            ("🔄 Actualizar Lista", self.actualizar_lista_clientes)
        ]
        
        for texto, comando in botones_principales:
            btn = ttk.Button(frame_acciones, text=texto, command=comando)
            btn.pack(fill=tk.X, pady=3)
        
        # Frame para búsqueda rápida
        frame_busqueda = ttk.LabelFrame(parent, text="Búsqueda Rápida", padding=10)
        frame_busqueda.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(frame_busqueda, text="Buscar:").pack(anchor=tk.W)
        self.busqueda_var = tk.StringVar()
        self.busqueda_entry = ttk.Entry(frame_busqueda, textvariable=self.busqueda_var)
        self.busqueda_entry.pack(fill=tk.X, pady=5)
        self.busqueda_entry.bind('<KeyRelease>', self.buscar_cliente)
        
        ttk.Label(frame_busqueda, text="Por:").pack(anchor=tk.W)
        self.criterio_busqueda = ttk.Combobox(frame_busqueda, 
                                            values=['Nombre_Empresa', 'Sector', 'Localidad', 'Estado_Contacto', 'Interes'],
                                            state='readonly')
        self.criterio_busqueda.set('Nombre_Empresa')
        self.criterio_busqueda.pack(fill=tk.X, pady=5)
        
        # Botón búsqueda avanzada
        ttk.Button(frame_busqueda, text="🔍 Búsqueda Avanzada", 
                  command=self.mostrar_busqueda).pack(fill=tk.X, pady=5)
        
        # Frame para filtros rápidos
        frame_filtros = ttk.LabelFrame(parent, text="Filtros Rápidos", padding=10)
        frame_filtros.pack(fill=tk.X, padx=10, pady=10)
        
        filtros_rapidos = [
            ("📞 Por contactar", "Por contactar"),
            ("✅ Contactados", "Contactado"),
            ("🔄 En seguimiento", "En seguimiento"),
            ("❌ No interesados", "No interesado"),
            ("💰 Clientes", "Cliente"),
            ("🌟 Interés Alto", "Alto"),
            ("📨 Con propuesta", "SI")
        ]
        
        for texto, filtro in filtros_rapidos:
            btn = ttk.Button(frame_filtros, text=texto, 
                           command=lambda f=filtro: self.aplicar_filtro_rapido(f))
            btn.pack(fill=tk.X, pady=2)
        
        ttk.Button(frame_filtros, text="🧹 Limpiar Filtros", 
                  command=self.limpiar_filtros).pack(fill=tk.X, pady=5)
    
    def crear_panel_tabla(self, parent):
        """Crea el panel de la tabla a la derecha"""
        # Frame para la tabla
        frame_tabla = ttk.Frame(parent)
        frame_tabla.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Label de información
        self.info_label = ttk.Label(frame_tabla, text="Total de clientes: 0", 
                                   font=('Arial', 10))
        self.info_label.pack(anchor=tk.W, pady=5)
        
        # Scrollbars
        v_scroll = ttk.Scrollbar(frame_tabla, orient=tk.VERTICAL)
        h_scroll = ttk.Scrollbar(frame_tabla, orient=tk.HORIZONTAL)
        
        # Treeview
        self.tree = ttk.Treeview(frame_tabla, columns=self.columnas, show='headings',
                                yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set,
                                height=20)
        
        # Configurar columnas
        anchos_columnas = {
            'ID': 50,
            'Nombre_Empresa': 150,
            'Sector': 120,
            'Localidad': 100,
            'Telefono': 100,
            'Correo_Electronico': 150,
            'Estado_Contacto': 120,
            'Fecha_Contacto': 100,
            'Observaciones': 200,
            'Sitio_Web_Actual': 80,
            'Interes': 100,
            'Fecha_Proximo_Contacto': 120,
            'Es_Cliente': 80,
            'Solicito_Propuesta': 100,
            'Se_Le_Envio_Propuesta': 100,
            'Fecha_Envio_Propuesta': 120
        }
        
        for col in self.columnas:
            ancho = anchos_columnas.get(col, 100)
            self.tree.heading(col, text=col.replace('_', ' ').title())
            self.tree.column(col, width=ancho, minwidth=50)
        
        v_scroll.config(command=self.tree.yview)
        h_scroll.config(command=self.tree.xview)
        
        # Empaquetar treeview y scrollbars
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Bind events
        self.tree.bind('<Double-1>', self.editar_doble_click)
        self.tree.bind('<<TreeviewSelect>>', self.actualizar_info_seleccion)
    
    def actualizar_estadisticas_rapidas(self):
        """Actualiza las estadísticas rápidas en el panel izquierdo"""
        df = self.leer_clientes()
        total = len(df)
        
        stats = {
            'total_clientes': total,
            'sin_web': len(df[df['Sitio_Web_Actual'] == 'No tiene']),
            'interes_alto': len(df[df['Interes'] == 'Alto']),
            'es_cliente': len(df[df['Es_Cliente'] == 'SI']),
            'propuestas_env': len(df[df['Se_Le_Envio_Propuesta'] == 'SI'])
        }
        
        for key, value in stats.items():
            self.stats_labels[key].config(text=str(value))
    
    def actualizar_info_seleccion(self, event=None):
        """Actualiza la información de la selección actual"""
        seleccion = self.tree.selection()
        if seleccion:
            texto = f"Clientes seleccionados: {len(seleccion)}"
        else:
            texto = f"Total de clientes: {len(self.tree.get_children())}"
        self.info_label.config(text=texto)
    
    def aplicar_filtro_rapido(self, filtro):
        """Aplica filtros rápidos desde el panel izquierdo"""
        df = self.leer_clientes()
        
        if filtro in ['Por contactar', 'Contactado', 'En seguimiento', 'No interesado', 'Cliente']:
            resultados = df[df['Estado_Contacto'] == filtro]
        elif filtro == 'Alto':
            resultados = df[df['Interes'] == 'Alto']
        elif filtro == 'SI':
            resultados = df[df['Se_Le_Envio_Propuesta'] == 'SI']
        else:
            resultados = df
        
        self.actualizar_lista_clientes(resultados)
        messagebox.showinfo("Filtro", f"Mostrando {len(resultados)} clientes con filtro: {filtro}")
    
    def limpiar_filtros(self):
        """Limpia todos los filtros aplicados"""
        self.busqueda_var.set("")
        self.actualizar_lista_clientes()
        messagebox.showinfo("Filtros", "Todos los filtros han sido limpiados")
    
    def actualizar_lista_clientes(self, df=None):
        """Actualiza la lista de clientes en el Treeview"""
        if df is None:
            df = self.leer_clientes()
        
        # Limpiar treeview
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Insertar datos
        for _, row in df.iterrows():
            valores = [row[col] if pd.notna(row[col]) else '' for col in self.columnas]
            item = self.tree.insert('', tk.END, values=valores)
            
            # Aplicar colores según estado
            estado = row['Estado_Contacto'] if pd.notna(row['Estado_Contacto']) else 'Por contactar'
            self.aplicar_color_fila(item, estado, row)
        
        self.actualizar_estadisticas_rapidas()
        self.actualizar_info_seleccion()
    
    def aplicar_color_fila(self, item, estado, row):
        """Aplica colores a las filas según diferentes criterios"""
        # Colores base por estado
        colores_estado = {
            'Por contactar': '#fff3cd',
            'Contactado': '#d1ecf1', 
            'En seguimiento': '#d4edda',
            'No interesado': '#f8d7da',
            'Cliente': '#c3e6cb'
        }
        
        color = colores_estado.get(estado, 'white')
        
        # Resaltar clientes actuales
        if row.get('Es_Cliente') == 'SI':
            color = '#d4edda'  # Verde más fuerte
        
        # Resaltar interés alto
        if row.get('Interes') == 'Alto':
            color = '#fff3cd'  # Amarillo
        
        self.tree.item(item, tags=(estado,))
        self.tree.tag_configure(estado, background=color)
    
    def buscar_cliente(self, event=None):
        """Busca clientes según el criterio"""
        criterio = self.criterio_busqueda.get()
        valor = self.busqueda_var.get().lower()
        
        if not valor:
            self.actualizar_lista_clientes()
            return
        
        df = self.leer_clientes()
        if criterio in df.columns:
            resultados = df[df[criterio].astype(str).str.lower().str.contains(valor, na=False)]
            self.actualizar_lista_clientes(resultados)
    
    def mostrar_formulario_agregar(self):
        """Muestra formulario para agregar cliente"""
        self.formulario_cliente("Agregar Cliente")
    
    def mostrar_formulario_modificar(self):
        """Muestra formulario para modificar cliente seleccionado"""
        seleccion = self.tree.selection()
        if not seleccion:
            messagebox.showwarning("Advertencia", "Por favor selecciona un cliente para modificar.")
            return
        
        item = seleccion[0]
        valores = self.tree.item(item)['values']
        if not valores:
            return
            
        cliente_id = valores[0]
        
        df = self.leer_clientes()
        cliente_data = df[df['ID'] == cliente_id]
        
        if cliente_data.empty:
            messagebox.showerror("Error", "No se encontró el cliente seleccionado.")
            return
        
        cliente = cliente_data.iloc[0]
        self.formulario_cliente("Modificar Cliente", cliente)
    
    def editar_doble_click(self, event):
        """Edita cliente al hacer doble click"""
        self.mostrar_formulario_modificar()
    
    def formulario_cliente(self, titulo, cliente=None):
        """Crea formulario para agregar/modificar cliente"""
        formulario = tk.Toplevel(self.root)
        formulario.title(titulo)
        formulario.geometry("600x800")
        formulario.transient(self.root)
        formulario.grab_set()
        
        # Frame principal con scroll
        main_frame = ttk.Frame(formulario)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Canvas y scrollbar
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Frame del formulario
        form_frame = ttk.Frame(scrollable_frame, padding=20)
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(form_frame, text=titulo, font=('Arial', 16, 'bold')).pack(pady=10)
        
        # Campos del formulario organizados en dos columnas
        campos_frame = ttk.Frame(form_frame)
        campos_frame.pack(fill=tk.BOTH, expand=True)
        
        # Columna izquierda
        left_frame = ttk.Frame(campos_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        
        # Columna derecha  
        right_frame = ttk.Frame(campos_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5)
        
        campos = [
            # (campo, label, tipo, opciones, columna)
            ('Nombre_Empresa', 'Nombre de la empresa*', 'text', None, 'left'),
            ('Sector', 'Sector', 'text', None, 'left'),
            ('Localidad', 'Localidad', 'text', None, 'left'),
            ('Telefono', 'Teléfono', 'text', None, 'left'),
            ('Correo_Electronico', 'Correo electrónico', 'text', None, 'left'),
            ('Estado_Contacto', 'Estado de contacto', 'combo', 
             ['Por contactar', 'Contactado', 'En seguimiento', 'No interesado', 'Cliente'], 'left'),
            ('Interes', 'Nivel de interés', 'combo', ['No evaluado', 'Bajo', 'Medio', 'Alto'], 'left'),
            ('Sitio_Web_Actual', 'Sitio web actual', 'combo', ['No tiene', 'Tiene'], 'right'),
            ('Es_Cliente', '¿Es cliente actual?', 'combo', ['NO', 'SI'], 'right'),
            ('Solicito_Propuesta', '¿Solicitó propuesta?', 'combo', ['NO', 'SI'], 'right'),
            ('Se_Le_Envio_Propuesta', '¿Se le envió propuesta?', 'combo', ['NO', 'SI'], 'right'),
            ('Fecha_Proximo_Contacto', 'Fecha próximo contacto', 'text', None, 'right'),
            ('Fecha_Envio_Propuesta', 'Fecha envío propuesta', 'text', None, 'right'),
            ('Observaciones', 'Observaciones', 'text', None, 'full')  # Campo completo
        ]
        
        self.entries = {}
        
        for campo, label, tipo, opciones, columna in campos:
            frame_campo = ttk.Frame(left_frame if columna == 'left' else right_frame if columna == 'right' else campos_frame)
            frame_campo.pack(fill=tk.X, pady=3)
            
            ttk.Label(frame_campo, text=label, width=25, anchor='w').pack(side=tk.LEFT)
            
            if tipo == 'text':
                entry = ttk.Entry(frame_campo, width=30)
                if columna == 'full':
                    entry.pack(fill=tk.X, padx=5)
                else:
                    entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
            elif tipo == 'combo':
                entry = ttk.Combobox(frame_campo, values=opciones, width=28, state='readonly')
                if columna == 'full':
                    entry.pack(fill=tk.X, padx=5)
                else:
                    entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
                entry.set(opciones[0] if opciones else '')
            
            self.entries[campo] = entry
            
            # Llenar con datos existentes si los hay
            if cliente is not None and campo in cliente:
                valor = cliente[campo]
                if pd.notna(valor):
                    if tipo == 'combo':
                        entry.set(str(valor))
                    else:
                        entry.delete(0, tk.END)
                        entry.insert(0, str(valor))
        
        # Botones en la parte inferior
        botones_frame = ttk.Frame(form_frame)
        botones_frame.pack(fill=tk.X, pady=20)
        
        if cliente is None:
            ttk.Button(botones_frame, text="Guardar Cliente", 
                      command=lambda: self.guardar_nuevo_cliente(formulario)).pack(side=tk.LEFT, padx=10)
        else:
            ttk.Button(botones_frame, text="Actualizar Cliente", 
                      command=lambda: self.actualizar_cliente_existente(cliente['ID'], formulario)).pack(side=tk.LEFT, padx=10)
        
        ttk.Button(botones_frame, text="Cancelar", 
                  command=formulario.destroy).pack(side=tk.LEFT, padx=10)
        
        # Empaquetar canvas y scrollbar al final
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def guardar_nuevo_cliente(self, formulario):
        """Guarda un nuevo cliente"""
        datos = self.obtener_datos_formulario()
        if datos is None:
            return
        
        df = self.leer_clientes()
        nuevo_id = df['ID'].max() + 1 if not df.empty else 1
        datos['ID'] = nuevo_id
        datos['Fecha_Contacto'] = datetime.now().strftime('%Y-%m-%d')
        
        # Si se envió propuesta pero no hay fecha, usar fecha actual
        if datos.get('Se_Le_Envio_Propuesta') == 'SI' and not datos.get('Fecha_Envio_Propuesta'):
            datos['Fecha_Envio_Propuesta'] = datetime.now().strftime('%Y-%m-%d')
        
        df = pd.concat([df, pd.DataFrame([datos])], ignore_index=True)
        
        if self.guardar_clientes(df):
            messagebox.showinfo("Éxito", "Cliente agregado correctamente.")
            formulario.destroy()
            self.actualizar_lista_clientes()
    
    def actualizar_cliente_existente(self, cliente_id, formulario):
        """Actualiza un cliente existente"""
        datos = self.obtener_datos_formulario()
        if datos is None:
            return
        
        df = self.leer_clientes()
        
        # Si se envió propuesta pero no hay fecha, usar fecha actual
        if datos.get('Se_Le_Envio_Propuesta') == 'SI' and not datos.get('Fecha_Envio_Propuesta'):
            datos['Fecha_Envio_Propuesta'] = datetime.now().strftime('%Y-%m-%d')
        
        for campo, valor in datos.items():
            if campo in df.columns and campo != 'ID':
                df.loc[df['ID'] == cliente_id, campo] = valor
        
        if self.guardar_clientes(df):
            messagebox.showinfo("Éxito", "Cliente actualizado correctamente.")
            formulario.destroy()
            self.actualizar_lista_clientes()
    
    def obtener_datos_formulario(self):
        """Obtiene y valida los datos del formulario"""
        datos = {}
        
        # Validar campo obligatorio
        if not self.entries['Nombre_Empresa'].get().strip():
            messagebox.showerror("Error", "El nombre de la empresa es obligatorio.")
            return None
        
        for campo, entry in self.entries.items():
            if isinstance(entry, ttk.Combobox):
                datos[campo] = entry.get()
            else:
                datos[campo] = entry.get().strip()
        
        return datos
    
    def eliminar_cliente(self):
        """Elimina el cliente seleccionado"""
        seleccion = self.tree.selection()
        if not seleccion:
            messagebox.showwarning("Advertencia", "Por favor selecciona un cliente para eliminar.")
            return
        
        item = seleccion[0]
        valores = self.tree.item(item)['values']
        if not valores:
            return
            
        cliente_id = valores[0]
        nombre_empresa = valores[1]
        
        respuesta = messagebox.askyesno(
            "Confirmar eliminación", 
            f"¿Estás seguro de eliminar a '{nombre_empresa}' (ID: {cliente_id})?"
        )
        
        if respuesta:
            df = self.leer_clientes()
            df = df[df['ID'] != cliente_id]
            
            if self.guardar_clientes(df):
                messagebox.showinfo("Éxito", "Cliente eliminado correctamente.")
                self.actualizar_lista_clientes()
    
    def mostrar_busqueda(self):
        """Muestra ventana de búsqueda avanzada"""
        busqueda_win = tk.Toplevel(self.root)
        busqueda_win.title("Búsqueda Avanzada")
        busqueda_win.geometry("400x500")
        busqueda_win.transient(self.root)
        
        frame = ttk.Frame(busqueda_win, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="Búsqueda Avanzada", font=('Arial', 14, 'bold')).pack(pady=10)
        
        # Campos de búsqueda
        campos_busqueda = [
            ('Nombre_Empresa', 'Nombre de empresa', 'text'),
            ('Sector', 'Sector', 'text'),
            ('Localidad', 'Localidad', 'text'),
            ('Estado_Contacto', 'Estado de contacto', 'combo'),
            ('Interes', 'Nivel de interés', 'combo'),
            ('Es_Cliente', 'Es cliente', 'combo'),
            ('Solicito_Propuesta', 'Solicitó propuesta', 'combo'),
            ('Se_Le_Envio_Propuesta', 'Se le envió propuesta', 'combo')
        ]
        
        self.entries_busqueda = {}
        
        for campo, label, tipo in campos_busqueda:
            ttk.Label(frame, text=label).pack(anchor=tk.W, pady=2)
            
            if tipo == 'combo':
                if campo in ['Es_Cliente', 'Solicito_Propuesta', 'Se_Le_Envio_Propuesta']:
                    entry = ttk.Combobox(frame, values=['', 'SI', 'NO'], width=37, state='readonly')
                elif campo == 'Estado_Contacto':
                    entry = ttk.Combobox(frame, values=['', 'Por contactar', 'Contactado', 'En seguimiento', 'No interesado', 'Cliente'], 
                                       width=37, state='readonly')
                elif campo == 'Interes':
                    entry = ttk.Combobox(frame, values=['', 'No evaluado', 'Bajo', 'Medio', 'Alto'], 
                                       width=37, state='readonly')
                entry.set('')
            else:
                entry = ttk.Entry(frame, width=40)
            
            entry.pack(pady=2, fill=tk.X)
            self.entries_busqueda[campo] = entry
        
        # Botones
        frame_botones = ttk.Frame(frame)
        frame_botones.pack(pady=20)
        
        ttk.Button(frame_botones, text="Buscar", 
                  command=lambda: self.ejecutar_busqueda_avanzada(busqueda_win)).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_botones, text="Limpiar", 
                  command=self.limpiar_busqueda).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_botones, text="Cerrar", 
                  command=busqueda_win.destroy).pack(side=tk.LEFT, padx=5)
    
    def ejecutar_busqueda_avanzada(self, ventana):
        """Ejecuta búsqueda avanzada"""
        df = self.leer_clientes()
        
        for campo, entry in self.entries_busqueda.items():
            valor = entry.get().strip()
            if valor:
                if campo in ['Es_Cliente', 'Solicito_Propuesta', 'Se_Le_Envio_Propuesta', 'Estado_Contacto', 'Interes']:
                    df = df[df[campo] == valor]
                else:
                    df = df[df[campo].astype(str).str.lower().str.contains(valor.lower(), na=False)]
        
        self.actualizar_lista_clientes(df)
        ventana.destroy()
        
        if df.empty:
            messagebox.showinfo("Búsqueda", "No se encontraron resultados.")
        else:
            messagebox.showinfo("Búsqueda", f"Se encontraron {len(df)} resultados.")
    
    def limpiar_busqueda(self):
        """Limpia los campos de búsqueda"""
        for entry in self.entries_busqueda.values():
            if isinstance(entry, ttk.Combobox):
                entry.set('')
            else:
                entry.delete(0, tk.END)
    
    def mostrar_graficos(self):
        """Muestra ventana con gráficos"""
        graficos_win = tk.Toplevel(self.root)
        graficos_win.title("Análisis Visual de Clientes")
        graficos_win.geometry("1000x800")
        graficos_win.transient(self.root)
        
        notebook = ttk.Notebook(graficos_win)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Pestaña de gráficos principales
        frame_principales = ttk.Frame(notebook)
        notebook.add(frame_principales, text="Gráficos Principales")
        
        self.crear_grafico_estados(frame_principales)
        
        # Pestaña de nuevos campos
        frame_nuevos = ttk.Frame(notebook)
        notebook.add(frame_nuevos, text="Seguimiento Propuestas")
        
        self.crear_grafico_propuestas(frame_nuevos)
        
        # Pestaña de localidades
        frame_localidades = ttk.Frame(notebook)
        notebook.add(frame_localidades, text="Localidades")
        
        self.crear_grafico_localidades(frame_localidades)
        
        # Pestaña de estadísticas
        frame_stats = ttk.Frame(notebook)
        notebook.add(frame_stats, text="Estadísticas")
        
        self.mostrar_estadisticas(frame_stats)
    
    def crear_grafico_estados(self, parent):
        """Crea gráfico de estados en el frame padre"""
        df = self.leer_clientes()
        
        if df.empty:
            ttk.Label(parent, text="No hay datos para generar gráficos.").pack(pady=50)
            return
        
        fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(12, 10))
        fig.suptitle('ANÁLISIS VISUAL DE CLIENTES POTENCIALES', fontsize=16, fontweight='bold')
        
        # Gráfico 1: Estados de contacto
        estados_count = df['Estado_Contacto'].value_counts()
        colors1 = plt.cm.Set3(np.linspace(0, 1, len(estados_count)))
        ax1.pie(estados_count.values, labels=estados_count.index, autopct='%1.1f%%',
                colors=colors1, startangle=90)
        ax1.set_title('DISTRIBUCIÓN POR ESTADO DE CONTACTO')
        
        # Gráfico 2: Niveles de interés
        interes_count = df['Interes'].value_counts()
        colors2 = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4']
        bars = ax2.bar(interes_count.index, interes_count.values, 
                      color=colors2[:len(interes_count)])
        ax2.set_title('CLIENTES POR NIVEL DE INTERÉS')
        ax2.set_ylabel('Cantidad de Clientes')
        ax2.tick_params(axis='x', rotation=45)
        
        for bar in bars:
            height = bar.get_height()
            ax2.text(bar.get_x() + bar.get_width()/2., height,
                    f'{int(height)}', ha='center', va='bottom', fontweight='bold')
        
        # Gráfico 3: Top sectores
        sectores_count = df['Sector'].value_counts().head(8)
        colors3 = plt.cm.viridis(np.linspace(0, 1, len(sectores_count)))
        bars = ax3.barh(range(len(sectores_count)), sectores_count.values, color=colors3)
        ax3.set_title('TOP 8 SECTORES MÁS COMUNES')
        ax3.set_yticks(range(len(sectores_count)))
        ax3.set_yticklabels(sectores_count.index)
        ax3.set_xlabel('Cantidad de Clientes')
        
        for i, bar in enumerate(bars):
            width = bar.get_width()
            ax3.text(width, bar.get_y() + bar.get_height()/2.,
                    f'{int(width)}', ha='left', va='center', fontweight='bold')
        
        # Gráfico 4: Presencia web
        web_count = df['Sitio_Web_Actual'].value_counts()
        colors4 = ['#FF9999', '#66B3FF']
        ax4.pie(web_count.values, labels=web_count.index, autopct='%1.1f%%',
                colors=colors4, startangle=90)
        ax4.set_title('PRESENCIA WEB ACTUAL')
        
        canvas = FigureCanvasTkAgg(fig, parent)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def crear_grafico_propuestas(self, parent):
        """Crea gráficos para los nuevos campos de propuestas"""
        df = self.leer_clientes()
        
        if df.empty:
            ttk.Label(parent, text="No hay datos para generar gráficos.").pack(pady=50)
            return
        
        fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(12, 10))
        fig.suptitle('SEGUIMIENTO DE PROPUESTAS Y CLIENTES', fontsize=16, fontweight='bold')
        
        # Gráfico 1: Es cliente
        cliente_count = df['Es_Cliente'].fillna('NO').value_counts()
        colors1 = ['#FF9999', '#66B3FF']
        ax1.pie(cliente_count.values, labels=cliente_count.index, autopct='%1.1f%%',
                colors=colors1, startangle=90)
        ax1.set_title('CLIENTES ACTUALES')
        
        # Gráfico 2: Solicitud de propuestas
        solicitud_count = df['Solicito_Propuesta'].fillna('NO').value_counts()
        colors2 = ['#FF9999', '#66B3FF']
        bars2 = ax2.bar(solicitud_count.index, solicitud_count.values, color=colors2)
        ax2.set_title('SOLICITUD DE PROPUESTAS')
        ax2.set_ylabel('Cantidad')
        
        for bar in bars2:
            height = bar.get_height()
            ax2.text(bar.get_x() + bar.get_width()/2., height,
                    f'{int(height)}', ha='center', va='bottom', fontweight='bold')
        
        # Gráfico 3: Envío de propuestas
        envio_count = df['Se_Le_Envio_Propuesta'].fillna('NO').value_counts()
        colors3 = ['#FF9999', '#66B3FF']
        bars3 = ax3.bar(envio_count.index, envio_count.values, color=colors3)
        ax3.set_title('ENVÍO DE PROPUESTAS')
        ax3.set_ylabel('Cantidad')
        
        for bar in bars3:
            height = bar.get_height()
            ax3.text(bar.get_x() + bar.get_width()/2., height,
                    f'{int(height)}', ha='center', va='bottom', fontweight='bold')
        
        # Gráfico 4: Relación solicitud vs envío
        cross_data = df[df['Solicito_Propuesta'].notna() & df['Se_Le_Envio_Propuesta'].notna()]
        if not cross_data.empty:
            cross_tab = pd.crosstab(cross_data['Solicito_Propuesta'], cross_data['Se_Le_Envio_Propuesta'])
            cross_tab.plot(kind='bar', ax=ax4, color=['#FF9999', '#66B3FF'])
            ax4.set_title('SOLICITUD VS ENVÍO DE PROPUESTAS')
            ax4.set_ylabel('Cantidad')
            ax4.legend(title='Se envió propuesta')
            ax4.tick_params(axis='x', rotation=45)
        else:
            ax4.text(0.5, 0.5, 'No hay datos suficientes', ha='center', va='center', transform=ax4.transAxes)
            ax4.set_title('SOLICITUD VS ENVÍO DE PROPUESTAS')
        
        canvas = FigureCanvasTkAgg(fig, parent)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def crear_grafico_localidades(self, parent):
        """Crea gráfico de localidades"""
        df = self.leer_clientes()
        
        if df.empty:
            ttk.Label(parent, text="No hay datos para generar gráficos.").pack(pady=50)
            return
        
        localidades_count = df['Localidad'].value_counts().head(10)
        
        fig, ax = plt.subplots(figsize=(10, 6))
        colors = plt.cm.plasma(np.linspace(0, 1, len(localidades_count)))
        
        bars = ax.bar(localidades_count.index, localidades_count.values, color=colors)
        ax.set_title('TOP 10 LOCALIDADES CON MÁS CLIENTES POTENCIALES', fontweight='bold')
        ax.set_xlabel('Localidad')
        ax.set_ylabel('Cantidad de Clientes')
        plt.xticks(rotation=45, ha='right')
        
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height,
                   f'{int(height)}', ha='center', va='bottom', fontweight='bold')
        
        canvas = FigureCanvasTkAgg(fig, parent)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    
    def mostrar_estadisticas(self, parent):
        """Muestra estadísticas en formato texto"""
        df = self.leer_clientes()
        
        if df.empty:
            ttk.Label(parent, text="No hay datos para mostrar estadísticas.").pack(pady=50)
            return
        
        # Frame con scroll
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        text_stats = tk.Text(frame, yscrollcommand=scrollbar.set, wrap=tk.WORD, 
                           font=('Arial', 10), padx=10, pady=10)
        text_stats.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar.config(command=text_stats.yview)
        
        # Generar estadísticas
        total_clientes = len(df)
        sin_web = len(df[df['Sitio_Web_Actual'] == 'No tiene'])
        interes_alto = len(df[df['Interes'] == 'Alto'])
        es_cliente = len(df[df['Es_Cliente'] == 'SI'])
        solicito_propuesta = len(df[df['Solicito_Propuesta'] == 'SI'])
        envio_propuesta = len(df[df['Se_Le_Envio_Propuesta'] == 'SI'])
        
        stats_text = f"""
{'='*60}
        INFORME ESTADÍSTICO DE CLIENTES POTENCIALES
{'='*60}

📊 ESTADÍSTICAS PRINCIPALES:
   • Total de clientes: {total_clientes}
   • Clientes sin sitio web: {sin_web} ({sin_web/total_clientes*100:.1f}%)
   • Clientes con interés alto: {interes_alto}
   • Clientes actuales: {es_cliente} ({es_cliente/total_clientes*100:.1f}%)
   • Solicitaron propuesta: {solicito_propuesta} ({solicito_propuesta/total_clientes*100:.1f}%)
   • Se envió propuesta: {envio_propuesta} ({envio_propuesta/total_clientes*100:.1f}%)
   • Sectores únicos: {df['Sector'].nunique()}
   • Localidades únicas: {df['Localidad'].nunique()}

🎯 ESTADOS DE CONTACTO:
"""
        for estado, count in df['Estado_Contacto'].value_counts().items():
            porcentaje = (count / total_clientes) * 100
            stats_text += f"   • {estado}: {count} ({porcentaje:.1f}%)\n"

        stats_text += f"""
🏢 TOP SECTORES:
"""
        for sector, count in df['Sector'].value_counts().head(10).items():
            stats_text += f"   • {sector}: {count}\n"

        stats_text += f"""
📈 SEGUIMIENTO DE PROPUESTAS:
   • Clientes que solicitaron propuesta: {solicito_propuesta}
   • Propuestas enviadas: {envio_propuesta}
   • Tasa de conversión a cliente: {(es_cliente/max(envio_propuesta, 1))*100:.1f}%
"""

        text_stats.insert(tk.END, stats_text)
        text_stats.config(state=tk.DISABLED)

def main():
    root = tk.Tk()
    app = GestorClientesApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
