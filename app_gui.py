"""
app_gui.py
----------
Interfaz CustomTkinter con:
- Préstamos (registro rápido + AppSheet)
- Servicios / Estancia (con lista visible y export a Excel)
- Inventario (Mobiliario, Libros detallados, Equipos) con búsqueda, modificación y exportación
- Reportes con selector Escuela/Grado/Grupo
- Memoria persistente
"""

from __future__ import annotations
import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox
import tkinter.ttk as ttk
from datetime import datetime
from pathlib import Path

import customtkinter as ctk
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

from motor_datos import (
    BibliotecaRepositorio, calcular_kpis, exportar_excel,
    calcular_kpis_servicios, construir_df_servicios,
)
from persistencia import cargar_repositorio, guardar_repositorio
# Importamos las clases del inventario (Actividad 4)
from modelos import Mobiliario, LibroInventario, EquipoComputo

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

COLOR_PRIM = "#1F497D"
COLOR_OK   = "#2E7D32"
COLOR_ERR  = "#C62828"
COLOR_BG   = "#F5F7FA"
ARCHIVO_INVENTARIO = Path("inventario.json")


class BibliotecaApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("📚 Sistema de Gestión de Biblioteca Escolar")
        self.geometry("1280x820")
        self.minsize(1100, 720)
        self.repo = BibliotecaRepositorio()
        self.kpis_cache = None
        self.kpis_serv_cache = None
        self.inventario = []  # Lista para el inventario

        self._msg_inicial = None
        if cargar_repositorio(self.repo):
            self._msg_inicial = (
                f"✔ Datos previos cargados: {len(self.repo.alumnos)} alumnos, "
                f"{len(self.repo.registros)} registros."
            )
        else:
            self._msg_inicial = "Sesión nueva. Carga el padrón (⚙️ Configuración)."

        # Cargar inventario desde JSON si existe
        self._cargar_inventario_desde_json()

        self._build_header()
        self._build_tabs()

        self.after(30_000, self._autoguardar)
        self.protocol("WM_DELETE_WINDOW", self._al_cerrar)

    # ================= PERSISTENCIA DEL INVENTARIO =================
    def _cargar_inventario_desde_json(self):
        """Carga el inventario desde el archivo JSON si existe."""
        if not ARCHIVO_INVENTARIO.exists():
            return
        try:
            with open(ARCHIVO_INVENTARIO, "r", encoding="utf-8") as f:
                datos = json.load(f)
            for item in datos:
                tipo = item.get("tipo", "")
                if tipo == "Mobiliario":
                    prod = Mobiliario(
                        item["id"], item["nombre"], item["descripcion"],
                        item["cantidad"], item.get("material", "")
                    )
                elif tipo == "Libro":
                    prod = LibroInventario(
                        item["id"], item["nombre"], item["descripcion"],
                        item["cantidad"], item.get("autor", ""),
                        item.get("isbn", ""), item.get("estado", "Disponible")
                    )
                elif tipo == "Equipo de Cómputo":
                    prod = EquipoComputo(
                        item["id"], item["nombre"], item["descripcion"],
                        item["cantidad"], item.get("marca", "")
                    )
                else:
                    continue
                self.inventario.append(prod)
        except Exception as e:
            print(f"Error cargando inventario: {e}")

    def _guardar_inventario_json(self):
        """Guarda el inventario en un archivo JSON."""
        datos = []
        for prod in self.inventario:
            item = {
                "id": prod.id_producto,
                "nombre": prod.nombre,
                "descripcion": prod.descripcion,
                "cantidad": prod.cantidad,
                "tipo": prod.tipo_producto(),
            }
            if isinstance(prod, LibroInventario):
                item["autor"] = prod.autor
                item["isbn"] = prod.isbn
                item["estado"] = prod.estado
            elif isinstance(prod, Mobiliario):
                item["material"] = prod.material
            elif isinstance(prod, EquipoComputo):
                item["marca"] = prod.marca
            datos.append(item)
        try:
            with open(ARCHIVO_INVENTARIO, "w", encoding="utf-8") as f:
                json.dump(datos, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Error guardando inventario: {e}")

    # ================= HEADER =================
    def _build_header(self):
        header = ctk.CTkFrame(self, fg_color=COLOR_PRIM, height=64, corner_radius=0)
        header.pack(fill="x")
        ctk.CTkLabel(header, text="📚 Biblioteca Escolar",
                     font=ctk.CTkFont(size=22, weight="bold"),
                     text_color="white").pack(side="left", padx=20, pady=14)
        self.lbl_estado = ctk.CTkLabel(
            header, text=self._texto_header(), text_color="white",
            font=ctk.CTkFont(size=13),
        )
        self.lbl_estado.pack(side="right", padx=20)

    def _texto_header(self):
        return (f"Padrón: {len(self.repo.alumnos)} alumnos · "
                f"Registros: {len(self.repo.registros)} · "
                f"Inventario: {len(self.inventario)} productos")

    def _actualizar_header(self):
        self.lbl_estado.configure(text=self._texto_header())

    # ================= TABS =================
    def _build_tabs(self):
        self.tabs = ctk.CTkTabview(self, fg_color=COLOR_BG)
        self.tabs.pack(fill="both", expand=True, padx=10, pady=10)
        self.tab_prestamos = self.tabs.add("📖 Préstamos")
        self.tab_servicios = self.tabs.add("🏫 Servicios / Estancia")
        self.tab_inventario = self.tabs.add("📦 Inventario")
        self.tab_reportes  = self.tabs.add("📊 Reportes y Analítica")
        self.tab_config    = self.tabs.add("⚙️ Configuración")

        self._build_prestamos(self.tab_prestamos)
        self._build_servicios(self.tab_servicios)
        self._build_inventario(self.tab_inventario)
        self._build_reportes(self.tab_reportes)
        self._build_config(self.tab_config)

        if self._msg_inicial:
            self._log(self._msg_inicial)

    # ================= CONFIG =================
    def _build_config(self, parent):
        frame = ctk.CTkFrame(parent)
        frame.pack(fill="x", padx=20, pady=20)
        ctk.CTkLabel(frame, text="Carga de archivos base",
                     font=ctk.CTkFont(size=16, weight="bold")
                     ).grid(row=0, column=0, columnspan=2, pady=10, sticky="w")

        ctk.CTkButton(frame, text="📂 Cargar Padrón de Alumnos (CSV)",
                      command=self._cargar_padron, fg_color=COLOR_PRIM
                      ).grid(row=1, column=0, padx=10, pady=8, sticky="w")
        self.lbl_padron = ctk.CTkLabel(frame, text="—")
        self.lbl_padron.grid(row=1, column=1, padx=10, sticky="w")

        ctk.CTkButton(frame, text="📥 Importar Préstamos AppSheet (CSV)",
                      command=self._importar_appsheet, fg_color=COLOR_PRIM
                      ).grid(row=2, column=0, padx=10, pady=8, sticky="w")
        self.lbl_hist = ctk.CTkLabel(frame, text="—")
        self.lbl_hist.grid(row=2, column=1, padx=10, sticky="w")

        ctk.CTkButton(frame, text="💾 Guardar registros manuales (CSV)",
                      command=self._guardar_registros, fg_color="gray30"
                      ).grid(row=3, column=0, padx=10, pady=8, sticky="w")

        ctk.CTkButton(frame, text="🗑️ Borrar toda la memoria (JSON)",
                      command=self._borrar_memoria, fg_color=COLOR_ERR
                      ).grid(row=4, column=0, padx=10, pady=8, sticky="w")

        ctk.CTkLabel(parent, text="Registro de actividad:",
                     font=ctk.CTkFont(size=13, weight="bold")
                     ).pack(anchor="w", padx=20, pady=(10, 0))
        self.txt_log = ctk.CTkTextbox(parent, height=200)
        self.txt_log.pack(fill="both", expand=True, padx=20, pady=10)

    # ================= PRÉSTAMOS =================
    def _build_prestamos(self, parent):
        box = ctk.CTkFrame(parent)
        box.pack(fill="x", padx=20, pady=(15, 10))
        ctk.CTkLabel(box, text="Registro Rápido de Préstamo",
                     font=ctk.CTkFont(size=16, weight="bold")
                     ).grid(row=0, column=0, columnspan=4, sticky="w", pady=8, padx=10)

        ctk.CTkLabel(box, text="Folio:").grid(row=1, column=0, padx=10, pady=6, sticky="e")
        self.entry_folio = ctk.CTkEntry(box, width=220, placeholder_text="Escanea/teclea y Enter")
        self.entry_folio.grid(row=1, column=1, padx=6, pady=6, sticky="w")
        self.entry_folio.bind("<Return>", lambda e: self._buscar_folio())

        ctk.CTkButton(box, text="🔍 Buscar", width=110,
                      command=self._buscar_folio, fg_color=COLOR_PRIM
                      ).grid(row=1, column=2, padx=6, pady=6)

        self.lbl_info = ctk.CTkLabel(box, text="—", font=ctk.CTkFont(size=14))
        self.lbl_info.grid(row=2, column=0, columnspan=4, padx=10, pady=10, sticky="w")

        ctk.CTkButton(box, text="✅ Registrar préstamo", height=38,
                      command=self._registrar_prestamo, fg_color=COLOR_OK
                      ).grid(row=3, column=1, padx=6, pady=8)
        ctk.CTkButton(box, text="📚 Registrar devolución", height=38,
                      command=self._registrar_devolucion, fg_color="gray30"
                      ).grid(row=3, column=2, padx=6, pady=8)

        imp = ctk.CTkFrame(parent)
        imp.pack(fill="x", padx=20, pady=10)
        ctk.CTkLabel(imp, text="Captura Automática (AppSheet)",
                     font=ctk.CTkFont(size=16, weight="bold")
                     ).pack(anchor="w", padx=10, pady=8)
        ctk.CTkButton(imp, text="📥 Adjuntar CSV de AppSheet",
                      command=self._importar_appsheet, fg_color=COLOR_PRIM
                      ).pack(anchor="w", padx=10, pady=(0, 10))

        ctk.CTkLabel(parent, text="Últimos movimientos:",
                     font=ctk.CTkFont(size=13, weight="bold")
                     ).pack(anchor="w", padx=20, pady=(10, 4))
        self.tabla_mov = ctk.CTkTextbox(parent, height=200)
        self.tabla_mov.pack(fill="both", expand=True, padx=20, pady=(0, 15))
        self._refrescar_tabla_movimientos()

    # ================= SERVICIOS =================
    def _build_servicios(self, parent):
        # --- Formulario ---
        box = ctk.CTkFrame(parent)
        box.pack(fill="x", padx=20, pady=10)
        ctk.CTkLabel(box, text="Registro de Servicio / Estancia",
                     font=ctk.CTkFont(size=16, weight="bold")
                     ).grid(row=0, column=0, columnspan=6, sticky="w", pady=8, padx=10)

        ctk.CTkLabel(box, text="Grado:").grid(row=1, column=0, padx=6, pady=6, sticky="e")
        self.cmb_grado = ctk.CTkComboBox(
            box, values=["", "1", "2", "3"], width=80,
            command=lambda _: self._actualizar_lista_alumnos(),
        )
        self.cmb_grado.grid(row=1, column=1, padx=4, pady=6, sticky="w")

        ctk.CTkLabel(box, text="Grupo:").grid(row=1, column=2, padx=6, pady=6, sticky="e")
        self.cmb_grupo = ctk.CTkComboBox(
            box, values=["", "A", "B", "C", "D", "E", "F"], width=80,
            command=lambda _: self._actualizar_lista_alumnos(),
        )
        self.cmb_grupo.grid(row=1, column=3, padx=4, pady=6, sticky="w")

        ctk.CTkLabel(box, text="Buscar alumno:").grid(row=2, column=0, padx=6, pady=6, sticky="e")
        self.entry_busqueda = ctk.CTkEntry(box, width=320, placeholder_text="Ej. Al → Alicia, Alberto…")
        self.entry_busqueda.grid(row=2, column=1, columnspan=2, padx=4, pady=6, sticky="w")
        self.entry_busqueda.bind("<KeyRelease>", lambda e: self._actualizar_lista_alumnos())

        self.lista_alumnos = tk.Listbox(box, height=5, width=60, font=("Segoe UI", 11))
        self.lista_alumnos.grid(row=3, column=0, columnspan=6, padx=10, pady=6, sticky="we")
        self.lista_alumnos.bind("<<ListboxSelect>>", self._seleccionar_alumno)
        self.alumnos_filtrados = []

        ctk.CTkLabel(box, text="Servicio:").grid(row=4, column=0, padx=6, pady=6, sticky="e")
        self.cmb_servicio = ctk.CTkComboBox(box, width=260, values=[
            "Trabajo al margen del grupo", "Lectura en sala", "Ajedrez",
            "Maestro responsable", "Evento grupal", "Otro",
        ])
        self.cmb_servicio.grid(row=4, column=1, columnspan=2, padx=4, pady=6, sticky="w")

        ctk.CTkLabel(box, text="Responsable:").grid(row=5, column=0, padx=6, pady=6, sticky="e")
        self.entry_responsable = ctk.CTkEntry(box, width=260)
        self.entry_responsable.grid(row=5, column=1, columnspan=2, padx=4, pady=6, sticky="w")

        ctk.CTkLabel(box, text="Evento grupal:").grid(row=6, column=0, padx=6, pady=6, sticky="e")
        self.entry_evento = ctk.CTkEntry(box, width=260, placeholder_text="Nombre del evento (opcional)")
        self.entry_evento.grid(row=6, column=1, padx=4, pady=6, sticky="w")
        ctk.CTkLabel(box, text="Personas:").grid(row=6, column=2, padx=6, pady=6, sticky="e")
        self.entry_personas = ctk.CTkEntry(box, width=80)
        self.entry_personas.insert(0, "1")
        self.entry_personas.grid(row=6, column=3, padx=4, pady=6, sticky="w")

        ctk.CTkButton(box, text="✅ Registrar servicio", height=38,
                      command=self._registrar_servicio, fg_color=COLOR_OK
                      ).grid(row=7, column=1, padx=6, pady=10)
        ctk.CTkButton(box, text="👥 Registrar evento grupal", height=38,
                      command=self._registrar_evento, fg_color=COLOR_PRIM
                      ).grid(row=7, column=2, padx=6, pady=10)

        self.lbl_sel = ctk.CTkLabel(box, text="Alumno seleccionado: —",
                                     font=ctk.CTkFont(size=13))
        self.lbl_sel.grid(row=8, column=0, columnspan=6, padx=10, pady=6, sticky="w")

        # --- Lista visible de servicios ---
        ctk.CTkLabel(parent, text="Servicios registrados (más reciente arriba):",
                     font=ctk.CTkFont(size=13, weight="bold")
                     ).pack(anchor="w", padx=20, pady=(6, 2))

        barra = ctk.CTkFrame(parent, fg_color="transparent")
        barra.pack(fill="x", padx=20, pady=(0, 4))
        ctk.CTkButton(barra, text="🔄 Actualizar lista", fg_color=COLOR_PRIM,
                      command=self._refrescar_servicios).pack(side="left", padx=4)
        ctk.CTkButton(barra, text="📤 Exportar servicios a Excel", fg_color=COLOR_OK,
                      command=self._exportar_servicios_excel).pack(side="left", padx=4)
        self.lbl_contador_serv = ctk.CTkLabel(barra, text="Total: 0 servicios")
        self.lbl_contador_serv.pack(side="left", padx=12)

        self.txt_servicios = ctk.CTkTextbox(parent, height=180)
        self.txt_servicios.pack(fill="both", expand=True, padx=20, pady=(0, 12))

        self._refrescar_servicios()

    # ================= INVENTARIO =================
    def _build_inventario(self, parent):
        main_frame = ctk.CTkFrame(parent, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # --- Panel Izquierdo: Formulario ---
        form_frame = ctk.CTkFrame(main_frame, width=350)
        form_frame.pack(side="left", fill="y", padx=(0, 10))
        form_frame.pack_propagate(False)

        ctk.CTkLabel(form_frame, text="➕ Agregar Producto",
                     font=ctk.CTkFont(size=16, weight="bold")
                     ).pack(pady=(10, 5), padx=10, anchor="w")

        # Tipo de producto
        ctk.CTkLabel(form_frame, text="Tipo:").pack(padx=10, anchor="w")
        self.cmb_tipo_prod = ctk.CTkComboBox(
            form_frame, values=["Mobiliario", "Libro Detallado", "Equipo de Cómputo"],
            command=self._cambiar_formulario_inventario
        )
        self.cmb_tipo_prod.pack(padx=10, pady=(0, 10), fill="x")

        # Campos comunes
        self.entry_inv_nombre = self._crear_campo(form_frame, "Nombre (ej. Sillas)")
        self.entry_inv_desc = self._crear_campo(form_frame, "Descripción")
        self.entry_inv_cant = self._crear_campo(form_frame, "Cantidad")

        # Frame para campos específicos
        self.frame_especifico = ctk.CTkFrame(form_frame, fg_color="transparent")
        self.frame_especifico.pack(fill="x", padx=10, pady=5)
        self.entry_especifico = None

        # Botón Guardar
        ctk.CTkButton(form_frame, text="✅ Guardar Producto",
                      command=self._agregar_producto_inventario,
                      fg_color=COLOR_OK, height=40
                      ).pack(padx=10, pady=20, fill="x")

        # --- Panel Derecho: Tabla ---
        table_frame = ctk.CTkFrame(main_frame)
        table_frame.pack(side="right", fill="both", expand=True)

        ctk.CTkLabel(table_frame, text="📋 Inventario General",
                     font=ctk.CTkFont(size=16, weight="bold")
                     ).pack(pady=10, anchor="w")

        # --- Barra de búsqueda ---
        search_frame = ctk.CTkFrame(table_frame, fg_color="transparent")
        search_frame.pack(fill="x", padx=10, pady=(5, 0))
        ctk.CTkLabel(search_frame, text="🔍 Buscar:").pack(side="left", padx=(0, 5))
        self.entry_buscar_inv = ctk.CTkEntry(search_frame, placeholder_text="Nombre, tipo o descripción...")
        self.entry_buscar_inv.pack(side="left", fill="x", expand=True)
        self.entry_buscar_inv.bind("<KeyRelease>", lambda e: self._refrescar_inventario())

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", rowheight=25, font=("Segoe UI", 10))
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))

        self.tabla_inv = ttk.Treeview(
            table_frame, columns=("ID", "Nombre", "Desc", "Cantidad", "Tipo", "Extra"),
            show="headings", height=15
        )
        self.tabla_inv.heading("ID", text="ID")
        self.tabla_inv.heading("Nombre", text="Nombre")
        self.tabla_inv.heading("Desc", text="Descripción")
        self.tabla_inv.heading("Cantidad", text="Cantidad")
        self.tabla_inv.heading("Tipo", text="Tipo")
        self.tabla_inv.heading("Extra", text="Extra")

        self.tabla_inv.column("ID", width=60, anchor="center")
        self.tabla_inv.column("Nombre", width=120)
        self.tabla_inv.column("Desc", width=150)
        self.tabla_inv.column("Cantidad", width=70, anchor="center")
        self.tabla_inv.column("Tipo", width=100)
        self.tabla_inv.column("Extra", width=100)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tabla_inv.yview)
        self.tabla_inv.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.tabla_inv.pack(fill="both", expand=True, padx=10, pady=10)

        # --- Botones de acción ---
        btn_frame = ctk.CTkFrame(table_frame, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=5)
        ctk.CTkButton(btn_frame, text="🔄 Refrescar", fg_color=COLOR_PRIM,
                      command=self._refrescar_inventario).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="✏️ Modificar Cantidad", fg_color="#E65100",
                      command=self._modificar_cantidad_inventario).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="🗑️ Eliminar", fg_color=COLOR_ERR,
                      command=self._eliminar_producto).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="📤 Exportar Reporte (Excel)", fg_color=COLOR_OK,
                      command=self._exportar_inventario_excel).pack(side="left", padx=5)

        self._refrescar_inventario()

    # --- Funciones auxiliares del inventario ---
    def _crear_campo(self, parent, placeholder):
        entry = ctk.CTkEntry(parent, placeholder_text=placeholder)
        entry.pack(padx=10, pady=5, fill="x")
        return entry

    def _cambiar_formulario_inventario(self, tipo):
        for widget in self.frame_especifico.winfo_children():
            widget.destroy()

        if tipo == "Libro Detallado":
            ctk.CTkLabel(self.frame_especifico, text="Autor:").pack(anchor="w")
            self.entry_especifico = ctk.CTkEntry(self.frame_especifico)
            self.entry_especifico.pack(fill="x", pady=(0, 5))
            ctk.CTkLabel(self.frame_especifico, text="ISBN:").pack(anchor="w")
            self.entry_isbn = ctk.CTkEntry(self.frame_especifico)
            self.entry_isbn.pack(fill="x")
        elif tipo == "Mobiliario":
            ctk.CTkLabel(self.frame_especifico, text="Material:").pack(anchor="w")
            self.entry_especifico = ctk.CTkEntry(self.frame_especifico)
            self.entry_especifico.pack(fill="x")
        elif tipo == "Equipo de Cómputo":
            ctk.CTkLabel(self.frame_especifico, text="Marca:").pack(anchor="w")
            self.entry_especifico = ctk.CTkEntry(self.frame_especifico)
            self.entry_especifico.pack(fill="x")

    def _agregar_producto_inventario(self):
        nombre = self.entry_inv_nombre.get().strip()
        desc = self.entry_inv_desc.get().strip()
        try:
            cantidad = int(self.entry_inv_cant.get())
        except ValueError:
            messagebox.showerror("Error", "La cantidad debe ser un número.")
            return
        if not nombre:
            messagebox.showerror("Error", "El nombre es obligatorio.")
            return

        tipo = self.cmb_tipo_prod.get()
        id_prod = f"P{len(self.inventario) + 1:04d}"

        try:
            if tipo == "Mobiliario":
                material = self.entry_especifico.get() if self.entry_especifico else ""
                nuevo = Mobiliario(id_prod, nombre, desc, cantidad, material)
            elif tipo == "Libro Detallado":
                autor = self.entry_especifico.get() if self.entry_especifico else ""
                isbn = self.entry_isbn.get() if hasattr(self, 'entry_isbn') else ""
                nuevo = LibroInventario(id_prod, nombre, desc, cantidad, autor, isbn, "Disponible")
            elif tipo == "Equipo de Cómputo":
                marca = self.entry_especifico.get() if self.entry_especifico else ""
                nuevo = EquipoComputo(id_prod, nombre, desc, cantidad, marca)
            else:
                messagebox.showerror("Error", "Tipo no válido.")
                return

            self.inventario.append(nuevo)
            self._refrescar_inventario()
            self._actualizar_header()
            self._guardar_inventario_json()  # Persistencia
            self._limpiar_formulario_inventario()
            messagebox.showinfo("Éxito", f"Producto '{nombre}' agregado.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _limpiar_formulario_inventario(self):
        self.entry_inv_nombre.delete(0, "end")
        self.entry_inv_desc.delete(0, "end")
        self.entry_inv_cant.delete(0, "end")
        if self.entry_especifico:
            self.entry_especifico.delete(0, "end")
        if hasattr(self, 'entry_isbn'):
            self.entry_isbn.delete(0, "end")

    def _refrescar_inventario(self):
        """Refresca la tabla del inventario, aplicando filtro de búsqueda si existe."""
        for item in self.tabla_inv.get_children():
            self.tabla_inv.delete(item)

        # Obtener texto de búsqueda (si existe el campo)
        filtro = ""
        if hasattr(self, 'entry_buscar_inv'):
            filtro = self.entry_buscar_inv.get().strip().lower()

        for prod in self.inventario:
            # Aplicar filtro
            if filtro:
                nombre = prod.nombre.lower()
                tipo = prod.tipo_producto().lower()
                desc = prod.descripcion.lower()
                if filtro not in nombre and filtro not in tipo and filtro not in desc:
                    continue

            # Extraer detalles específicos según el tipo
            extra = ""
            if isinstance(prod, LibroInventario):
                extra = f"Autor: {prod.autor}"
            elif isinstance(prod, Mobiliario):
                extra = f"Material: {prod.material}"
            elif isinstance(prod, EquipoComputo):
                extra = f"Marca: {prod.marca}"

            self.tabla_inv.insert("", "end", values=(
                prod.id_producto, prod.nombre, prod.descripcion,
                prod.cantidad, prod.tipo_producto(), extra
            ))

    def _eliminar_producto(self):
        sel = self.tabla_inv.selection()
        if not sel:
            messagebox.showwarning("Atención", "Selecciona un producto de la tabla.")
            return
        item = self.tabla_inv.item(sel[0])
        id_prod = item['values'][0]
        if not messagebox.askyesno("Confirmar", f"¿Eliminar el producto {id_prod}?"):
            return
        self.inventario = [p for p in self.inventario if p.id_producto != id_prod]
        self._refrescar_inventario()
        self._actualizar_header()
        self._guardar_inventario_json()  # Persistencia
        messagebox.showinfo("Eliminado", f"Producto {id_prod} eliminado.")

    def _modificar_cantidad_inventario(self):
        """Permite modificar la cantidad de un producto seleccionado."""
        sel = self.tabla_inv.selection()
        if not sel:
            messagebox.showwarning("Atención", "Selecciona un producto de la tabla.")
            return
        item = self.tabla_inv.item(sel[0])
        id_prod = item['values'][0]
        nombre_prod = item['values'][1]
        cantidad_actual = item['values'][3]

        ventana = ctk.CTkToplevel(self)
        ventana.title(f"Modificar cantidad: {nombre_prod}")
        ventana.geometry("400x220")
        ventana.grab_set()

        ctk.CTkLabel(ventana, text=f"Producto: {nombre_prod}",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(pady=10)
        ctk.CTkLabel(ventana, text=f"Cantidad actual: {cantidad_actual}").pack(pady=5)
        ctk.CTkLabel(ventana, text="Nueva cantidad:").pack(pady=(10, 0))
        entry_nueva = ctk.CTkEntry(ventana, width=150)
        entry_nueva.pack(pady=5)
        entry_nueva.insert(0, str(cantidad_actual))

        def guardar():
            try:
                nueva_cant = int(entry_nueva.get())
                if nueva_cant < 0:
                    messagebox.showerror("Error", "La cantidad no puede ser negativa.")
                    return
                for prod in self.inventario:
                    if prod.id_producto == id_prod:
                        prod.cantidad = nueva_cant
                        break
                self._refrescar_inventario()
                self._actualizar_header()
                self._guardar_inventario_json()  # Persistencia
                messagebox.showinfo("OK", f"Cantidad actualizada a {nueva_cant}.")
                ventana.destroy()
            except ValueError:
                messagebox.showerror("Error", "Ingresa un número válido.")

        ctk.CTkButton(ventana, text="✅ Guardar", command=guardar,
                      fg_color=COLOR_OK).pack(pady=15)

    def _exportar_inventario_excel(self):
        """Exporta el inventario completo a un archivo Excel."""
        if not self.inventario:
            messagebox.showinfo("Sin datos", "No hay productos en el inventario.")
            return

        escritorio = Path.home() / "Desktop"
        nombre = f"Inventario_Biblioteca_{datetime.now():%Y-%m-%d_%H%M}.xlsx"
        ruta = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialdir=str(escritorio if escritorio.exists() else Path.home()),
            initialfile=nombre,
            filetypes=[("Excel", "*.xlsx")],
        )
        if not ruta:
            return

        try:
            import pandas as pd
            from openpyxl.styles import PatternFill, Font, Alignment

            datos = []
            for prod in self.inventario:
                extra = ""
                if isinstance(prod, LibroInventario):
                    extra = f"Autor: {prod.autor} | ISBN: {prod.isbn}"
                elif isinstance(prod, Mobiliario):
                    extra = f"Material: {prod.material}"
                elif isinstance(prod, EquipoComputo):
                    extra = f"Marca: {prod.marca}"

                datos.append({
                    "ID": prod.id_producto,
                    "Nombre": prod.nombre,
                    "Descripción": prod.descripcion,
                    "Cantidad": prod.cantidad,
                    "Tipo": prod.tipo_producto(),
                    "Detalles Extra": extra,
                })

            df = pd.DataFrame(datos)
            with pd.ExcelWriter(ruta, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="Inventario", index=False)
                ws = writer.sheets["Inventario"]
                for c_idx in range(1, ws.max_column + 1):
                    cell = ws.cell(row=1, column=c_idx)
                    cell.fill = PatternFill("solid", start_color="1F497D")
                    cell.font = Font(color="FFFFFF", bold=True)
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                ws.column_dimensions['A'].width = 10
                ws.column_dimensions['B'].width = 25
                ws.column_dimensions['C'].width = 35
                ws.column_dimensions['D'].width = 12
                ws.column_dimensions['E'].width = 20
                ws.column_dimensions['F'].width = 40

            self._log(f"Inventario exportado a Excel: {ruta}")
            messagebox.showinfo("OK", f"Reporte exportado exitosamente:\n{ruta}")
        except PermissionError:
            messagebox.showerror("Archivo bloqueado",
                                 f"El archivo está abierto en Excel. Ciérralo y vuelve a intentar:\n{ruta}")
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"{type(e).__name__}: {e}")

    # ================= REPORTES =================
    def _build_reportes(self, parent):
        top = ctk.CTkFrame(parent)
        top.pack(fill="x", padx=20, pady=10)
        ctk.CTkButton(top, text="🔄 Calcular", fg_color=COLOR_PRIM,
                      command=self._calcular_reportes).pack(side="left", padx=6, pady=6)
        ctk.CTkButton(top, text="📤 Exportar Excel (préstamos + servicios)",
                      fg_color=COLOR_OK, command=self._exportar_excel
                      ).pack(side="left", padx=6, pady=6)
        ctk.CTkButton(top, text="📋 Filtrar deudores", fg_color="gray30",
                      command=self._filtrar_deudores).pack(side="left", padx=6, pady=6)

        self.kpi_frame = ctk.CTkFrame(parent, fg_color="transparent")
        self.kpi_frame.pack(fill="x", padx=20, pady=8)
        colores = ["#1F497D", "#2E7D32", "#E65100", "#6A1B9A"]
        self.kpi_labels = {}
        for i, (k, txt) in enumerate([
            ("alumnos", "👥 Alumnos"), ("prestamos", "📚 Préstamos"),
            ("indice", "📈 Índice"), ("circulacion", "🔄 En circulación"),
        ]):
            card = ctk.CTkFrame(self.kpi_frame, fg_color=colores[i], corner_radius=12)
            card.grid(row=0, column=i, padx=6, pady=6, sticky="we")
            self.kpi_frame.grid_columnconfigure(i, weight=1)
            ctk.CTkLabel(card, text=txt, font=ctk.CTkFont(size=12, weight="bold"),
                         text_color="white").pack(pady=(10, 0))
            lbl = ctk.CTkLabel(card, text="—", font=ctk.CTkFont(size=24, weight="bold"),
                               text_color="white")
            lbl.pack(pady=(0, 10))
            self.kpi_labels[k] = lbl

        selector = ctk.CTkFrame(parent, fg_color="transparent")
        selector.pack(fill="x", padx=20, pady=(4, 0))
        ctk.CTkLabel(selector, text="Ver métricas por:",
                     font=ctk.CTkFont(size=13, weight="bold")
                     ).pack(side="left", padx=(0, 10))
        self.vista_actual = ctk.StringVar(value="escuela")
        for valor, etiqueta in [
            ("escuela", "🏫 Escuela"), ("grado", "🎓 Grado"), ("grupo", "👥 Grupo"),
        ]:
            ctk.CTkRadioButton(selector, text=etiqueta, variable=self.vista_actual,
                               value=valor, command=self._render_grafica
                               ).pack(side="left", padx=8)

        self.fig = Figure(figsize=(7, 3.0), dpi=90)
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=parent)
        self.canvas.get_tk_widget().pack(fill="both", expand=True, padx=20, pady=8)

        ctk.CTkLabel(parent, text="Deudores activos (con correo):",
                     font=ctk.CTkFont(size=13, weight="bold")
                     ).pack(anchor="w", padx=20, pady=(6, 0))
        self.txt_deudores = ctk.CTkTextbox(parent, height=140)
        self.txt_deudores.pack(fill="both", expand=True, padx=20, pady=(0, 10))

    # ================= LÓGICA GENERAL =================
    def _log(self, msg: str):
        if hasattr(self, "txt_log"):
            self.txt_log.insert("end", f"[{datetime.now():%H:%M:%S}] {msg}\n")
            self.txt_log.see("end")

    def _autoguardar(self):
        try:
            guardar_repositorio(self.repo)
            self._guardar_inventario_json()  # Persistencia del inventario
        except Exception as e:
            print("Autoguardado falló:", e)
        self.after(30_000, self._autoguardar)

    def _al_cerrar(self):
        try:
            guardar_repositorio(self.repo)
            self._guardar_inventario_json()  # Persistencia del inventario
        except Exception as e:
            print("Guardado al cerrar falló:", e)
        self.destroy()

    def _borrar_memoria(self):
        if not messagebox.askyesno("Confirmar", "¿Borrar TODA la memoria guardada (incluyendo inventario)?"):
            return
        try:
            from persistencia import ARCHIVO_DATOS
            if ARCHIVO_DATOS.exists():
                ARCHIVO_DATOS.unlink()
            if ARCHIVO_INVENTARIO.exists():
                ARCHIVO_INVENTARIO.unlink()
            self.repo.alumnos.clear()
            self.repo.registros.clear()
            self.inventario.clear()
            self._actualizar_header()
            self._refrescar_servicios()
            self._refrescar_tabla_movimientos()
            self._refrescar_inventario()
            self._log("Memoria borrada (incluyendo inventario).")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _cargar_padron(self):
        ruta = filedialog.askopenfilename(filetypes=[("CSV", "*.csv")])
        if not ruta:
            return
        try:
            n = self.repo.cargar_padron(ruta)
            self.lbl_padron.configure(text=f"✔ {n} alumnos desde {os.path.basename(ruta)}")
            self._actualizar_header()
            self._log(f"Padrón cargado: {n} alumnos.")
            self._actualizar_lista_alumnos()
            guardar_repositorio(self.repo)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el padrón:\n{e}")

    def _importar_appsheet(self):
        ruta = filedialog.askopenfilename(filetypes=[("CSV", "*.csv")])
        if not ruta:
            return
        try:
            n = self.repo.cargar_historial(ruta, origen="AppSheet")
            self.lbl_hist.configure(text=f"✔ {n} préstamos importados")
            self._log(f"Importados {n} préstamos de AppSheet.")
            self._actualizar_header()
            self._refrescar_tabla_movimientos()
            guardar_repositorio(self.repo)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo importar:\n{e}")

    def _guardar_registros(self):
        ruta = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if not ruta:
            return
        try:
            self.repo.guardar_registros_csv(ruta)
            messagebox.showinfo("OK", "Registros guardados.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # ---------- PRÉSTAMOS ----------
    def _buscar_folio(self):
        folio = self.entry_folio.get().strip()
        if not folio:
            return
        alumno = self.repo.buscar_alumno(folio)
        if alumno:
            self.lbl_info.configure(
                text=f"✅ {alumno.nombre_completo} • {alumno.grado}°{alumno.grupo} • {alumno.correo}",
                text_color=COLOR_OK,
            )
            self._alumno_actual = alumno
        else:
            self.lbl_info.configure(text=f"⚠️ Folio {folio} no encontrado.",
                                     text_color=COLOR_ERR)
            self._alumno_actual = None

    def _registrar_prestamo(self):
        alumno = getattr(self, "_alumno_actual", None)
        if not alumno:
            messagebox.showwarning("Atención", "Primero busca un folio válido.")
            return
        self.repo.agregar_prestamo(alumno.folio, estado="Prestado")
        self._log(f"Préstamo registrado: {alumno.nombre_completo}")
        self._refrescar_tabla_movimientos()
        self._actualizar_header()
        guardar_repositorio(self.repo)
        self.entry_folio.delete(0, "end")
        self.lbl_info.configure(text="—", text_color="black")
        self._alumno_actual = None

    def _registrar_devolucion(self):
        folio = self.entry_folio.get().strip()
        if not folio:
            messagebox.showwarning("Atención", "Ingresa el folio para devolución.")
            return
        activos = [
            r for r in self.repo.registros
            if hasattr(r, "folio_alumno") and r.folio_alumno == folio
            and getattr(r, "estado", "").lower() == "prestado"
        ]
        if not activos:
            messagebox.showinfo("Sin préstamos", f"No hay préstamos activos para {folio}.")
            return
        activos[0].estado = "Devuelto"
        activos[0].fecha_devolucion = datetime.now()
        self._log(f"Devolución registrada para folio {folio}.")
        self._refrescar_tabla_movimientos()
        guardar_repositorio(self.repo)
        self.entry_folio.delete(0, "end")

    def _refrescar_tabla_movimientos(self):
        self.tabla_mov.delete("1.0", "end")
        for r in reversed(self.repo.registros[-20:]):
            self.tabla_mov.insert("end", str(r) + "\n")

    # ---------- SERVICIOS ----------
    def _actualizar_lista_alumnos(self):
        self.lista_alumnos.delete(0, "end")
        if not self.repo.alumnos:
            return
        grado = self.cmb_grado.get().strip()
        grupo = self.cmb_grupo.get().strip()
        texto = self.entry_busqueda.get().strip()
        self.alumnos_filtrados = self.repo.buscar_alumnos_fuzzy(texto, grado, grupo)
        for a in self.alumnos_filtrados[:150]:
            self.lista_alumnos.insert("end", f"{a.nombre_completo}  ·  {a.grado}°{a.grupo}")

    def _seleccionar_alumno(self, _evt=None):
        sel = self.lista_alumnos.curselection()
        if not sel:
            return
        alumno = self.alumnos_filtrados[sel[0]]
        self._alumno_sel = alumno
        self.lbl_sel.configure(
            text=f"Alumno seleccionado: {alumno.nombre_completo} ({alumno.grado}°{alumno.grupo})"
        )

    def _registrar_servicio(self):
        alumno = getattr(self, "_alumno_sel", None)
        if not alumno:
            messagebox.showwarning("Atención", "Selecciona un alumno de la lista.")
            return
        try:
            self.repo.agregar_servicio(
                tipo_servicio=self.cmb_servicio.get(),
                folio_alumno=alumno.folio,
                nombre_alumno=alumno.nombre_completo,
                grado=alumno.grado,
                grupo=alumno.grupo,
                responsable=self.entry_responsable.get().strip(),
                cantidad_personas=1,
            )
            self._log(f"Servicio '{self.cmb_servicio.get()}' → {alumno.nombre_completo}")
            self._actualizar_header()
            self._refrescar_servicios()
            guardar_repositorio(self.repo)
            messagebox.showinfo("OK", "Servicio registrado.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _registrar_evento(self):
        nombre_evento = self.entry_evento.get().strip()
        if not nombre_evento:
            messagebox.showwarning("Atención", "Escribe el nombre del evento.")
            return
        try:
            personas = int(self.entry_personas.get() or "1")
        except ValueError:
            personas = 1
        self.repo.agregar_servicio(
            tipo_servicio="Evento grupal",
            nombre_alumno=nombre_evento,
            responsable=self.entry_responsable.get().strip(),
            cantidad_personas=personas,
            origen="Evento",
        )
        self._log(f"Evento '{nombre_evento}' ({personas} personas).")
        self._actualizar_header()
        self._refrescar_servicios()
        guardar_repositorio(self.repo)
        messagebox.showinfo("OK", "Evento registrado.")
        self.entry_evento.delete(0, "end")

    def _refrescar_servicios(self):
        if not hasattr(self, "txt_servicios"):
            return
        self.txt_servicios.delete("1.0", "end")
        df = construir_df_servicios(self.repo)
        if df.empty:
            self.txt_servicios.insert("end", "(Aún no hay servicios registrados)\n")
            self.lbl_contador_serv.configure(text="Total: 0 servicios")
            return
        self.lbl_contador_serv.configure(
            text=f"Total: {len(df)} servicios · {df['Tipo_Servicio'].nunique()} tipos"
        )
        self.txt_servicios.insert(
            "end",
            f"{'FECHA/HORA':<20} | {'TIPO':<26} | {'ALUMNO / EVENTO':<28} | GRADO | RESPONSABLE\n"
        )
        self.txt_servicios.insert("end", "─" * 120 + "\n")
        for _, row in df.head(100).iterrows():
            tipo = str(row.get("Tipo_Servicio", ""))[:26]
            alumno = str(row.get("Nombre_Alumno", ""))[:28] or "—"
            grado = f"{row.get('Grado', '')}°{row.get('Grupo', '')}".strip("° ")
            if not grado or grado == "°":
                grado = "—"
            resp = str(row.get("Responsable", ""))[:18] or "—"
            fecha = str(row.get("Fecha_Hora", ""))[:19]
            self.txt_servicios.insert(
                "end",
                f"{fecha:<20} | {tipo:<26} | {alumno:<28} | {grado:<5} | {resp}\n"
            )
        if len(df) > 100:
            self.txt_servicios.insert("end", f"\n... y {len(df) - 100} más (exporta a Excel)\n")

    def _exportar_servicios_excel(self):
        df = construir_df_servicios(self.repo)
        if df.empty:
            messagebox.showinfo("Sin datos", "Aún no hay servicios registrados.")
            return
        escritorio = Path.home() / "Desktop"
        nombre = f"Servicios_Biblioteca_{datetime.now():%Y-%m-%d_%H%M}.xlsx"
        ruta = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialdir=str(escritorio if escritorio.exists() else Path.home()),
            initialfile=nombre,
            filetypes=[("Excel", "*.xlsx")],
        )
        if not ruta:
            return
        try:
            import pandas as pd
            from openpyxl.styles import PatternFill, Font, Alignment
            ks = calcular_kpis_servicios(self.repo)
            with pd.ExcelWriter(ruta, engine="openpyxl") as writer:
                ks["df"].to_excel(writer, sheet_name="Servicios_Listado", index=False)
                ks["por_tipo"].to_excel(writer, sheet_name="Por_Tipo", index=False)
                ks["por_grado"].to_excel(writer, sheet_name="Por_Grado", index=False)
                ks["top_alumnos"].to_excel(writer, sheet_name="Top_Alumnos", index=False)
                for hoja in writer.sheets:
                    ws = writer.sheets[hoja]
                    for c_idx in range(1, ws.max_column + 1):
                        cell = ws.cell(row=1, column=c_idx)
                        cell.fill = PatternFill("solid", start_color="1F497D")
                        cell.font = Font(color="FFFFFF", bold=True)
                        cell.alignment = Alignment(horizontal="center", vertical="center")
            self._log(f"Servicios exportados: {ruta}")
            messagebox.showinfo("OK", f"Reporte de servicios exportado:\n{ruta}")
        except PermissionError:
            messagebox.showerror("Archivo bloqueado",
                                 f"Ciérralo en Excel:\n{ruta}")
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"{type(e).__name__}: {e}")

    # ---------- REPORTES ----------
    def _calcular_reportes(self):
        try:
            self.kpis_cache = calcular_kpis(self.repo)
            self.kpis_serv_cache = calcular_kpis_servicios(self.repo)
        except Exception as e:
            messagebox.showerror("Error", f"Error calculando:\n{e}")
            return

        g = self.kpis_cache["global"].iloc[0]
        self.kpi_labels["alumnos"].configure(text=str(int(g["ALUMNOS_MATRICULADOS"])))
        self.kpi_labels["prestamos"].configure(text=str(int(g["PRESTAMOS_ANUAL"])))
        self.kpi_labels["indice"].configure(text=f"{g['INDICE_ANUAL']:.2f}")
        circul = int(self.kpis_cache["resumen"].iloc[0]["VALOR"])
        self.kpi_labels["circulacion"].configure(text=str(circul))

        self.txt_deudores.delete("1.0", "end")
        df_d = self.kpis_cache["deudores"]
        if df_d.empty:
            self.txt_deudores.insert("end", "(Sin deudores activos)\n")
        else:
            for _, row in df_d.head(50).iterrows():
                self.txt_deudores.insert(
                    "end",
                    f"{row.get('NOMBRE_COMPLETO','—')} | {row.get('GRADO','')}°{row.get('GRUPO','')} | "
                    f"{row.get('CORREO','—')} | {row.get('FECHA_SALIDA','—')}\n"
                )

        self._render_grafica()
        self._log("Reportes calculados.")

    def _render_grafica(self):
        if self.kpis_cache is None:
            return
        self.ax.clear()
        vista = self.vista_actual.get()
        if vista == "escuela":
            g = self.kpis_cache["global"].iloc[0]
            valores = [int(g["PRESTAMOS_ANUAL"])]
            etiquetas = ["Escuela"]
            titulo, color = "Préstamos Totales (Escuela)", "#1F497D"
        elif vista == "grado":
            df = self.kpis_cache["por_grado"]
            valores = df["PRESTAMOS_ANUAL"].tolist()
            etiquetas = [f"{g}°" for g in df["GRADO"]]
            titulo, color = "Préstamos por Grado", "#2E7D32"
        else:
            df = self.kpis_cache["por_grupo"]
            valores = df["PRESTAMOS_ANUAL"].tolist()
            etiquetas = [f"{r['GRADO']}°{r['GRUPO']}" for _, r in df.iterrows()]
            titulo, color = "Préstamos por Grupo", "#E65100"

        if not valores:
            self.ax.text(0.5, 0.5, "Sin datos", ha="center", va="center")
        else:
            barras = self.ax.bar(etiquetas, valores, color=color,
                                  edgecolor="white", linewidth=1.2)
            self.ax.set_title(titulo, fontsize=12, fontweight="bold")
            self.ax.set_ylabel("Préstamos")
            self.ax.spines["top"].set_visible(False)
            self.ax.spines["right"].set_visible(False)
            self.ax.grid(axis="y", alpha=0.3, linestyle="--")
            for b, v in zip(barras, valores):
                self.ax.text(b.get_x() + b.get_width() / 2, v, str(v),
                             ha="center", va="bottom", fontsize=9, fontweight="bold")
        self.fig.tight_layout()
        self.canvas.draw()

    def _filtrar_deudores(self):
        if self.kpis_cache is None:
            messagebox.showinfo("Info", "Calcula primero los reportes.")
            return
        ventana = ctk.CTkToplevel(self)
        ventana.title("Filtrar deudores")
        ventana.geometry("420x260")
        ctk.CTkLabel(ventana, text="Grado:").pack(pady=(15, 0))
        cmb_g = ctk.CTkComboBox(ventana, values=["", "1", "2", "3"]); cmb_g.pack()
        ctk.CTkLabel(ventana, text="Grupo:").pack(pady=(10, 0))
        cmb_gp = ctk.CTkComboBox(ventana, values=["", "A", "B", "C", "D", "E", "F"]); cmb_gp.pack()
        ctk.CTkLabel(ventana, text="Nombre contiene:").pack(pady=(10, 0))
        entry_n = ctk.CTkEntry(ventana, width=250); entry_n.pack()

        def aplicar():
            df = self.kpis_cache["deudores"].copy()
            if cmb_g.get(): df = df[df["GRADO"] == cmb_g.get()]
            if cmb_gp.get(): df = df[df["GRUPO"] == cmb_gp.get()]
            if entry_n.get():
                df = df[df["NOMBRE_COMPLETO"].str.contains(entry_n.get(), case=False, na=False)]
            self.txt_deudores.delete("1.0", "end")
            if df.empty:
                self.txt_deudores.insert("end", "(Sin deudores con ese filtro)\n")
            else:
                for _, row in df.head(100).iterrows():
                    self.txt_deudores.insert(
                        "end",
                        f"{row.get('NOMBRE_COMPLETO','—')} | {row.get('GRADO','')}°{row.get('GRUPO','')} | "
                        f"{row.get('CORREO','—')} | {row.get('FECHA_SALIDA','—')}\n"
                    )
            ventana.destroy()

        ctk.CTkButton(ventana, text="Aplicar", command=aplicar,
                      fg_color=COLOR_PRIM).pack(pady=15)

    def _exportar_excel(self):
        if self.kpis_cache is None:
            messagebox.showinfo("Info", "Calcula primero los reportes.")
            return
        escritorio = Path.home() / "Desktop"
        nombre = f"Analitica_Biblioteca_{datetime.now():%Y-%m-%d_%H%M}.xlsx"
        ruta = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialdir=str(escritorio if escritorio.exists() else Path.home()),
            initialfile=nombre,
            filetypes=[("Excel", "*.xlsx")],
        )
        if not ruta:
            return
        try:
            exportar_excel(self.kpis_cache, ruta, repo=self.repo)
            self._log(f"Excel exportado: {ruta}")
            messagebox.showinfo("OK", f"Reporte exportado:\n{ruta}")
        except PermissionError:
            messagebox.showerror("Archivo bloqueado",
                                 f"El archivo está abierto en Excel:\n{ruta}")
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error al exportar", f"{type(e).__name__}: {e}")


if __name__ == "__main__":
    app = BibliotecaApp()
    app.mainloop()