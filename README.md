# 📚 Sistema de Gestión de Biblioteca Escolar (Python + CustomTkinter)

Aplicación de escritorio desarrollada en **Python** para automatizar la gestión diaria de una biblioteca escolar: registro de préstamos, control de servicios/estancias, gestión de inventario, generación de KPIs y exportación de reportes a Excel.

> **Nota importante sobre los datos:** Los archivos CSV incluidos en este repositorio (`BD_NORMALIZADA.csv` y `Registro_Biblioteca.csv`) son **plantillas de ejemplo (datos ficticios)**. Se incluyen para demostrar el funcionamiento de la aplicación. Para usar el sistema en un entorno real, debes reemplazar estos archivos con tus propios datos.

---

## ✨ Características

### 📖 Módulo de Préstamos
- Registro rápido por folio de credencial.
- Autocompletado de nombre, grado y grupo.
- **Importación automática desde AppSheet** (formato CSV).

### 🏫 Módulo de Servicios / Estancia
- Registro de estancias, lectura, ajedrez, etc.
- **Búsqueda inteligente (fuzzy search)** por nombre con filtros.
- Registro de **eventos grupales**.
- Lista visible en vivo de servicios y exportación a Excel.

### 📦 Módulo de Inventario (NUEVO)
- **Clases Abstractas e Interfaces**: Uso de POO avanzada para definir productos.
- **Polimorfismo**: Manejo de diferentes tipos de productos (Mobiliario, Libros Detallados, Equipos de Cómputo) bajo una misma interfaz.
- Control de cantidades (conteo global) y catálogo detallado de libros.
- Búsqueda en tiempo real, modificación de stock y exportación de reportes a Excel.
- Persistencia automática en `inventario.json`.

### 📊 Módulo de Reportes y Analítica
- Índice Lector Global, por Grado y por Grupo.
- Top Lectores.
- **Lista de Deudores** con correo electrónico (exportable para campañas).
- Gráficas interactivas incrustadas con selector (Escuela / Grado / Grupo).

### ⚙️ Memoria y persistencia
- Guardado automático de los registros en un archivo JSON local.
- Los datos se conservan entre sesiones.

---

## 🚀 Instalación y uso

1. **Clonar el repositorio:**
   ```bash
   git clone https://github.com/Edmundo-GH/Biblioteca_ecolar_UC.git
   cd Biblioteca_ecolar_UC
