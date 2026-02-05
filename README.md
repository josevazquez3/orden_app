# Sistema de Órdenes del Día
## Colegio de Médicos de la Provincia de Buenos Aires

Sistema completo desarrollado en Python con Tkinter para gestionar las órdenes del día de las reuniones del Consejo Superior.

---

## 📋 Características

### ✨ Funcionalidades Principales

1. **Gestión de Reuniones**
   - Crear nuevas reuniones con todos los datos
   - Modalidad presencial o virtual
   - Selección de delegados titulares
   - Asignación de firmas (Presidente y Secretario)

2. **Gestión de Temas**
   - Alta, baja y modificación de temas
   - Categorización de temas
   - Historial completo de cada tema:
     * Fechas en las que se trató
     * Sedes donde se discutió
     * Cantidad de veces utilizado

3. **Gestión de Delegados**
   - Alta, baja y modificación
   - Titulares y suplentes
   - Asignación por distrito (I al X)

4. **Orden del Día**
   - Agregar temas
   - **Reordenar con botones Subir/Bajar**
   - Eliminar temas
   - Vista en tiempo real

5. **Generación de Documentos**
   - Vista previa
   - PDF profesional
   - Documento Word (.docx)
   - Formato idéntico al original

6. **Historial**
   - Todas las reuniones guardadas
   - Estadísticas por tema

---

## 🚀 Instalación

### Requisitos Previos
- Python 3.8 o superior
- pip (gestor de paquetes)

### Pasos de Instalación

1. **Descomprimir el proyecto**
```bash
# Descomprimir en una carpeta de su elección
```

2. **Instalar dependencias**
```bash
# En Windows:
pip install -r requirements.txt

# En Linux/Mac:
pip install -r requirements.txt --break-system-packages
```

3. **Ejecutar la aplicación**
```bash
python main.py
```

---

## 📁 Estructura del Proyecto
```
orden_dia_app/
│
├── models/
│   ├── __init__.py
│   └── database.py          # Base de datos SQLite
│
├── views/
│   ├── __init__.py
│   ├── main_view.py         # Ventana principal
│   └── dialogs.py           # Diálogos (nuevo tema, delegado, etc)
│
├── controllers/
│   ├── __init__.py
│   └── main_controller.py   # Controlador principal (MVC)
│
├── utils/
│   ├── __init__.py
│   └── document_generator.py # Generador de PDF y DOCX
│
├── main.py                   # Archivo de ejecución
├── requirements.txt          # Dependencias
├── README.md                # Esta documentación
└── orden_dia.db             # Base de datos (se crea automáticamente)
```

---

## 🎯 Uso del Sistema

### 1️⃣ Tab "Nueva Reunión"

**Crear una reunión:**
1. Completar datos (fecha, hora, lugar, sede, tipo)
2. Los delegados titulares se muestran automáticamente
3. Click en "➕ Agregar Tema" para agregar temas
4. Usar "⬆️ Subir" y "⬇️ Bajar" para reordenar
5. Seleccionar Presidente y Secretario
6. Click en "📄 Generar PDF" o "📝 Generar DOC"

### 2️⃣ Tab "Gestión de Temas"

**Crear tema:**
1. Click en "➕ Nuevo Tema"
2. Escribir descripción y categoría (opcional)
3. Guardar

**Ver historial de un tema:**
1. Seleccionar tema
2. Click en "📊 Ver Historial"

### 3️⃣ Tab "Gestión de Delegados"

**Crear delegado:**
1. Click en "➕ Nuevo Delegado"
2. Completar datos
3. Marcar si es titular
4. Guardar

### 4️⃣ Tab "Historial"

Ver todas las reuniones realizadas.

---

## 🔧 Características Técnicas

- **Arquitectura:** MVC (Modelo-Vista-Controlador)
- **Base de datos:** SQLite (archivo local)
- **Interfaz:** Tkinter
- **Generación PDF:** ReportLab
- **Generación Word:** python-docx

---

## 📝 Datos Iniciales

La primera vez que ejecutes el sistema, se cargarán automáticamente los 10 delegados titulares del documento original:

1. Dr. JULIO C. MORENO (Dist. I)
2. Dr. JORGE E. AGUGLIARO (Dist. II)
3. Dr. MAURICIO ESKINAZI (Dist. III)
4. Dr. RUBEN H. TUCCI (Dist. IV) - **Presidente por defecto**
5. Dr. JULIO D. DUNOGENT (Dist. V) - **Secretario por defecto**
6. Dr. JORGE OSCAR LUSARDI (Dist. VI)
7. Dr. HORACIO MARIO CARDUS (Dist. VII)
8. Dr. TOMAS GUANELLA (Dist. VIII)
9. Dr. GUSTAVO ARTURI (Dist. IX)
10. Dra. ROSA ANA DE FINO (Dist. X)

---

## 📂 Documentos Generados

Los documentos se guardan en la carpeta `outputs/`:
- Formato: `ORDEN_DEL_DIA_YYYYMMDD_HHMMSS.pdf`
- Formato: `ORDEN_DEL_DIA_YYYYMMDD_HHMMSS.docx`

---

## ❓ Preguntas Frecuentes

**P: ¿Cómo reordeno los temas?**  
R: Selecciona un tema en el orden y usa los botones "⬆️ Subir" o "⬇️ Bajar".

**P: ¿Se guardan las reuniones?**  
R: Sí, todas se guardan en la base de datos automáticamente.

**P: ¿Puedo modificar los delegados?**  
R: Sí, tanto en el tab "Gestión de Delegados" como al preparar una reunión.

**P: ¿Los temas se pueden reutilizar?**  
R: Sí, una vez creados están disponibles para todas las reuniones.

---

## 🆘 Solución de Problemas

**Error: "No module named 'tkinter'"**
- En Linux: `sudo apt-get install python3-tk`
- En Mac: tkinter viene con Python

**Error al generar PDF:**
- Verificar que reportlab esté instalado: `pip list | grep reportlab`

**Error al generar DOCX:**
- Verificar que python-docx esté instalado: `pip list | grep python-docx`

---

## 📧 Contacto

Sistema desarrollado por José - Secretario Administrativo  
Colegio de Médicos de la Provincia de Buenos Aires

---

**Versión:** 1.0.0  
**Fecha:** Enero 2026