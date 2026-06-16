# 🧬 Plataforma de Gestión – Bitácora de Biológicos
**Jurisdicción / Distrito 06 Xonacatlán · CESAEM · Programa de Vacunación Universal**

---

## Descripción

Aplicación web desarrollada con **Streamlit + Pandas** para la gestión, control y visualización automatizada de la bitácora de productos biológicos. Cubre los 5 módulos solicitados:

| Módulo | Descripción |
|--------|-------------|
| 📊 Dashboard / KPIs | Indicadores globales: dosis recibidas, distribuidas, desperdiciadas y eficiencia |
| 📋 Bitácora General | Vista filtrable de entradas (Coordinación) y salidas (distribución) con exportación |
| 📄 Anexo A | Reporte institucional con encabezado oficial, filtros y descarga |
| 👩‍⚕️ Informe Enfermería | Tabla dinámica por responsable con totales por biológico |
| 🏭 Inventario por Lote | Saldos en tiempo real con alertas VENCIDO / RIESGO / VIGENTE |

---

## Requisitos del Sistema

- **Python 3.9 o superior**
- **pip** actualizado

## Instalación Local

```bash
# 1. Clonar o descomprimir el proyecto
cd bitacora_app

# 2. (Opcional pero recomendado) Crear entorno virtual
python -m venv venv
source venv/bin/activate        # Linux / macOS
venv\Scripts\activate           # Windows

# 3. Instalar dependencias
pip install -r requirements.txt

# 4. Ejecutar la aplicación
streamlit run app.py
```

La aplicación abrirá automáticamente en **http://localhost:8501**

---

## Estructura del Proyecto

```
bitacora_app/
├── app.py                  # Aplicación principal (todos los módulos)
├── requirements.txt        # Dependencias Python
├── README.md               # Este archivo
└── .streamlit/
    └── config.toml         # Tema y configuración de Streamlit
```

---

## Cómo Usar la Aplicación

### 1. Cargar el Archivo
- En el **panel lateral izquierdo**, haz clic en "📂 Cargar libro Excel"
- Arrastra y suelta (Drag & Drop) o selecciona el archivo `Bitacora_de_bioliguicos.xlsx`
- El sistema procesará automáticamente las pestañas: `Base entradas`, `Base salidas`, `ANEXO A FERNANDO`, `Informe`

### 2. Navegar por los Módulos
Usa el selector de radio en el panel izquierdo para cambiar entre módulos.

### 3. Filtros Disponibles
- **Biológico**: Selector múltiple o individual por vacuna (BCG, HEXAVALENTE, TDPa, etc.)
- **Rango de Fechas**: Calendario de inicio y fin para acotar los registros
- **Estado de Lote**: Vigente / Riesgo / Vencido (en módulo de inventario)

### 4. Exportaciones
Cada módulo tiene botones de descarga:
- **Excel (.xlsx)**: Tabla activa con formato profesional
- **PDF (.pdf)**: Reporte tabular en hoja A4 horizontal (requiere `reportlab`)

---

## Lógica de Cálculo – Inventario

```
SALDO_DISPONIBLE = DOSIS_RECIBIDAS - DOSIS_DISTRIBUIDAS
```

- **DOSIS_RECIBIDAS**: suma de `NUMERO DE DOSIS2` en `Base entradas`, agrupada por Biológico + Lote
- **DOSIS_DISTRIBUIDAS**: suma de la columna del biológico en `Base salidas` donde el lote coincide (columnas `LT [Biológico]`)
- **Alertas automáticas**:
  - 🔴 `VENCIDO`: Fecha de caducidad anterior a hoy
  - 🟡 `RIESGO`: Menos de 90 días para vencer
  - 🟢 `VIGENTE`: Más de 90 días de vigencia

---

## Pestañas del Excel Procesadas

| Pestaña | Fila de Encabezado | Uso |
|---------|-------------------|-----|
| `Base entradas` | Fila 5 (índice 4) | Registros de ingreso desde Coordinación |
| `Base salidas` | Fila 4 (índice 3) | Movimientos de distribución por destino |
| `ANEXO A FERNANDO` | Fila 13 (índice 12) | Reporte oficial por biológico |
| `Informe` | Fila 9 (índice 8) | Tabla dinámica por responsable |

---

## Despliegue en Streamlit Cloud (Gratis)

1. Sube el proyecto a un repositorio de **GitHub**
2. Ve a [share.streamlit.io](https://share.streamlit.io)
3. Conecta tu cuenta de GitHub
4. Selecciona el repositorio y el archivo `app.py`
5. Haz clic en **Deploy**

La URL pública estará disponible en minutos.

---

## Despliegue con Docker

```dockerfile
FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt
COPY . .
EXPOSE 8501
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
```

```bash
docker build -t bitacora-biologicos .
docker run -p 8501:8501 bitacora-biologicos
```

---

## Notas Técnicas

- El archivo Excel **no se almacena en el servidor**; se procesa en memoria por sesión.
- `@st.cache_data` evita reprocesar el archivo al navegar entre módulos.
- Si `reportlab` no está instalado, los botones de PDF no aparecen (la app sigue funcionando).
- El sistema es compatible con futuros archivos siempre que mantengan la misma estructura de pestañas y encabezados.

---

## Soporte

Para ajustes en columnas o nuevas pestañas, modificar el diccionario `col_map_e` / `col_map_a` en `app.py` según la estructura del nuevo archivo.
