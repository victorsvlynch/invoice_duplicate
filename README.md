# 📑 Sistema de Verificación de Facturas Duplicadas (Streamlit + Docker)

Sistema web interactivo desarrollado con **Streamlit** y contenedorizado con **Docker** para verificar facturas y comprobantes contra la **Columna P** del archivo Excel corporativo **`Requ & OC.xlsx`**.

---

## 🚀 Características Principales

1. **Búsqueda Flexible y Multicriterio**:
   - **Coincidencia Exacta**: Busca coincidencia idéntica de Serie y Correlativo (ej. `F85S-0000064535`).
   - **Por Correlativo con ceros**: Identifica comprobantes coincidiendo la numeración con ceros a la izquierda (ej. `0000064535`).
   - **Por Número sin ceros**: Identifica comprobantes ignorando los ceros a la izquierda (ej. `64535`).
   - **Alerta de Serie Diferente**: Detecta si el número correlativo ya existe registrado bajo otra serie distinta (ej. `F001-0000064535` vs `F85S-0000064535`), previniendo errores de digitación de operadores.
   - **Detección de Duplicados Internos**: Alerta si en la misma lista pegada por el usuario un comprobante fue ingresado dos o más veces.

2. **Carga y Sincronización del Archivo Excel (`Requ & OC.xlsx`)**:
   - **Opción A (Arrastrar y Soltar)**: Sube directamente el archivo Excel desde la interfaz web. Se guarda automáticamente en memoria local para no tener que subirlo cada vez.
   - **Opción B (Carpeta Sincronizada `data/`)**: Si tienes sincronizado SharePoint / OneDrive en tu computadora, puedes vincular o copiar el archivo a `data/Requ & OC.xlsx` y el contenedor Docker lo leerá en tiempo real.
   - **Opción C (Descarga desde SharePoint)**: Permite ingresar el enlace directo o configurar cookies de sesión.

3. **Detección Inteligente de Columnas**:
   - Selecciona automáticamente la **Columna P** (16ª columna, índice 15) del Excel.
   - Permite elegir cualquier otra columna o cambiar la fila de encabezados si el formato del archivo varía.

4. **Reportes y Exportación**:
   - Métricas en tiempo real: Total Consultados, Duplicados en Excel, Nuevos / No Encontrados.
   - Descarga de reporte completo en **Excel estilizado (.xlsx)** con pestañas de *Resumen*, *Detalle de Duplicados* y *No Encontrados*.
   - Exportación rápida en CSV.
   - Botón de un solo clic para copiar la lista limpia de comprobantes no duplicados listos para registrar.

---

## 🐳 Instrucciones de Ejecución con Docker

### 1. Iniciar el contenedor con Docker Compose:

Abre tu terminal (PowerShell o CMD) en esta carpeta y ejecuta:

```bash
docker compose up --build -d
```

### 2. Abrir la aplicación:

Abre tu navegador web en:
👉 **[http://localhost:8504](http://localhost:8504)**

### 3. Detener el contenedor:

```bash
docker compose down
```

---

## 💻 Ejecución Local (Sin Docker)

Si prefieres ejecutarlo directamente en tu entorno Python local:

```bash
# 1. Instalar dependencias
pip install -r requirements.txt

# 2. Iniciar Streamlit
streamlit run app.py
```

---

## 📂 Estructura del Proyecto

```text
Invoice_Duplicate/
├── .dockerignore
├── Dockerfile                   # Imagen ligera Python 3.11-slim
├── docker-compose.yml           # Orquestación con volumen ./data montado
├── requirements.txt             # Streamlit, Pandas, Openpyxl, Requests
├── README.md                    # Documentación completa
├── app.py                       # Aplicación principal Streamlit
├── .streamlit/
│   └── config.toml              # Tema y configuración de tamaño de subida (250MB)
├── src/
│   ├── __init__.py
│   ├── matcher.py               # Motor de indexación O(1) y coincidencia multicriterio
│   ├── excel_loader.py          # Conexión SharePoint, lector y selector de Columna P
│   └── report_generator.py      # Generador de reportes Excel estilizados
└── data/                        # Volumen persistente para 'Requ & OC.xlsx'
```

---

## 📌 Ejemplo de Uso

1. Abre la aplicación en [http://localhost:8504](http://localhost:8504).
2. Sube tu archivo `Requ & OC.xlsx` en la barra lateral izquierda (o usa el archivo ya existente en `data/`).
3. Haz clic en el botón **"📋 Cargar Ejemplo de Prueba"** o pega tu lista de facturas:
   ```text
   F85S-0000064535
   F85S-0000064619
   F72P-0000047299
   F73P-0000016233
   F75P-0000097864
   0000064535
   64535
   ```
4. Presiona **"🔍 Verificar Duplicados en Excel"**.
5. Revisa los resultados en las pestañas interactivas y descarga el reporte en Excel.
