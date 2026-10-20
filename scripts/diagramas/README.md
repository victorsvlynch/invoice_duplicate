# 📐 Diagramas de Arquitectura y Flujos del Proceso (`scripts/diagramas`)

Este directorio centraliza los diagramas de arquitectura, modelos de flujo de datos y secuencias de interacción para el proyecto **Sistema de Verificación de Facturas Duplicadas (DAG)**.

---

## 📂 Recursos Disponibles

### 1. Diagrama de Arquitectura Interactivo con Archify
- **`arquitectura.html`**:
  - Renderizado autónomo de alta fidelidad compilado con **Archify CLI**.
  - Soporte de modo oscuro/claro, zoom continuo, pan interactivo y exportación (PNG/SVG).
  - Incluye **10 componentes** organizados en 3 zonas de seguridad y confianza:
    1. *Zona Pública de Cliente:* Operador Logístico (Navegador Web UI).
    2. *Zona de Ejecución Docker (:8504):* Docker Gateway (:8504), Streamlit App (`app.py`), Matcher Engine $O(1)$ (`matcher.py`), Excel Loader (Columna P) y Generador de Reportes (`report_generator.py`).
    3. *Zona Externa & Cloud:* SharePoint M365 (`Requ & OC.xlsx`), script `auto_push.ps1` y GitHub Actions (`release.yml`).
  - **Fichas técnicas y tarjetas de diseño** con detalles interactivos accesibles al hacer clic.
- **`arquitectura.json`**:
  - Especificación tipada estandarizada compatible con el esquema **Archify v1**.

### 2. Flujo de Secuencia y Verificación (Mermaid)
- **`flujo_verificacion_facturas.html`**:
  - Visualizador interactivo completo con soporte de tema claro/oscuro.
  - Describe el flujo completo desde la carga de `Requ & OC.xlsx` (local o SharePoint) hasta el motor de evaluación multicriterio y la exportación del reporte Excel.
- **`secuencia_verificacion.mmd`**:
  - Archivo fuente en notación Mermaid estándar para integración en herramientas de diagramado o documentación Markdown.

---

## 🚀 Cómo Re-renderizar o Actualizar con Archify

### En Windows (1-Clic):
```powershell
# PowerShell
.\scripts\diagramas\render_arquitectura.ps1

# O con apertura automática en navegador:
.\scripts\diagramas\render_arquitectura.ps1 -Open

# Símbolo del sistema (CMD / Batch):
scripts\diagramas\render_arquitectura.bat
```

---

## 👁️ Visualización Directa

Para abrir los diagramas:
1. Haz doble clic en [`arquitectura.html`](arquitectura.html) o en [`flujo_verificacion_facturas.html`](flujo_verificacion_facturas.html).
2. Ábrelos con cualquier navegador moderno (Chrome, Edge, Firefox).
