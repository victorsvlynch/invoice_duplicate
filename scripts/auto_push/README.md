# 🚀 Módulo de Automatización Git y Respaldos (`scripts/auto_push`)

Este directorio aloja las herramientas de automatización de control de versiones y empaquetado seguro para el proyecto **Sistema de Verificación de Facturas Duplicadas (DAG)**.

---

## 🛠️ Herramientas Incluidas

1. **`auto_push` (`auto_push.ps1` / `auto_push.bat` / `auto_push.py`)**:
   - Automatiza el ciclo `git add .`, generación de mensaje con IA, `git commit` y `git push`.
   - Utiliza la API de **Google Gemini** para analizar el `git diff` y generar un mensaje estructurado bajo el estándar **Conventional Commits** (`feat:`, `fix:`, `refactor:`, `chore:`).
   - Incluye mecanismo de reintento y fallback multi-modelo (`gemini-2.5-flash-lite`, `gemini-2.5-flash`, `gemini-flash-latest`).
   - Sincroniza con GitHub para que el pipeline de GitHub Actions (`release.yml`) detecte el prefijo y genere el **Release SemVer** automáticamente.

2. **`archify` (`archify.ps1` / `archify.bat`)**:
   - Empaquetador y respaldador estructurado del proyecto en formato ZIP con manifiesto JSON de trazabilidad.
   - **Modo Reportes (`-Mode Reports`):** Respalda reportes generados de verificación (`Reporte_Duplicados_*.xlsx`, `Resumen_Duplicados_*.csv`).
   - **Modo Snapshot Completo (`-Mode Full`):** Genera un snapshot comprimido del código fuente limpio, ignorando estrictamente `node_modules`, `__pycache__`, `venv`, `archives`, `.git` y temporales.
   - **Opción de Limpieza (`-CleanAfterArchive`):** Pone a salvo los archivos en el ZIP y purga los temporales locales.

---

## 💻 Instrucciones de Uso

### A. Ejecutar `auto_push`

```powershell
# Opción 1: PowerShell asistido por IA (genera el commit automáticamente)
.\scripts\auto_push\auto_push.ps1

# Opción 2: PowerShell con mensaje manual específico
.\scripts\auto_push\auto_push.ps1 -m "feat: optimizacion del motor matcher O(1)"

# Opción 3: Ejecutable Batch (1-clic en Windows)
scripts\auto_push\auto_push.bat
```

### B. Ejecutar `archify`

```powershell
# Generar un snapshot completo del código fuente
.\scripts\auto_push\archify.ps1 -Mode Full -Tag "release-v1.0"

# Respaldar reportes generados
.\scripts\auto_push\archify.ps1 -Mode Reports

# Simulación sin escritura (Dry-Run)
.\scripts\auto_push\archify.ps1 -Mode Full -DryRun

# Ejecución rápida vía Batch
scripts\auto_push\archify.bat Full
```
