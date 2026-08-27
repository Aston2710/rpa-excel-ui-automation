# Ambigüedad en la librería de UI Automation: dos implementaciones

**Fecha:** 2026-08-26
**Estado:** resuelto con dos ramas paralelas

## 1. El problema

El enunciado de la tarea (`README.md`) y la declaración de dependencias
(`pyproject.toml`) apuntan a **dos librerías distintas e incompatibles entre sí**.

### Evidencia a favor de `uiautomation`

`pyproject.toml` declara una única dependencia de ejecución:

```toml
dependencies = ["uiautomation>=2.0.29"]
```

Esa línea la introdujo el autor del enunciado en el commit `127ce7f`
(*"add: uiautomation"*), no es un artefacto de la implementación.

### Evidencia a favor de `pywinauto`

El texto del `README.md` contiene dos construcciones que **solo existen en
pywinauto**:

| Cita textual del README | Librería a la que pertenece |
|---|---|
| `child_window(control_type="...")` (§5, *Prohibición de Navegación Secuencial*) | pywinauto |
| "evaluar su presencia dinámicamente mediante el método `.exists()` de la ventana" (§4, Caso 02, paso 5) | pywinauto |

En `uiautomation` los equivalentes se escriben distinto: no existe
`child_window()` (se usan constructores tipados como `EditControl(AutomationId=...)`)
y el método de sondeo es `.Exists()` con mayúscula inicial.

### Por qué no se puede desempatar

La cronología de los commits del autor descarta la hipótesis de un descuido
temprano corregido después:

```
c2faec4  21-ago-2026 07:30   feat: initial commit with project setup and test plan
915e2b0  21-ago-2026 07:36   Update test case objectives to use 'bot' terminology
90f1bf4  21-ago-2026 07:38   Fix file paths in README for input and output
127ce7f  21-ago-2026 07:40   add: uiautomation          <-- se declara uiautomation
56b0ff5  21-ago-2026 07:45   Revise robustness criteria for file handling
912bbf6  21-ago-2026 07:53   Refine README for test cases and automation guidelines
                                                        <-- aparece child_window()
```

El ejemplo con sintaxis de pywinauto se agregó **13 minutos después** de declarar
`uiautomation`, en la última edición del enunciado. El autor tocó el README
sabiendo qué dependencia había fijado y aun así escribió el ejemplo de la otra
librería. No hay forma de deducir cuál de las dos señales es la intencional.

## 2. La decisión

Se entregan **ambas implementaciones**, con la misma arquitectura y los mismos
criterios de aceptación, en ramas separadas:

| Rama | Librería | Justificación |
|---|---|---|
| `main` | `uiautomation` | Es lo que declara `pyproject.toml`, el único artefacto ejecutable y verificable del enunciado. |
| `impl/pywinauto` | `pywinauto` | Es lo que sugiere la redacción literal del `README.md`, la última pieza que el autor editó. |

`main` es la entrega principal porque una dependencia declarada es un contrato
verificable (`pdm install` la instala), mientras que un ejemplo dentro de un
texto es ilustrativo. La rama alterna existe para que, si el criterio de
evaluación resulta ser el opuesto, la entrega no dependa de haber acertado la
adivinanza.

## 3. Qué cambia y qué no entre ramas

**No cambia:**

- La arquitectura: `ExcelManager` (aplicación y atajos globales) y
  `FileExplorer` (diálogos nativos y advertencias modales).
- Los `AutomationId` de los controles, obtenidos inspeccionando el árbol real de
  UI Automation de Office 16 en español. Son propiedades del sistema operativo,
  no de la librería cliente.
- El cumplimiento de los criterios de §5: sin `time.sleep`, sin `Tab`, sin
  coordenadas, `pathlib` exclusivo, logs en ambas clases.
- El ejecutor de los casos 01 y 02 y la verificación de integridad del origen.

**Cambia solo la capa de acceso a la UI:**

| Operación | `main` (uiautomation) | `impl/pywinauto` |
|---|---|---|
| Lanzar Excel | `subprocess.Popen` + `WindowControl(ClassName="XLMAIN")` | `Application(backend="uia").start(...)` |
| Localizar control | `dialog.EditControl(AutomationId="1148")` | `dialog.child_window(auto_id="1148", control_type="Edit")` |
| Esperar aparición | `.Exists(timeout, intervalo)` | `.exists(timeout=...)` / `.wait("exists ready")` |
| Esperar cierre | `.Disappears(timeout, intervalo)` | `.wait_not("exists", timeout=...)` |
| Accionar botón | `GetPattern(InvokePattern).Invoke()` | `.invoke()` |
| Escribir ruta | `GetPattern(ValuePattern).SetValue(...)` | `.set_edit_text(...)` |
| Atajo global | `auto.SendKeys("{Ctrl}{F12}")` | `send_keys("^{F12}")` |

## 4. Cómo ejecutar cada versión

```bash
# Entrega principal (uiautomation)
git checkout main
pdm install
pdm run bot          # Casos 01 y 02
pdm run test         # Pruebas unitarias

# Entrega alterna (pywinauto)
git checkout impl/pywinauto
pdm install
pdm run bot
pdm run test
```

## 5. Nota adicional sobre el flujo de apertura

Independiente de la librería, el cuadro de diálogo "Abrir" se invoca con el
atajo global `Ctrl+F12` y no navegando el Backstage (Archivo → Abrir →
Examinar), como sugieren las capturas de `img/openfile.png`.

Motivo: el ítem "Examinar" del Backstage está virtualizado y no se expone de
forma estable en el árbol de UI Automation de esta compilación de Office, lo que
lo convertiría precisamente en el tipo de automatización frágil que §5 prohíbe.
El uso del atajo está autorizado de forma explícita por el enunciado: *"envío de
teclas de método abreviado global (ej. F12)"*.
