# Manual de operación — Order Hub ISDIN Colombia

Este manual explica cómo está organizado el proyecto, cómo probar cambios sin afectar a
producción y cómo publicar una versión. Está escrito para quien no programa.

> Antes de hacer cualquier cosa, lee la sección **"Reglas que no se pueden romper"**.

---

## 1. Qué es esto

Order Hub es una aplicación web que corre dentro de Google (Apps Script). Automatiza el paso
de las órdenes de compra de los clientes a SAP:

1. Lee el archivo que el cliente deja en una carpeta de Google Drive (PDF o Excel).
2. Lo cruza contra las hojas maestras de Google Sheets (productos, bodega, destinatarios).
3. El usuario revisa el resultado en el asistente de la pantalla.
4. Se genera un archivo **CSV** con el formato de carga a SAP y se guarda en Drive.
5. Se registra la orden en la hoja `CONSOLIDADO` y se avisa a la KAM por correo.

Clientes: **MEDIPIEL** (PDF), **CMX** (Excel) y **CEN** (Excel consolidado con varios clientes finales).

---

## 2. Glosario (palabras que vas a ver)

| Palabra | Qué significa |
|---|---|
| **Apps Script** | El lugar de Google donde vive el código de la aplicación. Se abre desde script.google.com. |
| **Proyecto de Apps Script** | Una aplicación dentro de Apps Script. Tiene un **ID** (un código largo). |
| **Código fuente** | Los archivos `.js` y `.html` que escribimos. Están en la carpeta del proyecto. |
| **Repositorio / GitHub** | Donde se guarda el historial de cambios del código. |
| **Rama (branch)** | Una versión paralela del código. Por ejemplo `main`, `prueba`. |
| **Commit** | Una foto guardada de los cambios, con una descripción. |
| **clasp** | Herramienta que sube el código de tu computador a Apps Script. |
| **Push** | Subir el código de tu computador a Apps Script (`clasp push`). |
| **Implementación** | Una versión publicada de la aplicación que alguien puede abrir con un enlace. |
| **Versión** | Una copia fija del código dentro de Apps Script. Las implementaciones usan una versión fija. |
| **Terminal / PowerShell** | La ventana donde escribes comandos. |

---

## 3. Mapa de archivos

| Archivo | Para qué sirve |
|---|---|
| `main.js` | **Configuración central** (IDs de hojas y carpetas, clientes, KAM), el formato del CSV de SAP (`CSV_SAP`) y las funciones comunes a todos los clientes. |
| `medipiel_motor.js` | Lee los PDF de MEDIPIEL y guarda sus órdenes. |
| `cmx_motor.js` | Lee los Excel de CMX y guarda sus órdenes. |
| `cen_motor.js` | Lee el Excel consolidado de CEN (varios clientes) y guarda sus órdenes. |
| `wizard_ui.html` | El **asistente** que ve el usuario: fechas, materiales, back orders y guardado. |
| `index.html` | La pantalla principal con las tarjetas de cada cliente. |
| `appsscript.json` | Configuración de Apps Script (zona horaria, servicio de Drive). **No borrar.** |
| `CLAUDE.md` | Notas técnicas para la asistente de código. |
| `MANUAL.md` | Este manual. |

En `main.js` cada parte está marcada con comentarios. Busca las palabras **FORMATO ÚNICO**,
**PRUEBA** o **CONFIG** para encontrar lo importante.

---

## 4. Reglas que no se pueden romper

1. **Nunca subas código desde la carpeta `ORDER HUB` a la aplicación de producción sin probar antes.**
   Producción es la versión que usa el equipo de facturación.
2. **No compartas IDs ni correos reales en el repositorio.** El repositorio es privado.
3. **No borres archivos `.clasp.json` de producción.** Dicen a qué proyecto de Apps Script se sube el código.
4. **Antes de probar, revisa que estés en la rama `prueba`** (ver sección 6).
5. **No crees un `.clasp.json` dentro de la carpeta `ORDER HUB`.** Si existe, `clasp` puede subir el
   código a producción por error.
6. **En modo prueba los correos van a la dirección de la KAM de la copia.** Revísala antes (sección 7).
7. **En CEN, trabaja una persona a la vez.** Al iniciar un análisis CEN se borran las órdenes
   provisionales de cualquier sesión que no se haya confirmado con el CSV.

---

## 4.1. Flujo CEN con registro provisional

1. Al abrir CEN se analizan los Excel de la carpeta. Antes se borran las provisionales que
   quedaron de una sesión anterior.
2. Por cada cliente, al pulsar **"Guardar … y pasar al siguiente cliente"**, sus órdenes se
   escriben en `CONSOLIDADO` con la marca **PROVISIONAL** en la columna *Observaciones*. Así el
   siguiente cliente con las mismas iniciales continúa el consecutivo y no repite el número.
3. Al pulsar **"Generar CSV SAP unificado"**: se guarda el CSV, se quita la marca a las órdenes
   (quedan definitivas), se mueven los Excel terminados y se envían los correos.
4. Si el CSV no se puede guardar, las provisionales se borran. Si se recarga la página antes de
   generar el CSV, las provisionales se quedan en `CONSOLIDADO` hasta el siguiente análisis CEN,
   que las borra. **Apps Script no puede borrarlas en el momento exacto de la recarga.**

---

## 4.2. Destino de CMX

La ciudad y el destino de una orden de CMX salen de la hoja **Destinatarios**. El programa decide así:

1. Si el texto de la OC contiene el texto de la columna **Texto_OC** de un destino de CMX, usa ese destino.
2. Si CMX tiene un solo destino en Destinatarios, usa ese.
3. Si no, la orden queda **pendiente** y el asistente pide el **ID de destino** (columna `Id_destino`).
   - Con el ID se trae la ciudad de Destinatarios.
   - Si además se copia la dirección de la OC en el campo opcional, se guarda en **Texto_OC**, y la próxima OC con esa dirección se reconoce sola.

Para agregar un destino: una fila nueva en Destinatarios con `Id_origen` = 11033482, el `Id_destino`, la
ciudad y, en **Texto_OC**, el texto de la bodega tal como sale en la OC. Varias versiones del mismo
texto se separan con ` | `.

Nunca se asigna una ciudad por defecto. Si una orden queda pendiente, no se puede guardar hasta indicar el destino.

## 5. Herramientas que necesitas

- **Git** y **Node.js** (ya instalados en el equipo).
- **clasp** (ya instalado). Se ejecuta con `clasp.cmd` en PowerShell, porque `clasp` a secas
  está bloqueado por la política de scripts de Windows.
- Una cuenta con acceso a los archivos de prueba: `alejandra.martinez@isdin.com`.

---

## 6. Cómo probar cambios (entorno de prueba)

El entorno de prueba tiene su **propia copia** de la hoja de cálculo, sus propias carpetas
de Drive y su **propio proyecto de Apps Script**. Nada de lo que hagas ahí llega a producción.

Los IDs de prueba están en `main.js` (rama `prueba`):

| Qué | Dónde está |
|---|---|
| Hoja de cálculo (copia del consolidado) | `CONFIG.SPREADSHEET_ID` |
| Carpeta de pendientes MEDIPIEL | `CONFIG.CLIENTES.MEDIPIEL.FOLDER_ID` |
| Carpeta de pendientes CMX | `CONFIG.CLIENTES.CMX.FOLDER_ID` |
| Carpeta de pendientes CEN | `CONFIG.CLIENTES.CEN.FOLDER_ID` |
| Carpeta donde se guardan los CSV | `CONFIG.FOLDER_CSV_ID` |
| Modo prueba (prefijo `PRUEBA_` y correos apagados) | `CONFIG.PRUEBA` |

### 6.1. Preparar la carpeta de trabajo de pruebas

Para no mezclar los comandos con la carpeta original, la prueba se hace en una copia:
`C:\Users\AMSierra\Desktop\PRUEBA_ORDER_HUB`.

Si esa carpeta no existe, crearla así (PowerShell):

```powershell
cd "C:\Users\AMSierra\Desktop"
git clone --branch prueba --single-branch "C:\Users\AMSierra\Desktop\PROYECTOS\ORDER HUB\order-hub-colombia" "PRUEBA_ORDER_HUB"
```

### 6.2. Crear el proyecto de prueba en Apps Script (solo la primera vez)

```powershell
cd "C:\Users\AMSierra\Desktop\PRUEBA_ORDER_HUB"
git branch --show-current
clasp.cmd create --type standalone --title "ORDER HUB PRUEBA"
```

- `git branch --show-current` debe responder `prueba`.
- `create` escribe un archivo nuevo `.clasp.json` con el ID del proyecto de prueba.
- **No hagas `git add` de ese archivo.** Es propio de tu computador.

### 6.3. Subir el código de prueba

```powershell
clasp.cmd push
```

Si pregunta *"Manifest file has been updated. Do you want to push and overwrite?"*, responde **y**
(es un proyecto nuevo y vacío).

### 6.4. Abrir el proyecto y publicar la aplicación de prueba

```powershell
clasp.cmd open-script
```

En el editor que se abre:
1. Menú **Implementar → Nueva implementación**.
2. Tipo: **Aplicación web**.
3. **Ejecutar como:** Yo.
4. **Quién tiene acceso:** Solo yo.
5. **Implementar**. La primera vez pedirá **permisos**: elige tu cuenta y acepta.
6. Copia la **URL de la aplicación web** y ábrela en el navegador.

### 6.5. Cuando cambies código

Cada vez que cambies un archivo en `PRUEBA_ORDER_HUB`:
```powershell
clasp.cmd push
```
Luego recarga la URL de la aplicación. Si no ves el cambio, en el editor usa
**Implementar → Gestionar implementaciones → Editar → Nueva versión → Implementar**.

---

## 7. Modo prueba: qué cambia

Está en `main.js`, dentro de `CONFIG.PRUEBA`:

| Opción | Valor en prueba | Qué hace |
|---|---|---|
| `ACTIVO` | `true` | Activa el modo prueba. |
| `PREFIJO_ARCHIVOS` | `PRUEBA_` | Los archivos procesados y los CSV empiezan con `PRUEBA_`. Así se distinguen de los reales. |
| `CORREO_DESTINO` | vacío | Vacío: los correos van a la dirección de la KAM en `Asignacion_KAM` de la copia, con `[PRUEBA]` en el asunto. Si pones un correo, todos los avisos van ahí. |

Los archivos con `PRUEBA_` quedan en la carpeta de **procesados** real (es la misma en
producción y en prueba). **Bórralos a mano al terminar las pruebas.**

Antes de probar, revisa en la hoja de prueba, pestaña `Asignacion_KAM`, que **todos** los correos
sean tuyos. Si alguno no lo es, ese aviso llegaría a otra persona.

---

## 8. Lista de pruebas

### Prueba A — CSV sin encabezados y con el formato único
Procesa una orden de MEDIPIEL, una de CMX y una de CEN. Luego abre el CSV de
la carpeta **CSV PRUEBAS** y revisa:

- [ ] La primera línea empieza por `1;;ZDIR;Z011;10;00;...`. **No hay fila de encabezados.**
- [ ] El Contador y las columnas C a F salen solo en la primera línea de cada pedido.
- [ ] Las fechas van como `dd.MM.yyyy`, por ejemplo `26.10.2026`.
- [ ] La columna I (Num pedido cliente) es la **orden de compra** del cliente.
- [ ] La columna N (Material) sale vacía cuando la referencia no tiene material.
- [ ] El archivo se llama `PRUEBA_Pedido_...csv`.
- [ ] Compara con `Ejemplo estructura CSV para SAP.xlsx`.

### Prueba B — Desplegable de materiales en productos con stock OK
En el asistente de MEDIPIEL, con un producto que tenga stock:

- [ ] La **sección 2** (stock OK) muestra un desplegable con los lotes del producto.
- [ ] El lote preseleccionado es el primero con stock suficiente.
- [ ] Al cambiar de lote, el material enviado cambia (se ve en el resumen).
- [ ] Al elegir **"Digitar otro Material SAP manual"** aparece un campo de texto.
- [ ] Un material que **existe** (por ejemplo `690033678`) se acepta y muestra su stock real.
- [ ] Un material que **no existe** (por ejemplo `999999999`) se bloquea en rojo.
- [ ] La **sección 3** (back order) también tiene desplegable.

### Prueba D — Consecutivos de CEN con las mismas iniciales
Con dos clientes CEN de las mismas iniciales en el mismo Excel (en la copia, Cafam y Comfandi):
- [ ] Al guardar el primero, en CONSOLIDADO aparece su fila con la marca `PROVISIONAL` en *Observaciones*.
- [ ] El segundo cliente recibe el consecutivo siguiente, sin repetir el número.
- [ ] Al generar el CSV unificado, la columna *Observaciones* queda vacía en esas filas.
- [ ] Si recargas la página antes del CSV, al volver a abrir CEN las filas provisionales desaparecen.

### Prueba C — Renombrado y registro
- [ ] En **CONSOLIDADO** de la hoja de prueba aparecen las filas nuevas.
- [ ] El PDF/Excel procesado tiene el prefijo `PRUEBA_`.
- [ ] El consecutivo tiene la forma `AMS-26xxxx` (en la copia todas las iniciales KAM son AMS).

---

## 9. Problemas frecuentes

| Mensaje | Qué significa | Qué hacer |
|---|---|---|
| `clasp no se reconoce como...` | Se escribió `clasp` sin `.cmd`. | Usa `clasp.cmd`. |
| `la ejecución de scripts está deshabilitada` | Política de PowerShell. | Usa `clasp.cmd`, no `clasp`. |
| `Project file already exists` | Ya hay un `.clasp.json` en esa carpeta o en una carpeta superior. | Revisa la carpeta. En la carpeta de prueba no debe existir ninguno arriba. |
| `Unknown command "clasp open"` | Ese comando no existe. | Usa `clasp.cmd open-script`. |
| `Unknown command "clasp PUSH"` | Escrito en mayúsculas. | Escribe `push` en minúsculas. |
| `A file with this name already exists` | Conflicto de nombres en el proyecto. | Revisa el mensaje completo; no fuerces nada sin revisar. |
| Errores de permisos al abrir la URL | La primera vez Apps Script pide autorizar. | Implementar → Revisar permisos → tu cuenta → Permitir. |
| `Drive is not defined` al leer Excel de CEN | Falta el **servicio avanzado Drive API v2** en el proyecto. Lo usa `cen_convertirExcelASheet_` para convertir Excel en hoja de Google. | En el editor: **+ Servicios → Drive API → v2 → Añadir**. Si ya aparece, sube `appsscript.json` con `clasp.cmd push`. Luego crea una nueva versión. |

Si aparece un error que no está aquí, **copia el texto completo en rojo** de la terminal
y guárdalo en este manual en la sección 10 para revisarlo después.

---

## 10. Registro de incidencias

| Fecha | Qué pasó | Cómo se resolvió |
|---|---|---|
| 2026-10-28 | `clasp create` respondió "Project file already exists". Había un `.clasp.json` en la carpeta padre `ORDER HUB` apuntando a **producción**. | Se creó la carpeta `PRUEBA_ORDER_HUB` fuera de `ORDER HUB`. Producción no se modificó (comprobado: su ID de hoja sigue igual). |
| 2026-10-28 | Al leer la carpeta CEN de prueba: "Drive is not defined". | Falta activar el servicio avanzado Drive API v2 en el proyecto de prueba (ver tabla de problemas frecuentes). |
| 2026-10-28 | Dos clientes CEN con las mismas iniciales en el mismo Excel recibieron el mismo consecutivo (`AMS-260001` repetido). | Se registra cada cliente como provisional al guardarlo y el siguiente cliente continúa el número. Ver sección 4.1. |

---

## 11. Volver atrás si algo sale mal

- **En producción:** en el editor de Apps Script, **Implementar → Gestionar implementaciones**,
  y en la implementación de producción elige una **versión anterior** (por ejemplo la 55).
- **En el código:** cada cambio tiene un commit. Puedes ver el historial con `git log --oneline`.
  No borres commits sin revisar antes con quien corresponda.
