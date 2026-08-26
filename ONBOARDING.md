# CONTEXTO DEL PROYECTO — Ferrocentral / Jony

## 📁 Archivos del proyecto

| Archivo | Ruta | Descripción |
|---|---|---|
| **ferrocentral-control.html** | `C:\Users\jony9\Downloads\ferrocentral-control.html` | App principal de bodega — inventario INGRESO/SALIDA |
| **tren-carga.html** | `C:\Users\jony9\Downloads\Jony 11-25\tren-carga.html` | App de KPIs de viajes de tren CIE-LDR ↔ LDR-CIE |

---

## 🏭 APP 1: ferrocentral-control.html

### Qué es
PWA (Progressive Web App) de control de bodega para la Estación México · Bodega IDEMA · La Dorada. Archivo HTML único con todo el CSS y JS incluido. Se instala en Android via Chrome → "Agregar a pantalla de inicio".

### Datos
- **localStorage key:** `ferrocentral_v2` (registros), `ferrocentral_conf` (configuración)
- **DEMO array:** 325 registros hardcodeados del Excel original como baseline
- **Servidor local:** Python HTTP server en `C:\Users\jony9\Downloads` puerto 8080
- **URL local:** `http://10.193.94.95:8080/ferrocentral-control.html`

### Modelo de datos
Dos tipos de registro por operación:
- **INGRESO** (`tipo='INGRESO'`, `saldo=cantidadRecibidaKg`) — mercancía que entra
- **SALIDA** (`tipo='SALIDA'`, `cantidadDespachoKg=cant`, `ingresoRef=idIngreso`) — mercancía despachada

Campos clave: `id, tipo, cliente, producto, fechaRecibido, fechaDespacho, cantidadRecibidaKg, cantidadDespachoKg, saldo, estadoFinal, tipoTransporteEntrada, tipoTransporteSalida, locoPlacaEntrada, plataformaEntrada, locoPlacaSalida, plataformaSalida, origen, destino, semana, batchId`

### Funciones implementadas (todo lo del viernes + esta semana)

#### ✅ Registro INGRESO
- CAMIÓN: formulario simple (un ingreso por vez)
- TREN: formulario bulk con múltiples plataformas (`.bulk-item`)

#### ✅ Registro SALIDA TREN — estructura DOS NIVELES (último cambio grande)
- **Plataforma → Cargas** (nueva estructura)
- Cada plataforma (`.plat-card`) contiene múltiples cargas (`.cargo-item`)
- Botón "➕ Agregar carga a esta plataforma" por plataforma
- Botón "➕ Agregar plataforma" al fondo
- La locomotora es OPCIONAL al registrar (se puede agregar después)
- `batchId` = `Date.now()` asignado a todos los registros del mismo despacho
- Funciones: `addPlatItem()`, `removePlatItem()`, `addCargoToPlat(pid)`, `removeCargoItem(cargoId, pid)`, `updateCargoHint(sel)`, `_syncOptionLabels()`

#### ✅ Desplegables de stock inteligentes
- Solo muestran registros con `saldo > 0`
- Etiqueta: `🟢 Disponible: X kg` (completo) o `🟡 Disponible: X de Y kg` (parcial)
- `_syncOptionLabels()` actualiza en tiempo real las etiquetas en TODOS los dropdowns cuando cambia una selección o cantidad
- `🔴 Sin disponible` cuando otra carga en el mismo formulario ya tomó todo el stock
- Opciones NUNCA se deshabilitan (bug Android: al deshabilitar una opción seleccionada, Chrome resetea el valor del select a vacío)
- Resumen verde al tope: `📊 Stock disponible — Total: X kg`

#### ✅ Actualizar despacho de tren (en ⚙️ Sync)
- Busca por fecha + destino todos los SALIDA TREN de ese día
- Permite agregar locomotora y hora de salida a TODOS los registros a la vez
- Útil cuando no se tienen los datos de locomotora al momento de registrar

#### ✅ Importar Excel
- SheetJS (cdn `xlsx-0.20.3`) en `<head>`
- Detecta columnas automáticamente
- Deduplicación por `_recKey()`: `fecha|cliente|producto|kg|tipo`
- Botón en modal ⚙️ Sync

#### ✅ Editar registros
- Botón ✏️ Editar en cada tarjeta de registro
- Modal `m-edit` con todos los campos editables
- Recalcula `saldo` al editar cantidad de un INGRESO
- `guardarEdit()` llama `refreshSalDropdowns()` al terminar

#### ✅ Otras funciones
- `borrarReg(id)`: restaura saldo del INGRESO padre al borrar una SALIDA, botón DESHACER 8 seg
- `limpiarCorruptos()`: elimina registros con 0 kg o campos undefined
- `refreshSalDropdowns()`: reconstruye todos los dropdowns de stock al borrar/restaurar/editar
- `toggleRec(el)`: expand/collapse tarjeta, debounce 400ms, `onclick="toggleRec(this)"`
- `_submitting` flag: previene doble-tap (2 seg cooldown)
- Búsqueda: cubre 18 campos incluyendo locoPlacaEntrada, plataformaEntrada, etc.
- `isoWeek(fecha)`: calcula semana ISO

### Bugs conocidos resueltos
- `borrarReg` duplicado → segunda definición eliminada
- `_recKey` con `tipo=undefined` en registros DEMO → usa `(r.tipo||'INGRESO')`
- Opciones disabled en Android resetean valor del select → nunca se deshabilitan
- `toggleRec` con IDs inusuales → pasa `this` directamente desde onclick
- Cross-item validation en SALIDA TREN: suma totales por ingreso antes de validar

### Estructura de archivos CSS clave
```
.plat-card       → tarjeta de plataforma (fondo azul claro, borde dashed)
.cargo-item      → item de carga dentro de una plataforma
.cargo-add-btn   → botón "➕ Agregar carga a esta plataforma"
.bulk-item       → usado por INGRESO TREN (no por SALIDA)
.ci-ref          → select de ingreso a descontar en cada cargo-item
.ci-cant         → input de cantidad en cada cargo-item
.ci-hint         → div de pista/disponibilidad en cada cargo-item
```

---

## 🚂 APP 2: tren-carga.html

### Qué es
Dashboard de KPIs para viajes de tren intermodal Ciénaga ↔ La Dorada. Archivo HTML único standalone. Funciona offline.

### Datos
- **localStorage key:** `tren_v5`
- **Array DATOS (hardcodeado):** datos históricos de Int 1 a Int 6 CIE-LDR + Int 1 a Int 5 LDR-CIE
- **Migración:** `migrar()` IIFE convierte v2/v3/v4 a v5. `repararNombres()` IIFE corrige doble-rename.

### Estado actual de los viajes (en DATOS)
| ID | Tren | Corredor | Fecha salida | Fecha llegada | Plat. | Ton. |
|---|---|---|---|---|---|---|
| 1001 | Int 1 | CIE-LDR | 03/03/2026 20:19 | 05/03/2026 03:00 | 37/37 | 1.183,76 |
| 1002 | Int 2 | CIE-LDR | 17/03/2026 15:38 | 18/03/2026 18:18 | 35/35 | 1.167,31 |
| 1003 | Int 3 | CIE-LDR | 29/03/2026 14:02 | 30/03/2026 12:50 | 29/29 | 964,87 |
| 1004 | Int 4 | CIE-LDR | 19/04/2026 07:20 | 20/04/2026 18:55 | 33/33 | 1.092,75 |
| 1005 | Int 5 | CIE-LDR | 06/05/2026 13:00 | 07/05/2026 10:55 | 33/33 | 984,13 |
| 1006 | Int 6 | CIE-LDR | 14/05/2026 19:48 | 16/05/2026 06:20 | 32/31 | 980,24 |
| 2001 | Int 1 | LDR-CIE | 11/03/2026 10:30 | 13/03/2026 01:13 | 36/33 | 916,80 |
| 2002 | Int 2 | LDR-CIE | 24/03/2026 00:50 | 25/03/2026 20:35 | 29/29 | 648,00 ⚠️ descarrilamiento |
| 2003 | Int 3 | LDR-CIE | 11/04/2026 20:15 | 13/04/2026 01:15 | 33/23 | 611,11 |
| 2004 | Int 4 | LDR-CIE | 26/04/2026 00:55 | 27/04/2026 14:55 | 33/31 | 685,70 |
| 2005 | Int 5 | LDR-CIE | 09/05/2026 22:55 | 10/05/2026 21:25 | 33/30 | 601,11 |

**Pendiente agregar:** Int 6 LDR-CIE (retorno del Int 6 CIE-LDR que salió el 14-mayo) y posiblemente Int 7 CIE-LDR.

### Fix aplicado (mayo 2026)
**Problema:** Al actualizar DATOS en el HTML, los datos nuevos no aparecían porque localStorage ya tenía datos y el `if(trips.length===0)` los ignoraba.

**Solución:** En `window.onload`, ahora se hace un **merge**: se detectan los IDs nuevos en DATOS que no están en localStorage y se agregan automáticamente. Los registros manuales se preservan.

```javascript
const existingIds = new Set(trips.map(t => t.id));
const nuevos = DATOS.filter(d => !existingIds.has(d.id));
if (trips.length === 0) { trips = DATOS; save(); }
else if (nuevos.length > 0) {
  trips = [...trips, ...nuevos].sort((a,b) => a.fecha.localeCompare(b.fecha));
  save();
}
```

### KPIs calculados
- Backhaul (%), Utilización LDR-CIE, Factor carga CIE-LDR/LDR-CIE
- Ciclo completo promedio (días entre salidas CIE-LDR consecutivas) = N-1 intervalos para N viajes
- Desglose del ciclo: Tránsito CIE→LDR + Permanencia La Dorada + Tránsito LDR→CIE + Permanencia Ciénaga
- Permanencias: calculadas por parejas de viajes ida/vuelta por nombre de tren
- Trenes/mes, descarrilamientos, intermodales

### Cómo agregar un nuevo viaje al DATOS
Agregar al array `DATOS` en el HTML con el siguiente formato:
```javascript
{id:1007, tipo:'CIE-LDR', fecha:'2026-XX-XXT HH:MM', fechaLlegada:'2026-XX-XXT HH:MM',
 tren:'Int 7', ptotal:XX, pcarga:XX, pvacias:XX, tons:XXXX.XX, factor:XX.XX,
 desc:0, notas:'', inter:1,
 clientes:[{nombre:'GNL (Papel)',toneladas:XXX,plataformas:XX},{...}]}
// LDR-CIE usa id:2006
```
Después de agregarlo al HTML y recargar la página, el merge lo detecta y lo carga automáticamente.

---

## 📌 Notas para sesión nueva

1. **App ferrocentral** → siempre trabajar sobre `C:\Users\jony9\Downloads\ferrocentral-control.html`
2. **App tren-carga** → siempre trabajar sobre `C:\Users\jony9\Downloads\Jony 11-25\tren-carga.html`
3. El servidor Python en puerto 8080 sirve `C:\Users\jony9\Downloads` — para ver la app en el celular usar `http://10.193.94.95:8080/ferrocentral-control.html`
4. Leer el archivo antes de editar (requisito del Edit tool)
5. Usar PowerShell `[System.IO.File]::ReadAllText/WriteAllText` para reemplazos complejos con backticks/template literals
6. Los registros DEMO en ferrocentral tienen `tipo: undefined` → siempre usar `(r.tipo||'INGRESO')`
7. Formateo colombiano: `fmt()` usa separador de miles con punto, decimales con coma (es-CO)
