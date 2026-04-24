# Sistema de Sorteos — Movistar Arena RRHH

## Resumen
App web para gestionar sorteos de entradas de shows para empleados de Movistar Arena.
Desarrollada en HTML/JS puro (sin frameworks), conectada a Google Apps Script y Google Sheets.

---

## Archivos del proyecto

| Archivo | Descripción |
|---------|-------------|
| `index.html` | App principal de sorteo (dashboard + wizard) |
| `formulario_inscripcion.html` | Formulario público para que empleados se anoten |
| `sorteo_script.gs` | Google Apps Script — backend en Google Sheets |

---

## URLs

- **App**: `https://mateopiccardi-oss.github.io/sorteo-movistar-arena/`
- **Formulario**: `https://mateopiccardi-oss.github.io/sorteo-movistar-arena/formulario_inscripcion.html`
- **Apps Script**: `https://script.google.com/a/macros/movistararena.com.ar/s/AKfycby_KOpMJOKKZSr22ig0j4w-CWVu88z5ghSHzBQqxqOSEQAGj238dhsaSKQr6Pv_tddDaQ/exec`
- **Google Sheet**: `https://docs.google.com/spreadsheets/d/15YojsGUsDfhFBLOkBGxkLXvC7kSr6HUsJQgf99RhpWQ/edit`

---

## Arquitectura

```
Google Sheet (ID: 15YojsGUsDfhFBLOkBGxkLXvC7kSr6HUsJQgf99RhpWQ)
├── Hoja 1              ← empleados (Nombre Completo + Mail) ~284 empleados
├── Inscripciones       ← registros de inscripción al sorteo
├── Ganadores           ← ganadores confirmados (formato tabla)
└── Tracking Ganadores  ← tracking histórico (matriz: col=show, fila=ganador)
                          Fila 1 = fecha (DD/MM/YYYY)
                          Fila 2 = nombre del show (MAYUSCULAS)
                          Fila 3+ = ganadores (MAYUSCULAS)
                          Col A = colaboradores (fórmula, NO TOCAR)
                          Col B = conteo victorias (fórmula, NO TOCAR)
                          Col C en adelante = shows (más reciente en C)
                          Celda B123 = total tickets históricos (992)
```

---

## Estado actual del sistema

### Funcionalidades operativas
- ✅ Formulario público de inscripción con validación de mail corporativo
- ✅ Sincronización de inscriptos desde Sheets
- ✅ Sorteo animado con bombo
- ✅ Gestión de ganadores (reemplazar, agregar, confirmar)
- ✅ Envío de mails via Gmail con PDFs adjuntos desde Drive
- ✅ Import automático al cargar desde Tracking Ganadores
- ✅ Escritura automática en Tracking Ganadores al confirmar ganadores
- ✅ Dashboard con métricas: tickets entregados, top colaboradores, tasa de participación
- ✅ Eventos agrupados por mes con pestañas
- ✅ Historial con paginación (50/pág), filtros y exportación CSV
- ✅ Ranking y frecuencia de victorias (desplegables)
- ✅ Sincronización de ganadores con Sheets (backup)

### Problemas pendientes
- ~~⚠️ Las fechas de eventos a veces se muestran mal (viene del Tracking como Date object)~~ ✅ Resuelto
- ~~⚠️ El donut de tickets suma incorrectamente histórico + nuevos~~ ✅ Resuelto
- ~~⚠️ La app tarda ~1 segundo en cargar porque reimporta siempre del Tracking~~ ✅ Resuelto

---

## Decisiones técnicas importantes

### Storage
- **localStorage**: eventos + ganadores + sorteo activo + ticketsBase (persiste)
- **sessionStorage**: participantes por evento (sesión, hasta 110/día)
- **Tracking Ganadores**: fuente de verdad — siempre se importa al abrir la app

### Al abrir la app
1. Carga localStorage (eventos y ganadores locales)
2. Limpia todos los eventos y ganadores
3. 1 segundo después llama a `leerTracking()` y reimporta todo

### Tickets
- `ticketsBase` = 992 (viene de celda B123 del Tracking)
- Solo cuenta tickets de ganadores nuevos (sorteados desde la app, `_historico:false`)
- Ganadores importados del Tracking tienen `_historico:true`

### Anti-repetición
- Se verifica en **todas las fechas del mismo show** (no solo el evento actual)
- Se puede desactivar por evento

### IDs de ganadores
- Formato: `"g"+Date.now()+"_"+Math.floor(Math.random()*99999)` (sin puntos para evitar bugs en querySelector)

### Tracking Ganadores — escritura
- Siempre inserta columna nueva en **posición C** (empuja el resto a la derecha)
- Si el show ya existe, agrega ganadores abajo
- Fila 1 = fecha DD/MM/YYYY, Fila 2 = NOMBRE SHOW, Fila 3+ = GANADORES EN MAYUSCULAS
- NUNCA toca columnas A y B

---

## Apps Script — endpoints disponibles

| Action | Función | Descripción |
|--------|---------|-------------|
| `validarMail` | `validarMail(mail)` | Valida mail contra Hoja 1 |
| `inscribir` | `inscribir(mail, showId, showNombre, fecha)` | Registra inscripción |
| `getInscriptos` | `getInscriptos(showId)` | Devuelve inscriptos de un show |
| `guardarGanadores` | `guardarGanadores(ganadores, showId, showNombre, fecha, venue)` | Guarda en pestaña Ganadores |
| `enviarMails` | `enviarMails(showId, entradasXGan)` | Envía mails con PDFs desde Drive |
| `syncGanadores` | `syncGanadores(ganadores)` | Backup completo en pestaña Ganadores |
| `trackingGanadores` | `trackingGanadores(ganadores, showNombre, fecha)` | Escribe en Tracking Ganadores |
| `leerTracking` | `leerTracking()` | Lee Tracking y devuelve columnas + ticketsBase + totalEmpleados |
| `getShows` | `getShows()` | Lista shows disponibles |

---

## Drive — estructura de PDFs

```
📁 Sorteo Movistar Arena (en Mi unidad)
   📁 Entradas
      📁 [Nombre del show]   ← debe coincidir con ev.show en la app
         📄 entrada_01.pdf
         📄 entrada_02.pdf
```

---

## Estado del objeto S (estado global de la app)

```javascript
const S = {
  eventos: [],          // Array de eventos creados
  participantes: [],    // Session-only, se limpia al cerrar
  ganadores: [],        // Ganadores confirmados
  sorteo: {             // Sorteo en curso
    evId: null,
    lista: []
  },
  pdfs: [],             // PDFs cargados localmente (legacy)
  csv: { evId: null, rows: [] },
  evActivo: null,       // ID del evento activo en el wizard
  ticketsBase: 0,       // 992 tickets históricos (de B123)
  totalEmpleados: 0     // Total empleados de Hoja 1
}
```

---

## Estructura de un evento

```javascript
{
  id: "NOMBRE_SHOW",           // = ev.show para coincidir con carpeta Drive
  show: "NOMBRE_SHOW",         // Artista/show base
  nombre: "NOMBRE - DD/MM/YYYY", // Nombre completo del evento
  fecha: "YYYY-MM-DD",         // Fecha del show
  hora: "21:00",
  venue: "Movistar Arena, Buenos Aires",
  cantidad: 2,                 // Ganadores a sortear
  entradasXGan: 1,             // Entradas por ganador
  antiRep: true,               // Anti-repetición por show
  formUrl: "",                 // URL del Google Form
  creadoEn: "ISO string",
  _fromTracking: true          // true si fue creado por importación
}
```

---

## Estética

- Dark theme: `--bg:#0F1520`, `--bg2:#162030`, `--bg3:#1A2840`
- Acento cyan: `--ice:#00D4FF`, `--ice2:#60E8FF`
- Fuentes: Barlow Condensed (títulos) + Barlow (cuerpo)
- Logo Movistar Arena: PNG con fondo transparente (procesado)

---

## Flujo de trabajo típico

1. Empleados se anotan via formulario → van a pestaña Inscripciones
2. RRHH abre la app → sincroniza inscriptos (↻ Sincronizar)
3. Sortea → confirma ganadores
4. App escribe ganadores en Tracking Ganadores (col C, empuja resto)
5. Envía mails con entradas PDF adjuntas desde Drive
6. La próxima vez que abre la app → reimporta todo del Tracking automáticamente
