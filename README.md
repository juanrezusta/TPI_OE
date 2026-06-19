# 🚚 Sistema de Rendición de Cobros Virtuales
### Alejandro Negro S.R.L. — Logística y Distribución de Alimentos Congelados

Bot de Telegram + Panel de Control (Tkinter) para automatizar la rendición y conciliación de cobros virtuales de choferes de reparto.

> Desarrollado como Trabajo Práctico Integrador — Organización Empresarial — TUP Comisión 22  
> Autor: García Rezusta, Juan Manuel

---

## ¿Qué hace este sistema?

Los choferes envían mensajes al bot de Telegram indicando los cobros recibidos. El sistema:

1. **Parsea** el mensaje extrayendo monto y medio de pago
2. **Graba** en CSV local (respaldo) y Google Sheets (nube) simultáneamente
3. **Permite cancelar** un cobro respondiendo al mensaje original con "cancelar"
4. **Concilia** automáticamente los registros contra los reportes oficiales del banco y Mercado Pago
5. **Muestra** en tiempo real los últimos movimientos en el panel de control

---

## Estructura del proyecto

```
.
├── app.py                   # Archivo principal — todo en uno
├── credenciales.json        # Google Service Account — NO subir al repo
├── respaldo_YYYY-MM-DD.csv  # Generado automáticamente por fecha
└── last_log.txt             # Últimos 5 movimientos — generado automáticamente
```

---

## Requisitos

- Python 3.11+
- Bot de Telegram creado con [@BotFather](https://t.me/botfather)
- Google Cloud Service Account con acceso a Sheets API y Drive API
- Google Sheet creado y compartido con el email de la Service Account

### Instalación de dependencias

```bash
pip install python-telegram-bot customtkinter gspread google-auth pandas openpyxl
```

---

## Configuración

### 1. Credenciales de Google

1. Ir a [console.cloud.google.com](https://console.cloud.google.com)
2. Habilitar **Google Sheets API** y **Google Drive API**
3. Crear Service Account → descargar JSON → renombrarlo `credenciales.json`
4. Compartir el Google Sheet con el email de la Service Account (rol: Editor)

### 2. Token del bot

En `app.py`, reemplazar:

```python
TOKEN = "tu_token_aqui"
```

### 3. Nombre del Google Sheet

En `app.py`, asegurarse que coincida con el nombre exacto del Sheet en Drive:

```python
NOMBRE_SISTEMA = "Nombre de tu Sheet"
```

---

## Uso

### Iniciar el sistema

```bash
python app.py
```

Se abre el panel de control. Presionar **ENCENDER BOT** para comenzar a recibir mensajes.

---

## Comandos del bot — para choferes

El parser detecta el monto con expresión regular y clasifica el medio de pago por palabras clave.  
**Todo lo que no contenga una palabra de banco se asume Mercado Pago por defecto.**

| Formato | Ejemplos aceptados | Resultado |
|---|---|---|
| `mp [monto]` | `mp 350000` | Mercado Pago |
| `[monto] mp` | `350000 mp` | Mercado Pago |
| `mercado [monto]` | `mercado 80000` | Mercado Pago |
| `[monto]` solo | `350000` | Mercado Pago (default) |
| `[monto] galicia/bna/banco/bco/bnc/santander/bbva/macro/itau/brubank` | `50000 galicia` | Transferencia bancaria |
| Responder al mensaje + `cancelar` / `borrar` / `anular` | — | Elimina ese registro |

---

## Ejecutar conciliación

1. Descargar el reporte de Mercado Pago en `.xlsx` (portal de MP)
2. Descargar el extracto bancario en `.xlsx`
3. Cargarlos como pestañas en el Google Sheet (nombres deben contener "mp" o "banco")
4. Presionar **⚖️ EJECUTAR CONCILIACIÓN** en el panel

Los registros conciliados se pintan en **verde**. Los no encontrados quedan sin marcar para revisión manual.

---

## Cierre de lote

El cierre de lote es una acción **periódica y discrecional**. No es obligatorio hacerlo diariamente.

El administrativo acumula registros durante el período que considere (una semana, un mes o más) y ejecuta el cierre cuando lo decide. Al cerrar:

- Se descarga un backup completo a Excel local
- Se limpian todas las pestañas del Sheets para el próximo período

---

## Manejo de errores

| Situación | Respuesta |
|---|---|
| Mensaje sin número | Ignorado silenciosamente |
| Número solo sin contexto | Se asume Mercado Pago |
| `cancelar` sin responder a un mensaje | "Tenés que responder al pago que querés cancelar" |
| Pago ya eliminado | "No encontré ese pago. Capaz ya está borrado." |
| Sheets vacío al reiniciar | Sincronización automática desde CSV local |
| Excel fuente sin pestaña correcta | El conciliador informa el error específico |

---

## Estados de un registro (Máquina de estados)

```
[Registrado] → FALSE / sin observación
     ↓
[Conciliado automático] → TRUE / OK  (verde en Sheets)
     ↓ (alternativa)
[Conciliado manual] → TRUE / valor custom  (celeste en Sheets)
     ↓ (o en cualquier momento)
[Cancelado] → eliminado de CSV y Sheets
```

---

## Diagrama BPMN

El proceso completo está modelado en `bpmn_alejandro_negro.svg` (BPMN 2.0).  
Tres lanes: **Chofer · Sistema/Bot · Administrativo**  
Cinco gateways: ¿monto válido? · ¿banco o MP? · ¿registro encontrado? · ¿todo conciliado? · ¿cerrar lote?

---

## Seguridad

> ⚠️ **`credenciales.json` nunca debe subirse al repositorio.**

`.gitignore` recomendado:

```
credenciales.json
respaldo_*.csv
last_log.txt
*.xlsx
```

---

## Licencia

Uso privado. Todos los derechos reservados — Alejandro Negro S.R.L.
