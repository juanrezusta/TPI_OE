import os
import sys
import csv
import re
import time
import threading
import asyncio
import pandas as pd
import gspread
import customtkinter as ctk
from datetime import datetime, timedelta
from tkinter import messagebox
from oauth2client.service_account import ServiceAccountCredentials
from telegram import Update
from telegram.ext import ApplicationBuilder, MessageHandler, CommandHandler, filters, ContextTypes
from dotenv import load_dotenv #USO DOTENV PARA LEVANTAR SIN EXPONER EL TOKEN

# ==========================================
# CONFIGURACIONES GLOBALES
# ==========================================
load_dotenv()

TOKEN = os.getenv("BOT_TOKEN")

COLUMNAS_CSV = ["fecha", "hora", "chofer", "telegram_id", "monto", "destino", "msg_id"]  # estructura del csv de respaldo

# el sheet tiene que tener EXACTAMENTE este nombre en tu drive, sino no lo encuentra
NOMBRE_SISTEMA = "Control_Pagos_Activo"

# conectamos con google usando las credenciales del archivo json
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credenciales.json", scope)
client = gspread.authorize(creds)

BOT_ACTIVO = False
BOT_THREAD_INICIADO = False  # para que no se pueda encender el bot dos veces

# ==========================================
# 1. MÓDULO: GOOGLE SHEETS
# ==========================================

def preparar_sheet_diario(fecha):
    # busca el sheet por nombre, si no existe te avisa con un error claro
    try:
        ss = client.open(NOMBRE_SISTEMA)
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"ERROR CRÍTICO: No se encontró el Google Sheet llamado '{NOMBRE_SISTEMA}'.")
        print("Asegurate de haberlo creado en tu Drive y haberlo compartido con el mail del bot.")
        return None

    try:
        ws = ss.worksheet(fecha)  # busca la pestaña de hoy (formato YYYY-MM-DD)
        return ws
    except gspread.exceptions.WorksheetNotFound:
        # si no existe la pestaña del día, la crea con los headers y las reglas de colores
        ws = ss.add_worksheet(title=fecha, rows="1000", cols="10")
        headers = ["FECHA", "HORA", "CHOFER", "ID TEL", "S/D", "DESTINO", "MONTO", "CONCILIADO", ".", "MSG_ID"]
        ws.append_row(headers)
        
        # regla verde: conciliado por el script (TRUE + OK)
        # regla celeste: conciliado a mano por el admin (TRUE pero sin OK)
        requests = [
            {"addConditionalFormatRule": {"rule": {"ranges": [{"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 1000}], "booleanRule": {"condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": '=AND($H2=TRUE; $I2="OK")'}]}, "format": {"backgroundColor": {"red": 0.8, "green": 1.0, "blue": 0.8}}}}, "index": 0}},
            {"addConditionalFormatRule": {"rule": {"ranges": [{"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 1000}], "booleanRule": {"condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": '=AND($H2=TRUE; $I2<>"OK")'}]}, "format": {"backgroundColor": {"red": 0.8, "green": 0.9, "blue": 1.0}}}}, "index": 1}},
            {"updateSheetProperties": {"properties": {"sheetId": ws.id, "gridProperties": {"frozenRowCount": 1}}, "fields": "gridProperties.frozenRowCount"}}  # congela la primera fila
        ]
        ss.batch_update({"requests": requests})
        return ws
    except Exception as e:
        print(f"Error Helper: {e}")
        return None

# ==========================================
# 2. MÓDULO: LÓGICA DEL BOT
# ==========================================

def registrar_log(mensaje):
    # guarda los últimos 5 movimientos en un txt que lee el panel cada 1 segundo
    try:
        lineas = []
        if os.path.exists("last_log.txt"):
            with open("last_log.txt", "r", encoding="utf-8") as f:
                lineas = [l.strip() for l in f.readlines() if l.strip()]
        lineas.append(mensaje.strip())
        with open("last_log.txt", "w", encoding="utf-8") as f:
            f.write("\n".join(lineas[-5:]))  # solo los últimos 5, el resto se va
    except Exception as e:
        pass

def limpiar_entrada(texto):
    # el corazón del parser: extrae el monto y decide si es mp o banco
    texto_min = texto.lower()
    
    # si tiene alguna de estas palabras, es banco. sino, mp por defecto
    palabras_banco = ["banco", "bco", "bnc", "bc", "galicia", "bna", "santander", "bbva", "macro", "itau", "brubank"]
    destino = "banco" if any(p in texto_min for p in palabras_banco) else "mp"
    
    # saca todo lo que no sea número, punto o coma
    monto_sucio = re.sub(r'[^\d.,]', '', texto)
    
    # maneja los distintos formatos: 350.000,00 / 350,000.00 / 350000
    if "," in monto_sucio and "." in monto_sucio:
        if monto_sucio.rfind(",") > monto_sucio.rfind("."):
            monto_sucio = monto_sucio.replace(".", "").replace(",", ".")  # formato argentino
        else:
            monto_sucio = monto_sucio.replace(",", "")  # formato inglés
    elif "," in monto_sucio:
        monto_sucio = monto_sucio.replace(",", ".")
    
    try:
        return float(monto_sucio), destino
    except:
        return None, destino  # si no pudo parsear el monto, devuelve None y el bot ignora el mensaje

def sincronizar_desde_csv():
    # si la pestaña del día está vacía (por ejemplo después de reiniciar), levanta todo desde el csv local
    ahora = datetime.now() - timedelta(hours=3)  # ajuste a horario argentina
    fecha_h = ahora.strftime("%Y-%m-%d")
    archivo_csv = f"respaldo_{fecha_h}.csv"
    ws = preparar_sheet_diario(fecha_h)
    
    if not ws: return
    col_c = ws.col_values(3)
    if len([x for x in col_c if x.strip()]) > 1: return  # si ya tiene datos, no hace nada

    # fórmulas de totales generales para el pie del sheet
    f_gen_con = '=SUMIFS(INDIRECT("G2:G"&ROW()-1), INDIRECT("H2:H"&ROW()-1), TRUE, INDIRECT("C2:C"&ROW()-1), "<>TOTAL*")'
    f_gen_tot = '=SUMIF(INDIRECT("C2:C"&ROW()-1), "<>TOTAL*", INDIRECT("G2:G"&ROW()-1))'

    if not os.path.exists(archivo_csv) or os.path.getsize(archivo_csv) == 0:
        # si tampoco hay csv, crea solo la fila de total general vacía
        ws.append_row(["", "", "TOTAL GENERAL", "Conciliado:", f_gen_con, "Total:", f_gen_tot, "", "", ""], value_input_option='USER_ENTERED')
        ws.format("A2:J2", {"backgroundColor": {"red": 1, "green": 0.9, "blue": 0.6}, "textFormat": {"bold": True}})
        return

    df = pd.read_csv(archivo_csv, names=COLUMNAS_CSV, header=0, encoding='utf-8')
    if df.empty: return

    filas = []
    for chofer_n, grupo in df.groupby('chofer'):
        chofer_n = str(chofer_n).upper()
        for _, r in grupo.iterrows():
            monto_float = float(r['monto'])
            filas.append([r['fecha'], r['hora'], chofer_n, r['telegram_id'], "S/D", r['destino'], monto_float, False, "", f"'{r['msg_id']}"])
        
        # agrega una fila de subtotal por chofer con fórmulas dinámicas
        f_con = f'=SUMIFS(INDIRECT("G2:G"&ROW()-1), INDIRECT("C2:C"&ROW()-1), "{chofer_n}", INDIRECT("H2:H"&ROW()-1), TRUE)'
        f_tot = f'=SUMIF(INDIRECT("C2:C"&ROW()-1), "{chofer_n}", INDIRECT("G2:G"&ROW()-1))'
        filas.append(["", "", f"TOTAL {chofer_n}", "Conciliado:", f_con, "Total:", f_tot, "---", "---", "---"])

    filas.append(["", "", "TOTAL GENERAL", "Conciliado:", f_gen_con, "Total:", f_gen_tot, "", "", ""])
    ws.update(range_name=f"A2:J{len(filas)+1}", values=filas, value_input_option='USER_ENTERED')
    
    # pinta las filas de total: amarillo para general, gris para los individuales
    for i, f in enumerate(filas):
        if "TOTAL" in str(f[2]):
            color = {"red": 1, "green": 0.9, "blue": 0.6} if "GENERAL" in f[2] else {"red": 0.9, "green": 0.9, "blue": 0.9}
            ws.format(f"A{i+2}:J{i+2}", {"backgroundColor": color, "textFormat": {"bold": True}})

async def recibir_pago(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # se dispara cada vez que llega un mensaje al bot
    try:
        ahora = update.message.date - timedelta(hours=3)  # telegram manda en UTC, lo pasamos a argentina
        fecha_h, hora_h = ahora.strftime("%Y-%m-%d"), ahora.strftime("%H:%M:%S")
        chofer = update.message.from_user.first_name.upper()
        monto, destino = limpiar_entrada(update.message.text)
        if monto is None: return  # mensaje inválido, lo ignoramos y no respondemos nada

        # graba en el csv local primero (respaldo por si falla la nube)
        archivo_csv = f"respaldo_{fecha_h}.csv"
        file_exists = os.path.exists(archivo_csv)
        with open(archivo_csv, mode='a', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            if not file_exists: writer.writerow(COLUMNAS_CSV)  # si es el primero del día, escribe el header
            writer.writerow([fecha_h, hora_h, chofer, update.message.from_user.id, monto, destino, update.message.message_id])

        # ahora va a sheets
        ws = preparar_sheet_diario(fecha_h)
        col_c = ws.col_values(3)
        if len([x for x in col_c if x.strip()]) <= 1:
            sincronizar_desde_csv()  # pestaña vacía, sincroniza desde csv antes de insertar
            data = ws.get_all_values()
        else:
            data = ws.get_all_values()

        # busca dónde insertar la fila: antes del total del chofer, o antes del total general
        idx_chofer = next((i+1 for i, f in enumerate(data) if len(f) > 2 and f"TOTAL {chofer}" in str(f[2])), -1)
        idx_gen = next((i+1 for i, f in enumerate(data) if len(f) > 2 and "TOTAL GENERAL" in str(f[2])), -1)
        fila_dato = [fecha_h, hora_h, chofer, update.message.from_user.id, "S/D", destino, float(monto), False, "", f"'{update.message.message_id}"]

        if idx_chofer != -1:
            # el chofer ya tiene filas, inserta antes de su total
            ws.insert_row(fila_dato, index=idx_chofer, value_input_option='USER_ENTERED')
            ws.format(f"A{idx_chofer}:J{idx_chofer}", {"backgroundColor": {"red": 1, "green": 1, "blue": 1}, "textFormat": {"bold": False}})
        elif idx_gen != -1:
            # primer pago del chofer en el día, crea su sección con total
            ws.insert_row(fila_dato, index=idx_gen, value_input_option='USER_ENTERED')
            ws.format(f"A{idx_gen}:J{idx_gen}", {"backgroundColor": {"red": 1, "green": 1, "blue": 1}, "textFormat": {"bold": False}})
            f_con = f'=SUMIFS(INDIRECT("G2:G"&ROW()-1), INDIRECT("C2:C"&ROW()-1), "{chofer}", INDIRECT("H2:H"&ROW()-1), TRUE)'
            f_tot = f'=SUMIF(INDIRECT("C2:C"&ROW()-1), "{chofer}", INDIRECT("G2:G"&ROW()-1))'
            ws.insert_row(["", "", f"TOTAL {chofer}", "Conciliado:", f_con, "Total:", f_tot, "---", "---", "---"], index=idx_gen+1, value_input_option='USER_ENTERED')
            ws.format(f"A{idx_gen+1}:J{idx_gen+1}", {"backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9}, "textFormat": {"bold": True}})

        registrar_log(f"✅ {hora_h} | {chofer} | ${monto} ({destino.upper()})")
        await update.message.reply_text(f"✅ Anotado: ${monto} ({destino.upper()})")
    except Exception as e: pass

async def cancelar_pago(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # el chofer tiene que RESPONDER al mensaje que quiere cancelar y escribir "cancelar"
    try:
        if not update.message.reply_to_message:
            await update.message.reply_text("⚠️ Tenes que responder al pago que queres cancelar o borrar.")
            return
        
        msg_id_a_borrar = str(update.message.reply_to_message.message_id)  # el id del mensaje original
        ahora = update.message.date - timedelta(hours=3)
        fecha_h = ahora.strftime("%Y-%m-%d")

        ws = preparar_sheet_diario(fecha_h)
        if not ws: return
        data = ws.get_all_values()

        fila_a_borrar = -1
        chofer_borrado = "Chofer"
        monto_borrado = "0"
        destino_borrado = "S/D"

        # busca la fila que tenga ese msg_id en la columna J
        for i, fila in enumerate(data):
            if len(fila) >= 10:
                msg_sheet = str(fila[9]).replace("'", "").strip()
                if msg_sheet == msg_id_a_borrar:
                    fila_a_borrar = i + 1
                    chofer_borrado = str(fila[2])
                    monto_borrado = str(fila[6])
                    destino_borrado = str(fila[5]).upper()
                    break

        if fila_a_borrar != -1:
            ws.delete_rows(fila_a_borrar)  # borra del sheet
            
            # también lo borra del csv local
            archivo_csv = f"respaldo_{fecha_h}.csv"
            if os.path.exists(archivo_csv):
                with open(archivo_csv, 'r', encoding='utf-8') as f:
                    filas_csv = list(csv.reader(f))
                with open(archivo_csv, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    for row in filas_csv:
                        if len(row) > 6 and str(row[6]) != msg_id_a_borrar:
                            writer.writerow(row)  # escribe todo menos la fila a borrar
            
            registrar_log(f"❌ {ahora.strftime('%H:%M:%S')} | {chofer_borrado} | -${monto_borrado} ({destino_borrado}) ANULADO")
            await update.message.reply_text("🗑️ Pago anulado y borrado del sistema.")
        else:
            await update.message.reply_text("⚠️ No encontre ese pago. Capaz ya esta borrado.")
    except Exception as e: pass

def arrancar_hilo_bot():
    # corre el bot en un hilo separado para que no trabe el panel de tkinter
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    sincronizar_desde_csv()  # al arrancar, chequea si hay que sincronizar desde csv
    app = ApplicationBuilder().token(TOKEN).build()
    
    # primero los comandos de cancelación, después cualquier texto
    app.add_handler(CommandHandler(["cancelar", "borrar", "anular"], cancelar_pago))
    regex_borrar = re.compile(r'(?i)^(cancelar|anular|borrar|anulada|borrada)$')
    app.add_handler(MessageHandler(filters.Regex(regex_borrar), cancelar_pago))
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), recibir_pago))
    app.run_polling()

# ==========================================
# 3. MÓDULO: CONCILIADOR
# ==========================================

def encontrar_archivo_del_dia(fecha_str):
    # busca en la carpeta actual un excel que tenga la fecha de hoy en el nombre
    for archivo in os.listdir('.'):
        if fecha_str in archivo and (archivo.endswith('.xlsx') or archivo.endswith('.xls')):
            return archivo
    return None

def limpiar_monto_estricto(valor):
    # normaliza cualquier formato de número a float, devuelve None si no puede
    if pd.isna(valor) or str(valor).strip() == "": return None
    if isinstance(valor, (int, float)): return round(float(valor), 2)
    valor_str = str(valor).strip()
    if "." in valor_str and "," in valor_str:
        if valor_str.rfind(",") > valor_str.rfind("."): valor_str = valor_str.replace(".", "").replace(",", ".")
        else: valor_str = valor_str.replace(",", "")
    elif "," in valor_str: valor_str = valor_str.replace(",", ".")
    monto_sucio = re.sub(r'[^\d.]', '', valor_str)
    try: return round(float(monto_sucio), 2)
    except: return None

def extraer_montos_mp(archivo):
    # levanta el excel de mp y devuelve la lista de montos cobrados
    try:
        xl = pd.ExcelFile(archivo)
        # busca la pestaña que tenga "mp" o "mercado" en el nombre
        nombre_pestaña = next((s for s in xl.sheet_names if 'mp' in s.lower() or 'mercado' in s.lower()), None)
        if not nombre_pestaña: return None, "Error al levantar el archivo (MP): Pestaña no encontrada."
        
        df = pd.read_excel(archivo, sheet_name=nombre_pestaña)
        col_importe, col_tipo = None, None
        for col_original in df.columns:
            col_lower = str(col_original).lower()
            if "importe" in col_lower or "monto" in col_lower: col_importe = col_original
            if "tipo" in col_lower and "operac" in col_lower: col_tipo = col_original
                
        if not col_importe: return None, "Error al levantar el archivo (MP): Columna de importe no encontrada."
        
        # filtra solo los cobros recibidos, ignora comisiones y otros movimientos
        if col_tipo:
            mask = df[col_tipo].astype(str).str.lower().str.contains('recibido|ingreso|cobro', na=False)
            df = df[mask]
            
        montos = df[col_importe].apply(limpiar_monto_estricto).dropna()
        lista_montos = [round(float(m), 2) for m in montos if float(m) > 0]
        return lista_montos, "OK"
    except Exception as e: return None, "Error al levantar el archivo (MP). Revisa si está cerrado."

def extraer_montos_banco(archivo):
    # levanta el excel del banco y devuelve los créditos (transferencias recibidas)
    try:
        xl = pd.ExcelFile(archivo)
        nombre_pestaña = next((s for s in xl.sheet_names if 'banco' in s.lower()), None)
        if not nombre_pestaña: return None, "Error al levantar el archivo (Banco): Pestaña no encontrada."
        
        df = pd.read_excel(archivo, sheet_name=nombre_pestaña)
        col_credito = None
        for col_original in df.columns:
            col_lower = str(col_original).lower()
            if "ditos" in col_lower or "cred" in col_lower:
                col_credito = col_original
                break
                
        if col_credito:
            montos = df[col_credito].apply(limpiar_monto_estricto).dropna()
            lista_montos = [round(float(m), 2) for m in montos if float(m) > 0]
            return lista_montos, "OK"
        else: return None, "Error al levantar el archivo (Banco): Columna de Créditos no encontrada."
    except Exception as e: return None, "Error al levantar el archivo (Banco). Revisa si está cerrado."

def buscar_y_remover_monto(monto_buscado, lista_montos, tolerancia=2.0):
    # busca el monto en la lista con una tolerancia de $2 por centavos
    # cuando lo encuentra lo saca de la lista para no usarlo dos veces
    for i, monto_excel in enumerate(lista_montos):
        if abs(monto_excel - monto_buscado) <= tolerancia:
            lista_montos.pop(i)
            return True
    return False

def ejecutar_conciliacion_str():
    # el motor de la conciliación: cruza los registros del sheet con los excels del banco/mp
    reporte = []
    ahora = datetime.now() - timedelta(hours=3)
    fecha_h = ahora.strftime("%Y-%m-%d")

    archivo_excel = encontrar_archivo_del_dia(fecha_h)
    if not archivo_excel: return f"ERROR: No se encontró el Excel del día ({fecha_h})."

    try:
        ss = client.open(NOMBRE_SISTEMA)
        ws = ss.worksheet(fecha_h)
    except Exception as e:
        return f"ERROR: No se encontró la pestaña de hoy en Google Sheets.\nDetalle técnico: {e}"

    montos_mp, estado_mp = extraer_montos_mp(archivo_excel)
    if estado_mp != "OK": montos_mp = []
    
    montos_banco, estado_banco = extraer_montos_banco(archivo_excel)
    if estado_banco != "OK": montos_banco = []

    data = ws.get_all_values()
    celdas_a_actualizar = []
    pagos_conciliados = 0

    # pasada 1: descarta los montos que ya estaban conciliados para no usarlos dos veces
    for i, fila in enumerate(data):
        fila_sheet = i + 1
        if fila_sheet == 1 or len(fila) < 8 or "TOTAL" in str(fila[2]).upper(): continue
        destino = str(fila[5]).strip().lower()
        estado_actual = str(fila[7]).strip().upper()
        
        if estado_actual in ["TRUE", "VERDADERO"]:
            monto_sheet = limpiar_monto_estricto(fila[6])
            if monto_sheet is not None:
                if destino == "mp": buscar_y_remover_monto(monto_sheet, montos_mp)
                elif destino == "banco": buscar_y_remover_monto(monto_sheet, montos_banco)

    # pasada 2: concilia los que están en FALSE
    for i, fila in enumerate(data):
        fila_sheet = i + 1
        if fila_sheet == 1 or len(fila) < 8 or "TOTAL" in str(fila[2]).upper(): continue
        destino = str(fila[5]).strip().lower()
        estado_actual = str(fila[7]).strip().upper()
        
        if estado_actual not in ["FALSE", "FALSO"]: continue
        monto_raw = fila[6]
        monto_sheet = limpiar_monto_estricto(monto_raw)
        if monto_sheet is None: continue
            
        coincidencia = False
        if destino == "mp" and buscar_y_remover_monto(monto_sheet, montos_mp): coincidencia = True
        elif destino == "banco" and buscar_y_remover_monto(monto_sheet, montos_banco): coincidencia = True

        if coincidencia:
            # marca la fila como TRUE y OK en el sheet (se va a pintar verde automáticamente)
            celdas_a_actualizar.append(gspread.Cell(row=fila_sheet, col=8, value=True))
            celdas_a_actualizar.append(gspread.Cell(row=fila_sheet, col=9, value="OK"))
            pagos_conciliados += 1

    if celdas_a_actualizar:
        ws.update_cells(celdas_a_actualizar, value_input_option='USER_ENTERED')  # actualiza todo de una sola vez

    reporte.append("      REPORTE DE CONCILIACIÓN        ")
    reporte.append("=====================================")
    reporte.append(f"  - Archivo Mercado Pago : {estado_mp}")
    reporte.append(f"  - Archivo Banco        : {estado_banco}")
    reporte.append("-------------------------------------")
    reporte.append(f"  Conciliación Exitosa: {pagos_conciliados} importes chequeados.")
    
    return "\n".join(reporte)

def descargar_respaldo_nube():
    # descarga todo el sheet a excel local y limpia la nube para el próximo lote
    # OJO: esto borra todas las pestañas de datos, deeja solo "Instrucciones"
    try:
        ss = client.open(NOMBRE_SISTEMA)
        worksheets = ss.worksheets()
        
        ahora = datetime.now() - timedelta(hours=3)
        fecha_str = ahora.strftime("%Y-%m-%d_%H-%M")
        nombre_archivo = f"Historial_{NOMBRE_SISTEMA}_{fecha_str}.xlsx"
        
        # guarda cada pestaña como hoja del excel local
        with pd.ExcelWriter(nombre_archivo, engine='openpyxl') as writer:
            for ws in worksheets:
                datos = ws.get_all_values()
                if not datos:
                    df = pd.DataFrame()
                else:
                    df = pd.DataFrame(datos[1:], columns=datos[0])
                
                sheet_title = ws.title[:31]  # excel no acepta nombres de hoja de más de 31 caracteres
                df.to_excel(writer, sheet_name=sheet_title, index=False)
                
        # si no existe la pestaña de instrucciones, la crea con el manual de usuario
        nombres_pestañas = [ws.title.lower() for ws in worksheets]
        if "instrucciones" not in nombres_pestañas:
            ws_inst = ss.add_worksheet(title="Instrucciones", rows="30", cols="5")
            
            manual = [
                ["📘 MANUAL DE USUARIO - SISTEMA DE CONTROL DE COBROS v0.22.1"],
                [""],
                ["📱 1. REGISTRO DE PAGOS (TELEGRAM)"],
                ["Los choferes deben enviar el monto y el destino. No importa el orden ni las mayúsculas."],
                ["- Para Banco: Usar palabras como 'banco', 'galicia', 'bco'."],
                ["- Para Mercado Pago: Usar 'mp' o si solo mandan el número, el bot asume por defecto que es Mercado Pago."],
                ["- Ejemplos válidos: '15000 banco', 'mp 8500.50', 'Banco 10000'."],
                [""],
                ["❌ 2. ANULACIÓN DE PAGOS"],
                ["Si un chofer se equivoca al enviar un cobro, debe RESPONDER al mensaje de confirmación del bot con alguna de estas palabras:"],
                ["- 'cancelar', 'anular' o 'borrar'."],
                ["El bot detectará el mensaje original y eliminará el registro automáticamente de la planilla."],
                [""],
                ["⚖️ 3. CONCILIACIÓN AUTOMÁTICA"],
                ["En el programa de la PC, al tocar 'CONCILIAR HASTA AHORA', el sistema cruzará los datos de Google Sheets con tu Excel del Banco/MP."],
                ["- Los pagos que coincidan exactamente se marcarán en VERDE (TRUE y OK)."],
                ["- Tolerancia: El sistema perdona hasta $2.00 de diferencia por posibles errores de tipeo en los centavos."],
                [""],
                ["✍️ 4. CONCILIACIÓN MANUAL"],
                ["Si un pago no cruzó automáticamente pero vos verificaste que el dinero ingresó, podés validarlo a mano:"],
                ["- En Google Sheets, andá a la columna 'CONCILIADO' (Columna H) de ese pago y escribí la palabra: TRUE"],
                ["- Al hacer esto, la fila se pintará automáticamente de color CELESTE para indicar que fue validada por un humano."],
                [""],
                ["💾 5. CIERRE DE LOTE (DESCARGAR HISTORIAL)"],
                ["Al terminar la jornada o la semana, usá el botón 'DESCARGAR HISTORIAL' en el panel de la PC."],
                ["- Guardará una copia perfecta de todos los pagos en tu computadora en formato Excel."],
                ["- Limpiará todas las pestañas de datos de la nube para mantener el sistema rápido, dejando solo estas instrucciones."]
            ]
            
            ws_inst.update(range_name='A1:A28', values=manual, value_input_option='USER_ENTERED')
            
            # formato visual de la pestaña de instrucciones
            ws_inst.format("A1", {"textFormat": {"bold": True, "fontSize": 12, "foregroundColor": {"red": 1, "green": 1, "blue": 1}}, "backgroundColor": {"red": 0.18, "green": 0.28, "blue": 0.38}})
            formato_sub = {"textFormat": {"bold": True, "fontSize": 11, "foregroundColor": {"red": 0.1, "green": 0.3, "blue": 0.7}}}
            for celda in ["A3", "A9", "A14", "A19", "A24"]:
                ws_inst.format(celda, formato_sub)
            ss.batch_update({"requests": [{"updateDimensionProperties": {"range": {"sheetId": ws_inst.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1}, "properties": {"pixelSize": 850}, "fields": "pixelSize"}}]})
            worksheets = ss.worksheets()

        # borra todo menos instrucciones
        pestañas_borradas = 0
        for ws in worksheets:
            if ws.title.lower() != "instrucciones":
                ss.del_worksheet(ws)
                pestañas_borradas += 1
                
        return True, f"✅ Se guardó un backup de {len(worksheets)} pestañas en tu PC.\n🗑️ Se eliminaron {pestañas_borradas} pestañas de datos de la nube.\n\nEl sistema quedó limpio para el próximo turno."
    
    except gspread.exceptions.SpreadsheetNotFound:
        return False, f"No se encontró el Google Sheet '{NOMBRE_SISTEMA}'."
    except Exception as e:
        return False, f"Error durante el respaldo: {e}"


# ==========================================
# 4. MÓDULO: PANEL DE CONTROL (TKINTER)
# ==========================================

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class PanelControl(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Sistema de Control de Cobros - v0.22.1")
        self.geometry("520x520")
        self.protocol("WM_DELETE_WINDOW", self.cerrar_aplicacion)

        self.label_titulo = ctk.CTkLabel(self, text="PANEL DE CONTROL", font=ctk.CTkFont(size=22, weight="bold"))
        self.label_titulo.pack(pady=20)

        # botón de encender el bot, una vez encendido se bloquea hasta cerrar el programa
        self.btn_bot = ctk.CTkButton(self, text="ENCENDER BOT", command=self.toggle_bot, 
                                     fg_color="#2ecc71", hover_color="#27ae60", height=45,
                                     font=ctk.CTkFont(size=14, weight="bold"))
        self.btn_bot.pack(pady=10, padx=30, fill="x")

        self.frame_info = ctk.CTkFrame(self)
        self.frame_info.pack(pady=15, padx=30, fill="both", expand=True)

        self.lbl_titulo_info = ctk.CTkLabel(self.frame_info, text="ÚLTIMOS MOVIMIENTOS", font=ctk.CTkFont(size=12, weight="bold"))
        self.lbl_titulo_info.pack(pady=5)

        # acá se muestran los últimos movimientos, se actualiza cada 1 segundo
        self.lbl_info = ctk.CTkLabel(self.frame_info, text="Esperando registros...", font=ctk.CTkFont(size=13), text_color="#3498db", justify="left")
        self.lbl_info.pack(pady=10, padx=10, anchor="w")

        self.frame_botones = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_botones.pack(pady=10, padx=30, fill="x")

        self.btn_conc = ctk.CTkButton(self.frame_botones, text="⚖️ CONCILIAR HASTA AHORA", command=self.run_conciliacion,
                                      fg_color="#8e44ad", hover_color="#9b59b6", height=40)
        self.btn_conc.pack(side="left", fill="x", expand=True, padx=(0, 5))

        # botón de cierre de lote, tiene confirmación para que no se apriete sin querer
        self.btn_backup = ctk.CTkButton(self.frame_botones, text="💾 DESCARGAR HISTORIAL", command=self.run_backup,
                                        fg_color="#34495e", hover_color="#2c3e50", height=40)
        self.btn_backup.pack(side="right", fill="x", expand=False, padx=(5, 0))

        threading.Thread(target=self.monitor_logs, daemon=True).start()

    def toggle_bot(self):
        global BOT_THREAD_INICIADO
        if not BOT_THREAD_INICIADO:
            threading.Thread(target=arrancar_hilo_bot, daemon=True).start()
            BOT_THREAD_INICIADO = True
            self.btn_bot.configure(
                text="🤖 BOT ACTIVO (Cerrá el programa para apagar)", 
                fg_color="#043d1c", 
                state="disabled"  # se bloquea para que no lo enciendan dos veces
            )

    def monitor_logs(self):
        # hilo que lee el last_log.txt cada 1 segundo y actualiza el panel
        while True:
            if os.path.exists("last_log.txt"):
                try:
                    with open("last_log.txt", "r", encoding="utf-8") as f:
                        lineas = [line.strip() for line in f.readlines() if line.strip()]
                        contenido = "\n\n".join(lineas)
                        if contenido:
                            self.lbl_info.configure(text=contenido)
                except:
                    pass
            time.sleep(1)

    def run_conciliacion(self):
        # corre la conciliación en un hilo separado para no trabar la interfaz
        self.btn_conc.configure(state="disabled", text="⏳ PROCESANDO...")
        
        def tarea():
            try:
                reporte_final = ejecutar_conciliacion_str()
                if "ERROR" in reporte_final:
                    self.after(0, messagebox.showerror, "Error de Conciliación", reporte_final)
                else:
                    self.after(0, messagebox.showinfo, "Reporte de Conciliación", reporte_final)
            except Exception as e:
                self.after(0, messagebox.showerror, "Error", f"Fallo catastrófico: {e}")
            finally:
                self.after(0, self.btn_conc.configure, state="normal", text="⚖️ CONCILIAR HASTA AHORA")

        threading.Thread(target=tarea, daemon=True).start()

    def run_backup(self):
        # pide confirmación antes de borrar todo, por las dudas
        confirmacion = messagebox.askyesno(
            "Confirmar Cierre de Lote", 
            "¿Estás seguro de que querés descargar el historial y BORRAR todas las pestañas de la nube?\n\n⚠️ Asegurate de que la jornada haya terminado y el Bot esté apagado."
        )
        
        if not confirmacion:
            return  # si toca "No", no hace nada
            
        self.btn_backup.configure(state="disabled", text="⏳ CERRANDO LOTE...")
        
        def tarea():
            try:
                exito, mensaje = descargar_respaldo_nube()
                if exito:
                    self.after(0, messagebox.showinfo, "Respaldo y Limpieza Exitosa", mensaje)
                else:
                    self.after(0, messagebox.showerror, "Error de Respaldo", mensaje)
            except Exception as e:
                self.after(0, messagebox.showerror, "Error Crítico", f"Fallo al descargar: {e}")
            finally:
                self.after(0, self.btn_backup.configure, state="normal", text="💾 DESCARGAR HISTORIAL")

        threading.Thread(target=tarea, daemon=True).start()

    def cerrar_aplicacion(self):
        self.destroy()
        os._exit(0)  # mata todo el proceso limpio, incluyendo los hilos del bot

if __name__ == "__main__":
    app = PanelControl()
    app.mainloop()
