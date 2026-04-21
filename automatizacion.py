import os
import pandas as pd
import gspread
from datetime import datetime
from google.oauth2.service_account import Credentials
from playwright.sync_api import sync_playwright
from dotenv import load_dotenv

# --- CONFIGURACIÓN ---
load_dotenv()
USUARIO_REAL = os.getenv("USUARIO_SCE")
CONTRASENA_REAL = os.getenv("PASSWORD_SCE")
DOWNLOAD_DIR = r"C:\Users\cmancipe\Downloads"     
GOOGLE_SHEET_ID = "1WRaIZByE4nxC9x7NxVgW8QPjjFg0uOdww3bsSQE0GiU"
HOJA_NOMBRE = "Base de datos limpia"
LOOKER_URL = "https://lookerstudio.google.com/u/0/reporting/c820bd4c-9a70-43f2-ba84-3799b1fa213b/page/p_n7i63ucatd"


# --- FUNCIÓN AUXILIAR: Edita la fecha en la página actualmente visible ---
def editar_fecha_visible(page, fecha_texto):
    """
    Busca TODOS los elementos con 'Actualizado a:' en el DOM
    y edita únicamente el que esté VISIBLE en pantalla en este momento.
    """
    elementos = page.get_by_text("Actualizado a:", exact=False).all()
    
    editado = False
    for elemento in elementos:
        if elemento.is_visible():  # ← Solo actúa sobre el visible real
            elemento.dblclick(force=True)
            page.wait_for_timeout(500)
            page.keyboard.press("Control+a")
            page.keyboard.type(fecha_texto)
            page.keyboard.press("Escape")
            editado = True
            break  # Ya encontró el correcto, no seguir buscando
    
    return editado


# --- PASO 0: FUNCIÓN PARA GUARDAR SESIÓN ---
def guardar_sesion_google():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        page.goto("https://accounts.google.com")
        print("👉 POR FAVOR: Inicia sesión en Google.")
        input("🔐 Una vez logueado, presiona ENTER...")
        context.storage_state(path="google_session.json")
        print("✅ Sesión guardada.")
        browser.close()


# --- PASO 1: DESCARGAR EXCEL ---
def descargar_excel():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, downloads_path=DOWNLOAD_DIR)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        
        page.goto("https://sce.cumandes.com/")
        page.fill('input#LoginForm_username', USUARIO_REAL)
        page.fill('input#LoginForm_password.form-control', CONTRASENA_REAL)
        page.click('input#enterokay.btn.btn-success.btn-xs.btn-block')
        
        page.wait_for_url("**/site/index", timeout=60000) 
        page.goto("https://sce.cumandes.com/dashboard/projects/index")
        
        print("⏳ Esperando botón de descarga...")
        with page.expect_download(timeout=0) as dl:
            page.click("a#download-report.btn.blue-bg.text-uppercase.col-4")
         
        download = dl.value
        ruta = os.path.join(DOWNLOAD_DIR, download.suggested_filename)
        download.save_as(ruta)
        browser.close()
        return ruta


# --- PASO 2: LIMPIAR DATOS ---
def procesar_datos(ruta_excel):
    df = pd.read_excel(ruta_excel, sheet_name="Detalle de Cotización")
    df = df.map(lambda x: x.strftime("%Y-%m-%d") if pd.notnull(x) and hasattr(x, "strftime") else x)
    df = df.fillna("")
    return df


# --- PASO 3: SUBIR A GOOGLE SHEETS ---
def actualizar_sheets(df):
    creds = Credentials.from_service_account_file("credenciales.json", 
            scopes=["https://www.googleapis.com/auth/spreadsheets"])
    client = gspread.authorize(creds)
    sheet = client.open_by_key(GOOGLE_SHEET_ID).worksheet(HOJA_NOMBRE)
    sheet.clear()
    datos = df.values.tolist()
    sheet.update(range_name='A1', values=datos)
    print("✅ Google Sheets actualizado.")


# --- PASO 4: ACTUALIZAR FECHA EN LOOKER STUDIO ---
def actualizar_looker():
    if not os.path.exists("google_session.json"):
        print("❌ Error: No existe 'google_session.json'.")
        return

    ahora = datetime.now()
    fecha_texto = f"Actualizado a: {ahora.strftime('%d/%m/%Y %I:%M %p')}"

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(storage_state="google_session.json")
        page = context.new_page()
        
        print("🌐 Entrando a Looker Studio...")
        page.goto(LOOKER_URL, timeout=60000)

        try:
            print("⏳ Esperando el botón de edición...")
            page.click('button.edit-mode-button', timeout=20000)
            print("✏️ Modo edición activado.")
            page.wait_for_timeout(2000)
            
            # --- PÁGINA DE INICIO ---
            print("🏠 Editando fecha en la página de Inicio...")
            if editar_fecha_visible(page, fecha_texto):
                print("  ✅ Página de Inicio actualizada.")
            else:
                print("  ⚠️ No se encontró elemento visible en Inicio.")
            
            # --- CAMBIO A MODO VISTA PARA NAVEGAR ---
            print("⏳ Pasando a modo vista para navegar...")
            page.click('button.view-mode-button', timeout=20000)
            page.wait_for_timeout(2000)
            
            # --- ENTRAR AL INFORME ---
            print("📊 Entrando a INFORME DE PROYECCIÓN...")
            page.click('.button-control-mask')
            page.wait_for_timeout(4000)
            
            # --- VOLVER A EDICIÓN ---
            print("⏳ Reactivando el modo edición...")
            page.click('button.edit-mode-button', timeout=20000)
            page.wait_for_timeout(2000)
            
            # --- BUCLE DE HOJAS ---
            for i in range(1, 10):
                print(f"📄 Procesando hoja {i}...")
                
                try:
                    if editar_fecha_visible(page, fecha_texto):
                        print(f"  ✅ Hoja {i} actualizada.")
                    else:
                        print(f"  ⚠️ Hoja {i}: no hay elemento visible (puede no tener fecha).")
                except Exception as e_h:
                    print(f"  ❌ Error en hoja {i}: {e_h}")

                # Ir a la siguiente página
                if i < 10:
                    btn_sig = page.locator('button[aria-label="Página siguiente"]')
                    if btn_sig.is_visible():
                        btn_sig.click(force=True)
                        page.wait_for_timeout(3000)
                    else:
                        print("  🏁 No hay más páginas.")
                        break  # Salir si ya no hay más páginas

            # Guardar cambios
            page.keyboard.press("Control+s")
            print(f"\n✅ ¡Looker actualizado!: {fecha_texto}")

        except Exception as e:
            print(f"❌ Error durante la edición: {e}")
        
        finally:
            page.wait_for_timeout(2000)
            browser.close()


# --- FLUJO PRINCIPAL ---
if __name__ == "__main__":
    try:
        archivo = descargar_excel()
        datos = procesar_datos(archivo)
        actualizar_sheets(datos)
        actualizar_looker()
        print("🚀 Proceso terminado con éxito.")
    except Exception as e:
        print(f"❌ Error general: {e}")