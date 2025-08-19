from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook, load_workbook
import pandas as pd
import time
import os
import glob
import random
import xlwings as xw
import inspect
import sys
from datetime import datetime

# Obtener la ruta base del directorio donde está el script
base_dir = os.path.dirname(os.path.abspath(__file__))

# MODIFICACIÓN: Nueva ruta de entrada para clientes
input_excel_clientes = os.path.join(base_dir, "Data", "Clientes.xlsx")

# Definir rutas a las carpetas y archivos (mantenidas para compatibilidad)
input_folder_excel = os.path.join(base_dir, "data", "input", "Deudas")
output_folder_csv = os.path.join(base_dir, "data", "input", "DeudasCSV")
output_file_csv = os.path.join(base_dir, "data", "Resumen_deudas.csv")
output_file_xlsx = os.path.join(base_dir, "data", "Resumen_deudas.xlsx")

# Leer el archivo Excel
df = pd.read_excel(input_excel_clientes)

# Suposición de nombres de columnas
cuit_login_list = df['CUIT para ingresar'].tolist()
cuit_represent_list = df['CUIT representado'].tolist()
password_list = df['Contraseña'].tolist()
download_list = df['Ubicacion Descarga'].tolist()
clientes_list = df['Cliente'].tolist()

# Variable global para el driver (se recreará por cada cliente)
driver = None

def configurar_nuevo_navegador():
    """Configura y retorna un nuevo navegador Chrome limpio."""
    global driver
    
    # Configuración de opciones de Chrome
    options = Options()
    options.add_argument("--start-maximized")
    
    # Configurar preferencias de descarga
    prefs = {
        "download.prompt_for_download": True,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    
    # Inicializar driver nuevo
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    
    print("✅ Nuevo navegador Chrome configurado")
    return driver

def cerrar_sesion_y_navegador():
    """Cierra sesión completa y navegador - NUEVA FUNCIÓN MEJORADA."""
    global driver
    
    try:
        print("\n--- INICIANDO CIERRE COMPLETO DE SESIÓN ---")
        
        # PASO 1: Verificar cuántas pestañas están abiertas
        window_handles = driver.window_handles
        num_pestanas = len(window_handles)
        print(f"📊 Pestañas abiertas detectadas: {num_pestanas}")
        
        # PASO 2: Si hay más de 1 pestaña, cerrar las adicionales
        if num_pestanas > 1:
            print(f"🔄 Cerrando {num_pestanas - 1} pestañas adicionales...")
            
            # Ir a la última pestaña (SCT) y cerrarla
            for i in range(num_pestanas - 1, 0, -1):  # Desde la última hacia la segunda
                try:
                    driver.switch_to.window(window_handles[i])
                    print(f"🗂️ Cerrando pestaña {i + 1}: {driver.title[:50]}...")
                    driver.close()
                    time.sleep(1)
                except Exception as e:
                    print(f"⚠️ Error cerrando pestaña {i + 1}: {e}")
            
            # Volver a la pestaña principal (índice 0)
            driver.switch_to.window(window_handles[0])
            print("✅ Vuelto a la pestaña principal")
            time.sleep(1)
        
        # PASO 3: Intentar cerrar sesión en AFIP desde la pestaña principal
        try:
            print("🔒 Intentando cerrar sesión en AFIP...")
            
            # Buscar el icono de contribuyente AFIP
            icono_contribuyente = driver.find_element(By.ID, "iconoChicoContribuyenteAFIP")
            icono_contribuyente.click()
            time.sleep(1)
            
            # Buscar y hacer clic en el botón de salir
            boton_salir = driver.find_element(By.XPATH, '//*[@id="contBtnContribuyente"]/div[6]/button/div/div[2]')
            boton_salir.click()
            time.sleep(2)
            
            print("✅ Sesión cerrada exitosamente en AFIP")
            
        except Exception as e:
            print(f"⚠️ No se pudo cerrar sesión en AFIP (puede que no esté logueado): {e}")
        
        # PASO 4: Cerrar el navegador completamente
        print("🌐 Cerrando navegador completamente...")
        driver.quit()
        driver = None
        print("✅ Navegador cerrado exitosamente")
        
    except Exception as e:
        print(f"🚨 Error durante cierre completo: {e}")
        # Forzar cierre del navegador en caso de error
        try:
            if driver:
                driver.quit()
                driver = None
        except:
            pass
    
    print("--- CIERRE COMPLETO FINALIZADO ---\n")

# Crear el archivo de resultados
resultados = []

def human_typing(element, text):
    for char in str(text):
        element.send_keys(char)
        time.sleep(random.uniform(0.01, 0.03))

def actualizar_excel(row_index, mensaje):
    """Actualiza la última columna del archivo Excel con un mensaje de error."""
    df.at[row_index, 'Error'] = mensaje
    df.to_excel(input_excel_clientes, index=False)

def verificar_columnas_finales(df, cliente):
    """
    Verifica que solo estén las columnas correctas antes de generar Excel.
    """
    print(f"\n--- VERIFICANDO COLUMNAS FINALES PARA {cliente} ---")
    
    columnas_esperadas = ['Impuesto', 'Período', 'Ant/Cuota', 'Vencimiento', 'Saldo', 'Int. Resarcitorios']
    columnas_actuales = list(df.columns)
    
    print(f"Columnas actuales: {columnas_actuales}")
    print(f"Columnas esperadas: {columnas_esperadas}")
    
    # Verificar si hay columnas no deseadas
    columnas_extra = [col for col in columnas_actuales if col not in columnas_esperadas]
    if columnas_extra:
        print(f"⚠ Columnas extra encontradas: {columnas_extra}")
        
        # Eliminar columnas extra
        df_limpio = df[columnas_esperadas].copy()
        print(f"✓ Columnas extra eliminadas")
        return df_limpio
    else:
        print(f"✓ Solo columnas correctas presentes")
        return df

# MODIFICACIÓN: Nueva función para generar Excel en lugar de PDF
def generar_excel_desde_dataframe(df, cliente, ruta_excel):
    """Genera Excel directamente desde DataFrame - versión simplificada para anticipos."""
    try:
        print(f"\n--- GENERANDO EXCEL PARA {cliente} ---")
        
        if len(df) > 0:
            # Verificar columnas finales
            df_limpio = verificar_columnas_finales(df, cliente)
            
            # Guardar Excel sin formato especial
            df_limpio.to_excel(ruta_excel, index=False, sheet_name=f"Anticipos - {cliente}")
            print(f"DataFrame con {len(df_limpio)} registros guardado en Excel")
        else:
            # Crear Excel vacío con estructura básica
            df_vacio = pd.DataFrame(columns=['Impuesto', 'Período', 'Ant/Cuota', 'Vencimiento', 'Saldo', 'Int. Resarcitorios'])
            df_vacio.to_excel(ruta_excel, index=False, sheet_name=f"Anticipos - {cliente}")
            print("Excel vacío creado")
        
        print(f"✓ Excel generado exitosamente: {ruta_excel}")
        
    except Exception as e:
        print(f"Error generando Excel: {e}")
        import traceback
        traceback.print_exc()

def iniciar_sesion(cuit_ingresar, password, row_index):
    """Inicia sesión en el sitio web con el CUIT y contraseña proporcionados."""
    try:
        driver.get('https://auth.afip.gob.ar/contribuyente_/login.xhtml')
        element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'F1:username')))
        element.clear()
        time.sleep(2)

        human_typing(element, cuit_ingresar)
        driver.find_element(By.ID, 'F1:btnSiguiente').click()
        time.sleep(2)

        # Verificar si el CUIT es incorrecto
        try:
            error_message = driver.find_element(By.ID, 'F1:msg').text
            if error_message == "Número de CUIL/CUIT incorrecto":
                actualizar_excel(row_index, "Número de CUIL/CUIT incorrecto")
                return False
        except:
            pass

        element_pass = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'F1:password')))
        human_typing(element_pass, password)
        time.sleep(3)
        driver.find_element(By.ID, 'F1:btnIngresar').click()
        time.sleep(2)

        # Verificar si la contraseña es incorrecta
        try:
            error_message = driver.find_element(By.ID, 'F1:msg').text
            if error_message == "Clave o usuario incorrecto":
                actualizar_excel(row_index, "Clave incorrecta")
                return False
        except:
            pass

        return True
    except Exception as e:
        print(f"Error al iniciar sesión: {e}")
        actualizar_excel(row_index, "Error al iniciar sesión")
        return False

def ingresar_modulo(cuit_ingresar, password, row_index):
    """Ingresa al módulo específico del sistema de cuentas tributarias."""

    # Verificar si el botón "Ver todos" está presente y hacer clic
    boton_ver_todos = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, "Ver todos")))
    if boton_ver_todos:
        boton_ver_todos.click()
        time.sleep(2)

    # Buscar input del buscador y escribir
    buscador = driver.find_element(By.ID, 'buscadorInput')
    if buscador:
        human_typing(buscador, 'tas tr') 
        time.sleep(2)

    # Seleccionar la opción del menú
    opcion_menu = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'rbt-menu-item-0')))
    if opcion_menu:
        opcion_menu.click()
        time.sleep(2)

    # Manejar modal si aparece
    modales = driver.find_elements(By.CLASS_NAME, 'modal-content')
    if modales and modales[0].is_displayed():
        boton_continuar = driver.find_element(By.XPATH, '//button[text()="Continuar"]')
        if boton_continuar:
            boton_continuar.click()
            time.sleep(2)

    # Cambiar a la última pestaña abierta
    driver.switch_to.window(driver.window_handles[-1])

    # Verificar mensaje de error de autenticación
    error_message_elements = driver.find_elements(By.TAG_NAME, 'pre')
    if error_message_elements and error_message_elements[0].text == "Ha ocurrido un error al autenticar, intente nuevamente.":
        actualizar_excel(row_index, "Error autenticacion")
        driver.refresh()
        time.sleep(2)

    # Verificar si es necesario iniciar sesión nuevamente
    username_input = driver.find_elements(By.ID, 'F1:username')
    if username_input:
        username_input[0].clear()
        time.sleep(2)
        human_typing(username_input[0], cuit_ingresar)
        driver.find_element(By.ID, 'F1:btnSiguiente').click()
        time.sleep(2)

        password_input = driver.find_elements(By.ID, 'F1:password')
        if password_input:
            human_typing(password_input[0], password)
            time.sleep(2)
            driver.find_element(By.ID, 'F1:btnIngresar').click()
            time.sleep(1)
            actualizar_excel(row_index, "Error volver a iniciar sesion")

def seleccionar_cuit_representado(cuit_representado):
    """Selecciona el CUIT representado en el sistema."""
    try:
        select_present = EC.presence_of_element_located((By.NAME, "$PropertySelection"))
        if WebDriverWait(driver, 5).until(select_present):
            current_selection = Select(driver.find_element(By.NAME, "$PropertySelection")).first_selected_option.text
            if current_selection != str(cuit_representado):
                select_element = Select(driver.find_element(By.NAME, "$PropertySelection"))
                select_element.select_by_visible_text(str(cuit_representado))
    except Exception:
        try:
            cuit_element = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'span.cuit')))
            cuit_text = cuit_element.text.replace('-', '')
            if cuit_text != str(cuit_representado):
                print(f"El CUIT ingresado no coincide con el CUIT representado: {cuit_representado}")
                return False
        except Exception as e:
            print(f"Error al verificar CUIT: {e}")
            return False
    # Esperar que el popup esté visible y hacer clic en el botón de cerrar por XPATH
    try:
    # Usamos el XPATH para localizar el botón de cerrar
        xpath_popup = "/html/body/div[2]/div[2]/div/div/a"
        element_popup = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xpath_popup)))
        element_popup.click()
        print("Popup cerrado exitosamente.")
    except Exception as e:
        print(f"Error al intentar cerrar el popup: {e}")
    return True

def configurar_select_100_mejorado(driver):
    print(f"\n--- CONFIGURANDO SELECT A 100 REGISTROS (VERSIÓN MEJORADA) ---")
    
    try:
        # Esperar inicial
        time.sleep(1)
        print("✓ Esperando 1 segundos antes de configurar select...")
        
        # ESTRATEGIA 1: Buscar el select con múltiples selectores
        select_element = None
        selectores_select = [
            "select.mx-2.form-control.form-control-sm",
            "select[class*='form-control-sm']",
            "select[class*='mx-2']",
            "//div[@class='dtable__footer']//select",
            "//div[contains(@class, 'pagination')]//select",
            "//select[contains(@class, 'form-control')]",
            "//select"  # Último recurso
        ]
        
        for i, selector in enumerate(selectores_select):
            try:
                if selector.startswith("//"):
                    elements = driver.find_elements(By.XPATH, selector)
                else:
                    elements = driver.find_elements(By.CSS_SELECTOR, selector)
                
                if elements:
                    # Verificar cuál es el select correcto (que esté visible y tenga opciones)
                    for element in elements:
                        if element.is_displayed():
                            select_element = element
                            print(f"✓ Select encontrado con selector {i+1}: {selector}")
                            break
                    
                    if select_element:
                        break
                        
            except Exception as e:
                continue
        
        if not select_element:
            print("✗ No se encontró ningún select, continuando sin cambio...")
            time.sleep(1)
            return False
        
        # ESTRATEGIA 2: Analizar el select encontrado
        print(f"\n--- ANALIZANDO SELECT ENCONTRADO ---")
        
        # Hacer scroll al elemento
        driver.execute_script("arguments[0].scrollIntoView(true);", select_element)
        time.sleep(1)
        
        # Obtener información del select
        current_value = select_element.get_attribute('value')
        print(f"Valor actual del select: {current_value}")
        
        # ESTRATEGIA 3: Obtener opciones de manera más robusta
        opciones_info = driver.execute_script("""
            var select = arguments[0];
            var opciones = [];
            
            for (var i = 0; i < select.options.length; i++) {
                var option = select.options[i];
                opciones.push({
                    value: option.value,
                    text: option.text,
                    index: i
                });
            }
            
            return opciones;
        """, select_element)
        
        print(f"Opciones encontradas: {len(opciones_info)}")
        for opcion in opciones_info:
            print(f"  - Valor: '{opcion['value']}', Texto: '{opcion['text']}', Índice: {opcion['index']}")
        
        # Verificar si ya está en 100
        if current_value == "100":
            print("✓ Select ya está configurado en 100")
            time.sleep(1)
            return True
        
        # ESTRATEGIA 4: Buscar la opción 100
        opcion_100_encontrada = None
        for opcion in opciones_info:
            if opcion['value'] == '100' or opcion['text'] == '100':
                opcion_100_encontrada = opcion
                break
        
        if not opcion_100_encontrada:
            print("⚠ No se encontró opción '100' en el select")
            # Intentar con la opción más alta disponible
            valores_numericos = []
            for opcion in opciones_info:
                try:
                    if opcion['value'] and opcion['value'].isdigit():
                        valores_numericos.append(int(opcion['value']))
                except:
                    pass
            
            if valores_numericos:
                max_valor = max(valores_numericos)
                print(f"Usando valor máximo disponible: {max_valor}")
                target_value = str(max_valor)
                target_index = None
                for opcion in opciones_info:
                    if opcion['value'] == target_value:
                        target_index = opcion['index']
                        break
            else:
                print("✗ No se encontraron opciones válidas")
                time.sleep(1)
                return False
        else:
            target_value = "100"
            target_index = opcion_100_encontrada['index']
            print(f"✓ Opción 100 encontrada en índice {target_index}")
        
        # ESTRATEGIA 5: Múltiples métodos de cambio
        exito_cambio = False
                   
        # Método 2: Select by index
        if not exito_cambio:
            try:
                print("Intentando Método 2: Select by index...")
                from selenium.webdriver.support.ui import Select
                select_obj = Select(select_element)
                select_obj.select_by_index(target_index)
                time.sleep(1)
                
                new_value = select_element.get_attribute('value')
                if new_value == target_value:
                    print(f"✓ Método 2 exitoso: Select cambiado a {target_value}")
                    exito_cambio = True
                else:
                    print(f"✗ Método 2 falló: Valor sigue siendo {new_value}")
                    
            except Exception as e:
                print(f"✗ Método 2 falló: {e}")
                         
        # ESTRATEGIA 6: Verificación visual y de DOM
        if exito_cambio:
            print(f"\n--- VERIFICANDO CAMBIO ---")
            time.sleep(1)
            
            # Verificar valor del select
            valor_final = select_element.get_attribute('value')
            print(f"Valor final del select: {valor_final}")
            
            # Verificar información de paginación
            try:
                # Buscar elementos que muestren información de registros
                info_elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'registros') or contains(text(), 'Mostrando') or contains(text(), 'de')]")
                
                for elem in info_elements:
                    if elem.is_displayed():
                        texto = elem.text.strip()
                        if texto and ('registros' in texto.lower() or 'mostrando' in texto.lower()):
                            print(f"Información de paginación: {texto}")
                            break
                            
            except Exception as e:
                print(f"No se pudo obtener información de paginación: {e}")
            
            # Verificar número de filas visibles en la tabla
            try:
                filas_visibles = driver.find_elements(By.XPATH, "//tbody//tr[@role='row']")
                print(f"Filas visibles en la tabla: {len(filas_visibles)}")
                
                if len(filas_visibles) > 10:
                    print("✓ El cambio parece haber funcionado (más de 10 filas visibles)")
                else:
                    print("⚠ Posible problema: solo se ven 10 o menos filas")
                    
            except Exception as e:
                print(f"No se pudo contar filas visibles: {e}")
        
        # Esperar antes de continuar
        print("✓ Esperando 2 segundos antes de extraer datos...")
        time.sleep(2)
        
        return exito_cambio
        
    except Exception as e:
        print(f"✗ Error general configurando select: {e}")
        time.sleep(1)
        return False

def exportar_desde_html(ubicacion_descarga, cuit_representado, cliente):
    try:
        print(f"=== INICIANDO EXTRACCIÓN HTML PARA CLIENTE: {cliente} ===")
        
        # Verificar que estamos en la página correcta
        print(f"URL actual: {driver.current_url}")
        print(f"Título de la página: {driver.title}")
        
        # Esperar a que la página se cargue completamente
        time.sleep(3)
        # PASO 1: Verificar si hay iframe y cambiar a él
        print(f"\n--- VERIFICANDO Y CAMBIANDO AL IFRAME ---")
        
        iframe_encontrado = False

        try:
            # Buscar iframe específico del SCT
            iframe_selector = "iframe[src*='homeContribuyente']"
            iframe_element = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, iframe_selector)))
            
            print(f"✓ Iframe encontrado: {iframe_element.get_attribute('src')}")
            
            # Cambiar al iframe
            driver.switch_to.frame(iframe_element)
            iframe_encontrado = True
            print("✓ Cambiado al iframe exitosamente")
            
            # Esperar a que el contenido del iframe se cargue COMPLETAMENTE
            time.sleep(3)  # Aumentar tiempo de espera
            
            # Esperar a que Vue.js termine de renderizar
            WebDriverWait(driver, 20).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            print("✓ Contenido del iframe cargado completamente")
            
        except Exception as e:
            print(f"✗ Error cambiando al iframe: {e}")
            print("Continuando en el documento principal...")
        
        # PASO 2: BÚSQUEDA MEJORADA del elemento "$ Deudas"
        print(f"\n--- BÚSQUEDA MEJORADA DE ELEMENTO '$ DEUDAS' ---")
        
        elemento_deudas = None
        numero_deudas = 0

        try:
            # PRIMERA BÚSQUEDA: Esperar explícitamente a que aparezcan las pestañas
            print("Esperando a que las pestañas se carguen...")
            
            try:
                # Esperar a que aparezca cualquier elemento de navegación
                WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "[role='tablist'], .nav-tabs, .tab-content"))
                )
                print("✓ Elementos de navegación detectados")
            except:
                print("⚠ No se detectaron elementos de navegación estándar")
            
            # SEGUNDA BÚSQUEDA: Buscar TODOS los elementos que contengan "Deudas"
            print("Buscando TODOS los elementos con 'Deudas'...")
            
            # Usar JavaScript para buscar elementos
            elementos_deudas_js = driver.execute_script("""
                var elementos = [];
                var allElements = document.querySelectorAll('*');
                
                for (var i = 0; i < allElements.length; i++) {
                    var element = allElements[i];
                    if (element.textContent && element.textContent.includes('Deudas')) {
                        elementos.push({
                            tagName: element.tagName,
                            className: element.className,
                            id: element.id,
                            textContent: element.textContent.substring(0, 100),
                            isVisible: element.offsetParent !== null,
                            role: element.getAttribute('role'),
                            href: element.href || ''
                        });
                    }
                }
                
                return elementos;
            """)
            
            print(f"JavaScript encontró {len(elementos_deudas_js)} elementos con 'Deudas':")
            for i, elem in enumerate(elementos_deudas_js[:10]):  # Mostrar primeros 10
                print(f"  {i+1}. Tag: {elem['tagName']}, Texto: '{elem['textContent'][:50]}...', Visible: {elem['isVisible']}")
                print(f"      Clase: {elem['className']}, Role: {elem['role']}")
            
            # TERCERA BÚSQUEDA: Intentar selectores más amplios
            print("\nBuscando con selectores amplios...")
            
            selectores_amplios = [
                # Buscar cualquier elemento que contenga "Deudas"
                "//*[contains(text(), 'Deudas')]",
                "//*[contains(., 'Deudas')]",
                # Buscar elementos clickeables
                "//a[contains(text(), 'Deudas')]",
                "//button[contains(text(), 'Deudas')]",
                "//div[contains(text(), 'Deudas')]",
                "//span[contains(text(), 'Deudas')]",
                "//li[contains(text(), 'Deudas')]",
                # Buscar por atributos comunes de Bootstrap/Vue
                "//*[@data-*][contains(text(), 'Deudas')]",
                "//*[@v-*][contains(text(), 'Deudas')]",
                # Buscar por clases de Bootstrap
                "//*[contains(@class, 'nav')][contains(text(), 'Deudas')]",
                "//*[contains(@class, 'tab')][contains(text(), 'Deudas')]",
                "//*[contains(@class, 'btn')][contains(text(), 'Deudas')]"
            ]
            
            for i, selector in enumerate(selectores_amplios, 1):
                try:
                    elementos = driver.find_elements(By.XPATH, selector)
                    if elementos:
                        print(f"  Selector {i} encontró {len(elementos)} elementos")
                        
                        for j, elem in enumerate(elementos):
                            try:
                                if elem.is_displayed():
                                    elem_texto = elem.text.strip()
                                    if 'Deudas' in elem_texto:
                                        print(f"    ✓ Elemento visible: '{elem_texto}'")
                                        
                                        # Este es nuestro candidato
                                        elemento_deudas = elem
                                        
                                        # Buscar número de deudas
                                        import re
                                        numeros = re.findall(r'\d+', elem_texto)
                                        if numeros:
                                            numero_deudas = int(numeros[0])
                                            print(f"    ★ Número de deudas: {numero_deudas}")
                                        else:
                                            numero_deudas = 1
                                            
                                        break
                                        
                            except Exception as e:
                                continue
                        
                        if elemento_deudas:
                            break
                            
                except Exception as e:
                    continue
            
            # CUARTA BÚSQUEDA: Si todavía no encuentra, hacer una búsqueda exhaustiva
            if not elemento_deudas:
                print("\n--- BÚSQUEDA EXHAUSTIVA ---")
                
                # Guardar HTML completo del iframe para análisis
                iframe_html = driver.page_source
                html_iframe_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), f"debug_iframe_completo_{cliente}.html")
                with open(html_iframe_file, 'w', encoding='utf-8') as f:
                    f.write(iframe_html)
                print(f"HTML completo del iframe guardado: {html_iframe_file}")
                
                # Buscar "Deudas" en el HTML
                if 'Deudas' in iframe_html:
                    print("✓ 'Deudas' encontrado en el HTML del iframe")
                    
                    # Intentar hacer clic por coordenadas si es necesario
                    try:
                        # Buscar cualquier elemento que contenga el texto
                        elemento_cualquiera = driver.find_element(By.XPATH, "//*[contains(text(), 'Deudas')]")
                        if elemento_cualquiera:
                            elemento_deudas = elemento_cualquiera
                            numero_deudas = 1
                            print("✓ Elemento encontrado con búsqueda de emergencia")
                    except:
                        pass
                else:
                    print("✗ 'Deudas' NO encontrado en el HTML del iframe")
                    
        except Exception as e:
            print(f"Error en búsqueda de elemento Deudas: {e}")

        if not elemento_deudas:
            print("✗ No se encontró el elemento '$ Deudas'")
            
            # Generar Excel vacío y salir
            nombre_excel = f"Anticipos - {cliente}.xlsx"
            ruta_excel = os.path.join(ubicacion_descarga, nombre_excel)
            
            df_vacio = pd.DataFrame()
            generar_excel_desde_dataframe(df_vacio, cliente, ruta_excel)
            
            # Volver al contenido principal antes de salir
            if iframe_encontrado:
                driver.switch_to.default_content()
            
            return
        
        print(f"✓ Elemento '$ Deudas' encontrado con {numero_deudas} deudas")

        # PASO 3: Decidir si hacer clic o generar Excel vacío
        datos_tabla = []
        
        if numero_deudas >= 1:
            print(f"\n--- HACIENDO CLIC EN '$ DEUDAS' (tiene {numero_deudas} deudas) ---")
            
            try:
                # Hacer scroll al elemento para asegurar que esté visible
                driver.execute_script("arguments[0].scrollIntoView(true);", elemento_deudas)
                time.sleep(2)
                
                # Intentar clic normal primero
                elemento_deudas.click()
                print("✓ Clic normal en '$ Deudas' realizado")
                time.sleep(3)  # Esperar más tiempo para que cargue la tabla

                # USAR LA FUNCIÓN MEJORADA PARA CONFIGURAR SELECT
                exito_select = configurar_select_100_mejorado(driver)
            
                if not exito_select:
                    print("⚠ No se pudo configurar el select, continuando con los registros disponibles...")             
            except Exception as e:
                print(f"Error en clic normal: {e}")
                try:
                    # Intentar clic con JavaScript
                    driver.execute_script("arguments[0].click();", elemento_deudas)
                    print("✓ Clic con JavaScript realizado")
                    time.sleep(3)
                except Exception as e2:
                    print(f"Error en clic JavaScript: {e2}")
                    
                    # Volver al contenido principal antes de salir
                    if iframe_encontrado:
                        driver.switch_to.default_content()
                    return
            
            # PASO 3.5: CONFIGURAR SELECT A 100 REGISTROS
            print(f"\n--- CONFIGURANDO SELECT A 100 REGISTROS ---")

            try:
                # Esperar 4 segundos antes de empezar a configurar
                time.sleep(2)
                print("✓ Esperando 2 segundos antes de configurar select...")
                
                # Esperar a que el select esté presente
                WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "select.mx-2.form-control.form-control-sm"))
                )
                
                # Buscar el select en el footer de la tabla
                try:
                    select_element = driver.find_element(By.CSS_SELECTOR, "select.mx-2.form-control.form-control-sm")
                    print("✓ Select encontrado con CSS selector")
                except:
                    # Fallback: buscar por múltiples selectores
                    selectores_fallback = [
                        "//select[contains(@class, 'form-control-sm')]",
                        "//select[contains(@class, 'mx-2')]", 
                        "//div[@class='dtable__footer']//select",
                        "//div[contains(@class, 'dtable')]//select"
                    ]
                    
                    select_element = None
                    for selector in selectores_fallback:
                        try:
                            select_element = driver.find_element(By.XPATH, selector)
                            print(f"✓ Select encontrado con selector fallback: {selector}")
                            break
                        except:
                            continue
                    
                    if not select_element:
                        print("⚠ No se encontró el select, continuando sin cambiar...")
                        # Continuar sin el select, pero esperar antes de extraer datos
                        time.sleep(2)
                        print("✓ Esperando 2 segundos antes de extraer datos...")
                    else:
                        # Procesar el select encontrado
                        pass
                
                if 'select_element' in locals() and select_element:
                    # Verificar el valor actual del select
                    current_value = select_element.get_attribute('value')
                    print(f"Valor actual del select: {current_value}")
                    
                    # Buscar todas las opciones disponibles
                    options = select_element.find_elements(By.TAG_NAME, "option")
                    print(f"Opciones disponibles: {[opt.text for opt in options]}")
                    
                    # Verificar si ya está en 100
                    if current_value == "100":
                        print("✓ Select ya está configurado en 100")
                    else:
                        # Cambiar a 100
                        try:
                            # Método 1: Usar Select de Selenium
                            from selenium.webdriver.support.ui import Select
                            select_obj = Select(select_element)
                            select_obj.select_by_value("100")
                            print("✓ Select cambiado a 100 usando Select()")
                            
                        except Exception as e1:
                            print(f"Método 1 falló: {e1}")
                            try:
                                # Método 2: Hacer clic en la opción 100
                                option_100 = select_element.find_element(By.XPATH, ".//option[@value='100']")
                                option_100.click()
                                print("✓ Select cambiado a 100 haciendo clic en option")
                                
                            except Exception as e2:
                                print(f"Método 2 falló: {e2}")
                                try:
                                    # Método 3: JavaScript
                                    driver.execute_script("arguments[0].value = '100'; arguments[0].dispatchEvent(new Event('change'));", select_element)
                                    print("✓ Select cambiado a 100 usando JavaScript")
                                    
                                except Exception as e3:
                                    print(f"Método 3 falló: {e3}")
                                    print("⚠ No se pudo cambiar el select, continuando...")
                    
                    # Esperar a que la tabla se actualice después del cambio
                    time.sleep(3)
                    print("✓ Esperando 3 segundos para que la tabla se actualice...")
                    
                    # Verificar el cambio
                    try:
                        new_value = select_element.get_attribute('value')
                        print(f"Nuevo valor del select: {new_value}")
                        
                        # Buscar el texto que indica cuántos registros se muestran
                        try:
                            registro_text_elements = driver.find_elements(By.XPATH, "//*[contains(text(), 'registros') or contains(text(), 'de')]")
                            for elem in registro_text_elements:
                                if 'registros' in elem.text or 'de' in elem.text:
                                    print(f"Información de registros: {elem.text}")
                                    break
                        except:
                            pass
                            
                    except Exception as e:
                        print(f"Error verificando el cambio: {e}")
                
                # Esperar 2 segundos antes de empezar a extraer datos
                time.sleep(2)
                print("✓ Esperando 2 segundos antes de extraer datos de la tabla...")

            except Exception as e:
                print(f"Error configurando select: {e}")
                # En caso de error, al menos esperar antes de continuar
                time.sleep(2)
                print("✓ Esperando 2 segundos antes de continuar (por error en select)...")
            
            # PASO 4: Extraer datos de la tabla (dentro del iframe) - VERSIÓN MODIFICADA PARA ANTICIPOS
            print(f"\n--- EXTRAYENDO DATOS CON FILTROS PARA ANTICIPOS ---")

            try:
                # Esperar a que la tabla se cargue dentro del iframe
                WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.XPATH, "//table[@role='table']")))
                
                # Buscar la tabla específica con 12 columnas
                tabla = None

                try:
                    tabla = driver.find_element(By.XPATH, "//table[@role='table'][@aria-colcount='12']")
                    aria_rowcount = tabla.get_attribute('aria-rowcount')
                    aria_colcount = tabla.get_attribute('aria-colcount')
                    print(f"✓ Tabla de 12 columnas encontrada: {aria_rowcount} filas, {aria_colcount} columnas")
                except:
                    # Fallback a búsqueda general
                    tablas = driver.find_elements(By.XPATH, "//table[@role='table']")
                    if tablas:
                        tabla = tablas[0]
                        print(f"ℹ Usando primera tabla como fallback")
                    else:
                        print("✗ No se encontró tabla")
                        if iframe_encontrado:
                            driver.switch_to.default_content()
                        return
                
                # MAPEO COMPLETO DE TODAS LAS COLUMNAS
                mapeo_columnas_completo = {
                    '1': 'Establecimiento',        # Para luego eliminar
                    '2': 'Concepto',              # Para luego eliminar  
                    '3': 'Subconcepto',           # Para luego eliminar
                    '4': 'Impuesto',              # ✓ MANTENER
                    '5': 'Concepto',              # Para luego eliminar (duplicado)
                    '6': 'Subconcepto',           # Para luego eliminar (duplicado)  
                    '7': 'Período',               # ✓ MANTENER
                    '8': 'Ant/Cuota',             # ✓ MANTENER
                    '9': 'Vencimiento',           # ✓ MANTENER
                    '10': 'Saldo',                # ✓ MANTENER
                    '11': 'Int. Resarcitorios',   # ✓ MANTENER
                    '12': 'Int. Punitorio'        # Para luego eliminar
                }
             
                print(f"Mapeo completo definido: {len(mapeo_columnas_completo)} columnas")

                # MODIFICACIÓN: FILTROS ESPECÍFICOS PARA ANTICIPOS
                # Solo Ganancias Sociedades
                impuestos_incluir = [
                    'ganancias sociedades'
                ]
                
                print(f"Filtros de impuestos para anticipos: {impuestos_incluir}")
                
                # Período: 2025
                # Vencimiento: entre 01/06/2025 y 21/07/2025
                fecha_vencimiento_inicio = datetime(2025, 8, 1).date()
                fecha_vencimiento_fin = datetime(2025, 8, 19).date()
                
                print(f"Filtro de período: 2025")
                print(f"Filtro de vencimiento: desde {fecha_vencimiento_inicio} hasta {fecha_vencimiento_fin}")

                # EXTRAER FILAS DE DATOS CON FILTROS ESPECÍFICOS PARA ANTICIPOS
                try:
                    filas_datos = tabla.find_elements(By.XPATH, ".//tbody//tr[@role='row']")
                    print(f"Filas de datos encontradas: {len(filas_datos)}")
                    
                    datos_extraidos = 0
                    datos_filtrados = 0

                    for i, fila in enumerate(filas_datos):
                        try:
                            print(f"\n--- Procesando fila {i+1} ---")
                            
                            # Extraer datos de TODAS las columnas primero
                            datos_fila_completa = {}
                            fila_valida = True
                            for aria_colindex, nombre_columna in mapeo_columnas_completo.items():
                                try:
                                    celda = fila.find_element(By.XPATH, f".//td[@aria-colindex='{aria_colindex}'][@role='cell']")
                                    texto_celda = celda.text.strip()

                                    # Limpiar valores monetarios
                                    if nombre_columna in ['Saldo', 'Int. Resarcitorios', 'Int. Punitorio']:
                                        if not texto_celda or texto_celda in ['', '-', 'N/A']:
                                            texto_celda = '0'
                                        else:
                                            # Limpiar formato monetario: $ 178.468,79 → 178468.79
                                            texto_limpio = texto_celda.replace('$', '').replace(' ', '').strip()

                                            # Si tiene formato argentino (puntos como separadores de miles, coma como decimal)
                                            if ',' in texto_limpio and '.' in texto_limpio:
                                                # Formato: 178.468,79 → 178468.79
                                                partes = texto_limpio.split(',')
                                                if len(partes) == 2:
                                                    parte_entera = partes[0].replace('.', '')
                                                    parte_decimal = partes[1]
                                                    texto_celda = f"{parte_entera}.{parte_decimal}"
                                                else:
                                                    texto_celda = texto_limpio.replace('.', '').replace(',', '.')
                                            elif ',' in texto_limpio:
                                                # Solo coma decimal: 1234,56 → 1234.56
                                                texto_celda = texto_limpio.replace(',', '.')
                                            elif '.' in texto_limpio:
                                                # Verificar si es separador de miles o decimal
                                                if len(texto_limpio.split('.')[-1]) <= 2:
                                                    # Probablemente decimal
                                                    texto_celda = texto_limpio
                                                else:
                                                    # Probablemente separador de miles
                                                    texto_celda = texto_limpio.replace('.', '')
                                            else:
                                                texto_celda = texto_limpio
                                            # Validar que sea numérico
                                            try:
                                                float(texto_celda)
                                            except ValueError:
                                                texto_celda = '0'

                                    datos_fila_completa[nombre_columna] = texto_celda
                                    print(f"  {nombre_columna} (col-{aria_colindex}): '{texto_celda}'")

                                except Exception as e:
                                    # Manejo de errores por columna
                                    if nombre_columna in ['Saldo', 'Int. Resarcitorios', 'Int. Punitorio']:
                                        datos_fila_completa[nombre_columna] = '0'
                                        print(f"  {nombre_columna} (col-{aria_colindex}): '0' (por defecto)")
                                    else:
                                        datos_fila_completa[nombre_columna] = ''
                                        print(f"  {nombre_columna} (col-{aria_colindex}): '' (error: {str(e)[:50]}...)")
                                        if nombre_columna in ['Impuesto', 'Vencimiento', 'Período']:  # Campos críticos para anticipos
                                            fila_valida = False
                            
                            # MODIFICACIÓN: APLICAR FILTROS ESPECÍFICOS PARA ANTICIPOS
                            if fila_valida:
                                
                                # FILTRO 1: Verificar impuesto (solo Ganancias Sociedades)
                                impuesto_texto = datos_fila_completa.get('Impuesto', '').lower()
                                impuesto_valido = 'ganancias sociedades' in impuesto_texto
                                
                                if not impuesto_valido:
                                    print(f"  ✗ Fila {i+1} descartada: no es Ganancias Sociedades ('{impuesto_texto}')")
                                    continue
                                
                                # FILTRO 2: Verificar período (debe ser 2025)
                                periodo_texto = datos_fila_completa.get('Período', '')
                                periodo_valido = '2025' in periodo_texto
                                
                                if not periodo_valido:
                                    print(f"  ✗ Fila {i+1} descartada: período no es 2025 ('{periodo_texto}')")
                                    continue
                                
                                # FILTRO 3: Verificar fecha de vencimiento (debe estar entre 01/06/2025 y 30/06/2025)
                                fecha_vencimiento_texto = datos_fila_completa.get('Vencimiento', '')
                                fecha_vencida_valida = False
                                
                                if fecha_vencimiento_texto:
                                    try:
                                        # Parsear fecha formato dd/mm/yyyy
                                        fecha_vencimiento = datetime.strptime(fecha_vencimiento_texto, "%d/%m/%Y").date()
                                        
                                        # Verificar si está en el rango junio 2025
                                        if fecha_vencimiento_inicio <= fecha_vencimiento <= fecha_vencimiento_fin:
                                            fecha_vencida_valida = True
                                            print(f"  ✓ Fecha de vencimiento válida para anticipos: {fecha_vencimiento}")
                                        else:
                                            print(f"  ✗ Fecha fuera del rango junio 2025: {fecha_vencimiento}")
                                            continue
                                            
                                    except ValueError:
                                        print(f"  ✗ Formato de fecha inválido: '{fecha_vencimiento_texto}'")
                                        continue
                                else:
                                    print(f"  ✗ Sin fecha de vencimiento")
                                    continue
                                
                                # FILTRO 4: Verificar datos mínimos
                                tiene_datos_minimos = bool(impuesto_texto) and bool(fecha_vencimiento_texto) and bool(periodo_texto)
                                
                                if tiene_datos_minimos and impuesto_valido and periodo_valido and fecha_vencida_valida:
                                    # Agregar metadata de procesamiento
                                    datos_fila_completa['Fecha_Procesamiento'] = datetime.now().date().strftime("%Y-%m-%d")
                                    datos_fila_completa['Fuente'] = 'SCT_Web_Anticipos'
                                    
                                    datos_tabla.append(datos_fila_completa)
                                    datos_extraidos += 1
                                    
                                    print(f"  ✓ Fila {i+1} INCLUIDA en reporte de anticipos")
                                    print(f"    Resumen: {datos_fila_completa['Impuesto'][:30]}... | {datos_fila_completa['Período']} | {datos_fila_completa['Vencimiento']} | ${datos_fila_completa['Saldo']}")
                                else:
                                    print(f"  ✗ Fila {i+1} descartada: datos insuficientes")
                                    
                            else:
                                print(f"  ✗ Fila {i+1} descartada: fila inválida")
                            
                            datos_filtrados += 1    
                        except Exception as e:
                            print(f"  ✗ Error procesando fila {i+1}: {e}")
                            continue

                    print(f"\n✓ RESUMEN DE EXTRACCIÓN Y FILTRADO PARA ANTICIPOS:")
                    print(f"  - Filas procesadas: {len(filas_datos)}")
                    print(f"  - Filas filtradas: {datos_filtrados}")
                    print(f"  - Registros de anticipos incluidos: {datos_extraidos}")
                    print(f"  - Tasa de inclusión: {(datos_extraidos/len(filas_datos)*100):.1f}%" if len(filas_datos) > 0 else "  - Sin filas para procesar")
                    
                    # Mostrar resumen específico para anticipos
                    if datos_tabla:
                        periodos_encontrados = {}
                        for fila in datos_tabla:
                            periodo = fila['Período']
                            if periodo in periodos_encontrados:
                                periodos_encontrados[periodo] += 1
                            else:
                                periodos_encontrados[periodo] = 1
                        
                        print(f"\n  - Distribución por período:")
                        for periodo, cantidad in periodos_encontrados.items():
                            print(f"    {periodo}: {cantidad} registros")
                    
                    # Diagnóstico si no se extrajeron datos
                    if datos_extraidos == 0:
                        print(f"\n--- DIAGNÓSTICO: SIN ANTICIPOS ENCONTRADOS ---")
                        
                        # Verificar una fila de muestra para diagnóstico
                        if len(filas_datos) > 0:
                            print("Analizando primera fila para diagnóstico...")
                            fila_muestra = filas_datos[0]
                            
                            for aria_colindex, nombre_columna in mapeo_columnas_completo.items():
                                try:
                                    celda = fila_muestra.find_element(By.XPATH, f".//td[@aria-colindex='{aria_colindex}'][@role='cell']")
                                    texto = celda.text.strip()
                                    print(f"    {nombre_columna} (col-{aria_colindex}): '{texto[:50]}...'")
                                except:
                                    print(f"    {nombre_columna} (col-{aria_colindex}): ERROR - No encontrada")
                            
                            # Guardar HTML para análisis
                            tabla_html = tabla.get_attribute('outerHTML')
                            archivo_debug = os.path.join(os.path.dirname(os.path.abspath(__file__)), f"debug_anticipos_{cliente}.html")
                            with open(archivo_debug, 'w', encoding='utf-8') as f:
                                f.write(tabla_html)
                            print(f"    HTML guardado para análisis: {archivo_debug}") 

                except Exception as e:
                    print(f"Error extrayendo filas con filtros para anticipos: {e}")
                    import traceback
                    traceback.print_exc()

            except Exception as e:
                print(f"Error general en extracción filtrada para anticipos: {e}")
                import traceback
                traceback.print_exc()
                
                if iframe_encontrado:
                    driver.switch_to.default_content()
                return
        
        # PASO 5: Volver al contenido principal antes de generar Excel
        if iframe_encontrado:
            print("\n--- VOLVIENDO AL CONTENIDO PRINCIPAL ---")
            driver.switch_to.default_content()
            print("✓ Vuelto al contenido principal")
        
        # MODIFICACIÓN: PASO 6: Generar Excel en lugar de PDF
        print(f"\n--- GENERANDO EXCEL PARA ANTICIPOS ---")
        
        nombre_excel = f"Anticipos - {cliente}.xlsx"
        ruta_excel = os.path.join(ubicacion_descarga, nombre_excel)
        
        if datos_tabla:
            df = pd.DataFrame(datos_tabla)
            print(f"DataFrame creado con {len(df)} filas y {len(df.columns)} columnas")
            print(f"Columnas: {list(df.columns)}")
            
            # Los datos ya vienen filtrados específicamente para anticipos
            df_filtrado = df.copy()
            
            print(f"DataFrame final: {len(df_filtrado)} registros para Excel de anticipos")
            
        else:
            df_filtrado = pd.DataFrame()

        
        # Generar Excel usando la función modificada
        generar_excel_desde_dataframe(df_filtrado, cliente, ruta_excel)
        
        print(f"✓ Excel generado: {ruta_excel}")

    except Exception as e:
        print(f"✗ ERROR GENERAL: {e}")
        import traceback
        traceback.print_exc()
        
        # Asegurar que volvemos al contenido principal en caso de error
        try:
            driver.switch_to.default_content()
        except:
            pass

def procesar_cliente_completo(cuit_ingresar, cuit_representado, password, cliente, ubicacion_descarga, indice):
    """
    Función unificada que procesa completamente un cliente con sesión limpia.
    MODIFICACIÓN: Ahora recibe ubicacion_descarga como parámetro.
    """
    print(f"\n{'='*80}")
    print(f"🚀 INICIANDO PROCESAMIENTO DE CLIENTE: {cliente}")
    print(f"📋 CUIT Login: {cuit_ingresar} | CUIT Representado: {cuit_representado}")
    print(f"📁 Ubicación descarga: {ubicacion_descarga}")
    print(f"{'='*80}")
    
    try:
        # PASO 1: Configurar navegador nuevo y limpio
        print("🌐 PASO 1: Configurando navegador nuevo...")
        configurar_nuevo_navegador()
        
        # PASO 2: Iniciar sesión
        print("🔐 PASO 2: Iniciando sesión en AFIP...")
        control_sesion = iniciar_sesion(cuit_ingresar, password, indice)
        
        if not control_sesion:
            print(f"❌ Error en autenticación para {cliente}")
            return False
        
        # PASO 3: Ingresar al módulo SCT
        print("🏢 PASO 3: Ingresando al módulo de Sistema de Cuentas Tributarias...")
        ingresar_modulo(cuit_ingresar, password, indice)
        
        # PASO 4: Cerrar popup inicial
        try:
            xpath_popup = "/html/body/div[2]/div[2]/div/div/a"
            element_popup = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xpath_popup)))
            element_popup.click()
            print("✅ Popup inicial cerrado exitosamente")
        except Exception as e:
            print(f"⚠️ Error al intentar cerrar popup inicial: {e}")
        
        # PASO 5: Seleccionar CUIT representado
        print("🎯 PASO 5: Seleccionando CUIT representado...")
        if not seleccionar_cuit_representado(cuit_representado):
            print(f"❌ Error seleccionando CUIT representado para {cliente}")
            return False
        
        # MODIFICACIÓN: PASO 6: Extraer datos y generar Excel (no PDF)
        print("📊 PASO 6: Extrayendo datos y generando Excel de anticipos...")
        exportar_desde_html(ubicacion_descarga, cuit_representado, cliente)
        
        print(f"✅ CLIENTE {cliente} PROCESADO EXITOSAMENTE")
        return True
        
    except Exception as e:
        print(f"❌ ERROR GENERAL procesando cliente {cliente}: {e}")
        import traceback
        traceback.print_exc()
        actualizar_excel(indice, f"Error general: {str(e)[:50]}...")
        return False
    
    finally:
        # PASO 7: SIEMPRE cerrar sesión y navegador al final
        print("🔒 PASO 7: Cerrando sesión y navegador...")
        cerrar_sesion_y_navegador()
        print(f"🏁 PROCESAMIENTO DE {cliente} FINALIZADO\n")

# ========== VERIFICACIÓN DE FUNCIONES ==========
def verificar_funciones_disponibles():
    """Verifica que todas las funciones necesarias estén disponibles."""
    funciones_necesarias = ['generar_excel_desde_dataframe', 'exportar_desde_html', 'procesar_cliente_completo']
    
    current_module = sys.modules[__name__]
    
    print("=== VERIFICACIÓN DE FUNCIONES ===")
    for func_name in funciones_necesarias:
        if hasattr(current_module, func_name):
            print(f"✓ Función {func_name} disponible")
        else:
            print(f"✗ Función {func_name} NO disponible")
    
    # Mostrar algunas funciones disponibles
    all_functions = [name for name, obj in inspect.getmembers(current_module) if inspect.isfunction(obj)]
    print(f"Total funciones disponibles: {len(all_functions)}")

# Función para convertir Excel a CSV utilizando xlwings (mantenida para compatibilidad)
def excel_a_csv(input_folder, output_folder):
    for excel_file in glob.glob(os.path.join(input_folder, "*.xlsx")):
        try:
            app = xw.App(visible=False)
            wb = app.books.open(excel_file)
            sheet = wb.sheets[0]
            df = sheet.used_range.options(pd.DataFrame, header=1, index=False).value

            # Convertir la columna 'FechaVencimiento' a datetime, ajustar según sea necesario
            if 'FechaVencimiento' in df.columns:
                df['FechaVencimiento'] = pd.to_datetime(df['FechaVencimiento'], errors='coerce')

            wb.close()
            app.quit()

            base = os.path.basename(excel_file)
            csv_file = os.path.join(output_folder, base.replace('.xlsx', '.csv'))
            df.to_csv(csv_file, index=False, encoding='utf-8-sig', sep=';')
            print(f"Convertido {excel_file} a {csv_file}")
        except Exception as e:
            print(f"Error al convertir {excel_file} a CSV: {e}")

# Función para obtener el nombre del cliente a partir del nombre del archivo (mantenida para compatibilidad)
def obtener_nombre_cliente(filename):
    base = os.path.basename(filename)
    nombre_cliente = base.split('-')[1].strip()
    return nombre_cliente

# ========== VERIFICAR FUNCIONES AL INICIO ==========
print("=" * 60)
print("INICIANDO SISTEMA DE EXTRACCIÓN DE ANTICIPOS SCT")
print("=" * 60)
verificar_funciones_disponibles()
print("=" * 60)

# MODIFICACIÓN: Bucle principal para procesar anticipos
print("🚀 INICIANDO PROCESAMIENTO DE CLIENTES PARA ANTICIPOS")
print("📋 MODO: Extracción de Ganancias Sociedades - Período 2025 - Vencimiento Junio 2025")

# Crear directorio de salida si no existe
for ubicacion in download_list:
    if ubicacion and not os.path.exists(ubicacion):
        try:
            os.makedirs(ubicacion, exist_ok=True)
            print(f"📁 Directorio creado: {ubicacion}")
        except Exception as e:
            print(f"⚠️ Error creando directorio {ubicacion}: {e}")

indice = 0
for cuit_ingresar, cuit_representado, password, cliente, ubicacion_descarga in zip(cuit_login_list, cuit_represent_list, password_list, clientes_list, download_list):
    print(f"\n🔄 PROCESANDO CLIENTE {indice + 1}/{len(clientes_list)}")
    
    # Validar ubicación de descarga
    if not ubicacion_descarga or not os.path.exists(ubicacion_descarga):
        print(f"❌ Error: Ubicación de descarga inválida para {cliente}: {ubicacion_descarga}")
        actualizar_excel(indice, f"Ubicación descarga inválida: {ubicacion_descarga}")
        indice += 1
        continue
    
    # Procesar cliente con ubicación específica
    exito = procesar_cliente_completo(cuit_ingresar, cuit_representado, password, cliente, ubicacion_descarga, indice)
    
    if exito:
        print(f"✅ Cliente {cliente} completado exitosamente")
        print(f"📄 Excel de anticipos generado en: {ubicacion_descarga}")
    else:
        print(f"❌ Cliente {cliente} falló - ver logs para detalles")
    
    indice += 1

print("\n" + "="*60)
print("✅ PROCESAMIENTO DE TODOS LOS CLIENTES COMPLETADO")
print("📊 RESUMEN DE ANTICIPOS:")
print("   - Impuesto filtrado: Ganancias Sociedades")
print("   - Período filtrado: 2025")
print("   - Vencimiento filtrado: 01/06/2025 a 30/06/2025")
print("   - Formato de salida: Excel (.xlsx)")
print("   - Título de archivos: Anticipos - [Cliente]")
print("="*60)