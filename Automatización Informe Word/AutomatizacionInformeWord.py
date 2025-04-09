import os
import json
import time
from pathlib import Path
from bs4 import BeautifulSoup
from docxtpl import DocxTemplate

plantilla_word = "AutomatizacionW-BUENO2.docx"
carpeta_json = "json_files"
ruta_html = "html_files"

def convertir_html_a_json(ruta_entrada, carpeta_json):
    if not os.path.exists(carpeta_json):
        os.makedirs(carpeta_json)
    
    # Obtener el nombre de la carpeta de origen (nombre del servidor)
    nombre_carpeta_origen = os.path.basename(ruta_entrada)
    
    # Crear una subcarpeta dentro de json_files con el nombre de la carpeta de origen
    carpeta_destino = os.path.join(carpeta_json, nombre_carpeta_origen)
    if not os.path.exists(carpeta_destino):
        os.makedirs(carpeta_destino)
    
    for archivo in os.listdir(ruta_entrada):
        if archivo.endswith('.html'):
            ruta_archivo_entrada = os.path.join(ruta_entrada, archivo)
            nombre_json = archivo.replace('.html', '.json')
            ruta_archivo_salida = os.path.join(carpeta_destino, nombre_json)
            
            try:
                with open(ruta_archivo_entrada, 'r', encoding='utf-8') as file:
                    contenido_html = file.read()

                soup = BeautifulSoup(contenido_html, 'html.parser')
                contenido_html = soup.prettify()

                if archivo.startswith("Log") or archivo.endswith("Logs.html") or archivo.endswith("Log.html"):
                    datos_salida = procesar_logs(contenido_html)
                elif archivo.endswith("InstanciasTODO.html") or archivo.endswith("Instancias.html"):
                    datos_salida = procesar_instancias_sql(contenido_html)
                elif archivo.endswith("Sitios.html"):
                    datos_salida = procesar_sitios_web(contenido_html)
                elif archivo.startswith("IAD"):
                    datos_salida = procesar_IAD(contenido_html)
                else:
                    print(f"Tipo de archivo no reconocido: {archivo}")
                    continue

                with open(ruta_archivo_salida, 'w', encoding='utf-8') as file_json:
                    json.dump(datos_salida, file_json, indent=4, ensure_ascii=False)
                
                print(f"Convertido: {archivo} -> {os.path.basename(ruta_archivo_salida)}")
            
            except Exception as e:
                print(f"Error al procesar {archivo}: {e}")

def procesar_IAD(contenido_html):
    soup = BeautifulSoup(contenido_html, 'html.parser')
    resultado = {}
    elementos = soup.find_all(['h3', 'table'])
    
    titulo_actual = None
    for elemento in elementos:
        try:
            if elemento.name == 'h3':
                titulo_actual = elemento.text.strip()
                continue
            
            if elemento.name == 'table':
                filas = elemento.find_all('tr')
                if len(filas) <= 1:
                    continue
                
                encabezados = [th.text.strip() for th in filas[0].find_all(['th', 'td'])]
                datos_tabla = []
                
                for fila in filas[1:]:
                    columnas = fila.find_all('td')
                    if len(columnas) < len(encabezados):
                        continue
                    valores = [col.text.strip() for col in columnas[:len(encabezados)]]
                    datos_fila = dict(zip(encabezados, valores))
                    datos_tabla.append(datos_fila)
                
                if titulo_actual:
                    nombre_seccion = titulo_actual.replace(':', '').replace(' ', '')
                    if datos_tabla:
                        resultado[nombre_seccion] = datos_tabla if len(datos_tabla) > 1 else datos_tabla[0]
                
                titulo_actual = None
        
        except Exception as e:
            print(f"Error procesando elemento: {e}")
    
    return resultado

def procesar_instancias_sql(contenido_html):
    soup = BeautifulSoup(contenido_html, 'html.parser')
    tabla = soup.find_all('table')[1]
    filas = tabla.find_all('tr')[1:]
    
    instancias = []
    for fila in filas:
        columnas = fila.find_all('td')
        if len(columnas) >= 10:
            instancia = {
                'Nombre': columnas[0].text.strip(),
                'EstadoServicio': columnas[1].get('class', [''])[0] if columnas[1].get('class') else '',
                'Estado': columnas[1].text.strip(),
                'Version': columnas[2].text.strip(),
                'Edicion': columnas[3].text.strip(),
                'Collation': columnas[4].text.strip(),
                'RutaInstalacion': columnas[5].text.strip(),
                'TamañoMDF': columnas[6].text.strip(),
                'TamañoLDF': columnas[7].text.strip(),
                'TamañoBak': columnas[8].text.strip(),
                'TamañoTotal': columnas[9].text.strip()
            }
            instancias.append(instancia)
    
    return instancias

def procesar_sitios_web(contenido_html):
    soup = BeautifulSoup(contenido_html, 'html.parser')
    div = soup.find('div')
    tabla = div.find('table')
    filas = tabla.find_all('tr')[1:]
    
    sitios = []
    for fila in filas:
        columnas = fila.find_all('td')
        if len(columnas) >= 3:
            sitio = {
                'Nombre': columnas[0].text.strip(),
                'RutaFisica': columnas[1].text.strip(),
                'EspacioUtilizado': columnas[2].text.strip()
            }
            sitios.append(sitio)
    
    return sitios

def procesar_logs(archivo_html):
    soup = BeautifulSoup(archivo_html, 'html.parser')
    tabla = soup.find('table')
    if not tabla:
        raise ValueError("No se encontró la tabla en el archivo HTML.")
    
    filas = tabla.find_all('tr')[1:-4]
    detalles = []
    for fila in filas:
        columnas = fila.find_all('td')
        if len(columnas) >= 3:
            detalle = {
                'Disco': columnas[0].text.strip(),
                'TipoArchivo': columnas[1].text.strip(),
                'TamañoTotal': columnas[2].text.strip()
            }
            detalles.append(detalle)

    totales = tabla.find_all('tr')[-4:]
    resumen_totales = []
    for total in totales[:-1]:
        columnas = total.find_all('td')
        if len(columnas) >= 3:
            resumen_totales.append({
                'TipoArchivo': columnas[0].text.split(':')[0].strip().replace("Total archivos ", ""),
                'TamañoTotal': columnas[2].text.strip()
            })
    
    total_general = (
        totales[-1].find_all('td')[1].text.strip()
        if len(totales) > 0 and len(totales[-1].find_all('td')) > 1
        else "Total general no encontrado"
    )

    return {
        'Detalles': detalles,
        'Totales': resumen_totales,
        'TotalGeneral': total_general
    }

def generar_informe_word(carpeta_json, plantilla_word, nombre_salida):
    informe = DocxTemplate(plantilla_word)
    carpeta = Path(carpeta_json)
    fecha_actual = time.localtime()
    nombre = input("Ingrese su nombre: ")

    lista_instancias = []
    lista_sitios = []
    lista_logs = []
    lista_iad = []
    lista_iad1 = []
    lista_iad2 = []
    lista_iad3 = []
    lista_iad4 = []
    lista_iad5 = []
    lista_iad6 = []
    lista_iad7 = []
    lista_iad8 = []
    lista_iad9 = []
    lista_iad10 = []
    lista_iad11 = []
    lista_iad12 = []
    lista_iad13 = []


    # Procesar archivos JSON
    for archivo in carpeta.iterdir():
        if archivo.name.endswith("TODO.json"): # INFORMACIÓN INSTANCIAS
            with archivo.open("r", encoding="utf-8") as f:
                data = json.load(f)
            for i in data:
                ruta = i["RutaInstalacion"]
                partes_ruta = os.path.normpath(ruta).split(os.sep)
                ruta_contraida = os.path.join(*partes_ruta[-2:])
                insertar = {
                    "nombre": i["Nombre"],
                    "estado": i["Estado"],
                    "version": i["Version"],
                    "edicion": i["Edicion"],
                    "collation": i["Collation"],
                    "ruta": ruta_contraida,
                    "archivos_mdf": i["TamañoMDF"],
                    "archivos_ldf": i["TamañoLDF"],
                    "archivos_bak": i["TamañoBak"],
                    "total": i["TamañoTotal"],
                }
                lista_instancias.append(insertar)
        
        elif archivo.name.endswith("Sitios.json"): # INFORMACIÓN SITIOS
            with archivo.open("r") as f:
                data = json.load(f)
            for i in data:
                insertar = {
                    "nombre": i["Nombre"],
                    "ruta": i["RutaFisica"],
                    "espacio": i["EspacioUtilizado"],
                }
                lista_sitios.append(insertar)
        
        elif archivo.name.startswith('Logs') or archivo.name.endswith('Logs.json'): # INFORMACIÓN LOGS
            with archivo.open("r",encoding="utf-8") as f:
                data = json.load(f)
            for i in data["Detalles"]:
                insertar = {
                    "disco": i["Disco"],
                    "archivo": i["TipoArchivo"],
                    "total": i["TamañoTotal"],
                }
                lista_logs.append(insertar)

        elif archivo.name.startswith('IAD'): # INFORMACIÓN IAD
            with archivo.open("r") as f:
                data = json.load(f)

            # InformacionGeneral
            for i in [data["InformacionGeneral"]]:  
                dict_iad = {
                    "ubi": i["Ubicacion"],
                    "cliente": i["Cliente"],
                    "tipo": i["Tipo"],
                    "rol": i["Rol"],
                    "moni": i["Monitoreo"],
                }
                lista_iad.append(dict_iad)

            # InformaciondelSistema
            for i in [data["InformaciondelSistema"]]:  
                dict1_iad = {
                    "nombre": i["Name"],
                    "manufa": i["Manufacturer"],
                    "model": i["Model"],
                    "phys": i["Physical Processors"],
                    "totalphys": i["Total Physical Memory (Gb)"],
                    "dns": i["DnsHostName"],
                    "domain": i["Domain"],
                }
                lista_iad1.append(dict1_iad)

            # InformaciondelSistemaOperativo
            for i in [data["InformaciondelSistemaOperativo"]]:
                dict2_iad = {
                    "os": i["Operating System"],
                    "arch": i["Architecture"],
                    "version": i["Version"],
                    "install_date": i["Install Date"],
                    "windows_dir": i["WindowsDirectory"],
                }
                lista_iad2.append(dict2_iad)

            # InformacionsobreelestadoSistemaOperativo
            for i in [data["InformacionsobreelestadoSistemaOperativo"]]:
                dict3_iad = {
                    "desc": i["Description"],
                    "product_key_id": i["ProductKeyID2"],
                    "grace_period": i["GracePeriodRemaining"],
                    "license_status": i["License Status"],
                }
                lista_iad3.append(dict3_iad)

            # InformaciondelaMemoriaRAM 
            for i in data["InformaciondelaMemoriaRAM"]:
                dict4_iad = {
                    "locator": i["Device Locator"],
                    "manufa": i["Manufacturer"],
                    "speed": i["Speed"],
                    "capacity": i["Capacity (GB)"],
                }
                lista_iad4.append(dict4_iad)

            # InformaciondelosDiscos 
            for i in data["InformaciondelosDiscos"]:
                dict5_iad = {
                    "device": i["DeviceID"],
                    "fs": i["FileSystem"],
                    "vol_name": i["VolumeName"],
                    "total_size": i["Total Size (GB)"],
                    "free_space": i["Free Space (GB)"],
                }
                lista_iad5.append(dict5_iad)

            # Informaciondelosvolumenes 
            for i in data["Informaciondelosvolumenes"]:
                dict6_iad = {
                    "label": i["Label"],
                    "name": i["Name"],
                    "device_id": i["DeviceID"],
                    "system_volume": i["SystemVolume"],
                    "total_size": i["Total Size (GB)"],
                    "free_space": i["Free Space (GB)"],
                }
                lista_iad6.append(dict6_iad)

            # InformaciondelasTarjetasdeRed 
                network_data = data.get("InformaciondelasTarjetasdeRed", [])
                # Si es un diccionario único, convertirlo en lista
                if isinstance(network_data, dict):
                    network_data = [network_data]
                # Si es un string u otro tipo, convertirlo en lista con un diccionario
                elif not isinstance(network_data, list):
                    network_data = [{"Connection Name": str(network_data)}]
                
                for i in network_data:
                    dict7_iad = {
                        "conn_name": i.get("Connection Name", "N/A"),
                        "adapter_name": i.get("Adapter Name", "N/A"),
                        "type": i.get("Type", "N/A"),
                        "mac": i.get("MAC", "N/A"),
                        "enabled": i.get("Enabled", "N/A"),
                        "speed": i.get("Speed (Mbps)", "N/A"),
                        "ip": i.get("IPAddress", "N/A"),
                        "subnet": i.get("Subnet", "N/A"),
                        "gateway": i.get("Gateway", "N/A"),
                    }
                    lista_iad7.append(dict7_iad)

            # InformaciondelosUsuariosLocales 
            for i in data["InformaciondelosUsuariosLocales"]:
                dict8_iad = {
                    "name": i["Name"],
                    "status": i["Status"],
                    "disabled": i["Disabled"],
                    "description": i["Description"],
                    "local_account": i["LocalAccount"],
                }
                lista_iad8.append(dict8_iad)

            # InformaciondeHotFixinstalados 
            for i in data["InformaciondeHotFixinstalados"]:
                dict9_iad = {
                    "desc": i["Description"],
                    "hotfix_id": i["HotFixID"],
                    "installed_by": i["InstalledBy"],
                    "installed_on": i["InstalledOn"],
                }
                lista_iad9.append(dict9_iad)

            # InformaciondeVisordesucesos(Critical)
            for i in [data["InformaciondeVisordesucesos(Critical)"]]:
                dict10_iad = {
                    "system_log": i["System Log"],
                    "application_log": i["Application Log"],
                }
                lista_iad10.append(dict10_iad)

            # # InformaciondeNTP
            for i in [data["InformaciondeNTP"]]:
                dict11_iad = {
                    "leap_indicator": i["Leap Indicator"],
                    "stratum": i["Stratum"],
                    "precision": i["Precision"],
                    "root_delay": i["Root Delay"],
                    "root_dispersion": i["Root Dispersion"],
                    "ref_id": i["ReferenceId"],
                    "last_sync": i["Last Successful Sync Time"],
                    "source": i["Source"],
                    "poll_interval": i["Poll Interval"],
                }
                lista_iad11.append(dict11_iad)

            # InformaciondelosSoftwareInstalados
            for i in data["InformaciondelosSoftwareInstalados"]:
                dict12_iad = {
                    "name": i["DisplayName"],
                    "version": i["DisplayVersion"],
                    "publisher": i["Publisher"],
                }
                lista_iad12.append(dict12_iad)

            # InformaciondelosServicios
            for i in data["InformaciondelosServicios"]:
                dict13_iad = {
                    "name": i["Name"],
                    "display_name": i["DisplayName"],
                    "status": i["Status"],
                    "start_type": i["StartType"],
                }
                lista_iad13.append(dict13_iad)


    context = {
        'nombre': nombre,
        'dia': time.strftime("%d", fecha_actual),
        'mes': time.strftime("%m", fecha_actual),
        'anio': time.strftime("%Y", fecha_actual),
        'lista_sitios': lista_sitios,
        'lista_logs': lista_logs,
        'lista_instancias': lista_instancias,
        'lista_iad' : lista_iad,
        'lista_iad1' : lista_iad1,
        'lista_iad2' : lista_iad2,
        'lista_iad3' : lista_iad3,
        'lista_iad4' : lista_iad4,
        'lista_iad5' : lista_iad5,
        'lista_iad6' : lista_iad6,
        'lista_iad7' : lista_iad7,
        'lista_iad8' : lista_iad8,
        'lista_iad9' : lista_iad9,
        'lista_iad10' : lista_iad10,
        'lista_iad11' : lista_iad11,
        'lista_iad12' : lista_iad12,
        'lista_iad13' : lista_iad13,

    }

    informe.render(context)
    informe.save(nombre_salida)
    print("Documento creado exitosamente")

def main():
    while True:
        print("\nSeleccione una opción:")
        print("1. Convertir HTML a JSON")
        print("2. Generar informe Word desde JSON")
        print("3. Salir")
        
        opcion = input("\nIngrese el número deseado: ")
        
        if opcion == "1":
            if not os.path.exists(ruta_html):
                print(f"Error: La carpeta {ruta_html} no existe.")
                continue
                
            carpetas = [nombre for nombre in os.listdir(ruta_html) if os.path.isdir(os.path.join(ruta_html, nombre))]
            print(f"\nCarpetas encontradas: {carpetas}") 
            
            if carpetas:
                print("\nServidores disponibles con su HTML:")
                for i, carpeta in enumerate(carpetas, start=1):
                    print(f"{i}. {carpeta}")

                try:
                    seleccion = int(input("\nSeleccione un servidor: "))
                    if 1 <= seleccion <= len(carpetas):
                        nombre_carpeta = carpetas[seleccion - 1]
                        ruta_entrada = os.path.join(ruta_html, nombre_carpeta)
                        convertir_html_a_json(ruta_entrada, carpeta_json)
                    else:
                        print("\nNúmero no encontrado.")
                except ValueError:
                    print("\nEntrada no válida.")
            else:
                print(f"\nNo se encontraron carpetas en {ruta_html}")

        elif opcion == "2":
            if not os.path.exists(carpeta_json):
                print(f"Error: La carpeta {carpeta_json} no existe.")
                continue
                
            servidores = [nombre for nombre in os.listdir(carpeta_json) 
                        if os.path.isdir(os.path.join(carpeta_json, nombre))]
            
            if not servidores:
                print("\nNo hay servidores procesados. Primero convierte HTML a JSON.")
                continue
                
            print("\nServidores disponibles:")
            for i, servidor in enumerate(servidores, 1):
                print(f"{i}. {servidor}")
                
            try:
                seleccion = int(input("\nSeleccione un servidor: "))
                if 1 <= seleccion <= len(servidores):
                    servidor_seleccionado = servidores[seleccion - 1]
                    ruta_servidor = os.path.join(carpeta_json, servidor_seleccionado)
                    
                    # Generar informe solo para el servidor seleccionado
                    nombre_salida = input("\nIngrese nombre para el informe (ej: Informe_Servidor1.docx): ")
                    generar_informe_word(ruta_servidor, plantilla_word, nombre_salida)
                else:
                    print("\nNúmero fuera de rango")
            except ValueError:
                print("Entrada inválida. Debe ser un número.")
        
        elif opcion == "3":
            print("\nSaliendo del script")
            break
        
        else:
            print("Opción no válida. Por favor, intente nuevamente.")

if __name__ == "__main__":
    main()