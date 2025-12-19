import os
import shutil
import ctypes

# Ruta de descargas (ajustar si es necesario)
ruta_origen = r"C:\Users\Usuario\Desktop\ARCHIVOS BRANCO"


# Definimos las carpetas destino
carpeta_facturas = os.path.join(ruta_origen, "Facturas")
carpeta_transferencias = os.path.join(ruta_origen, "Transferencias de BCP")
carpeta_precomprobantes = os.path.join(ruta_origen, "Precomprobantes")
carpeta_pagos = os.path.join(ruta_origen, "Pagos a cuenta")
carpeta_remitos = os.path.join(ruta_origen, "Remitos")
carpeta_excel = os.path.join(ruta_origen, "Excel")

def organizar_archivos():
    # Creamos las carpetas si no existen
    for carpeta in [carpeta_facturas, carpeta_transferencias, carpeta_precomprobantes, carpeta_pagos, carpeta_remitos, carpeta_excel]:
        if not os.path.exists(carpeta):
            os.makedirs(carpeta)

    # Contadores para saber qué hizo el script al final
    movidos = 0

    print("Iniciando organización...")

    for archivo in os.listdir(ruta_origen):
        ruta_completa = os.path.join(ruta_origen, archivo)

        # IMPORTANTE: Ignorar si es una carpeta (para no mover carpetas dentro de carpetas)
        if os.path.isdir(ruta_completa):
            continue
        
        # Ignorar archivos temporales o de sistema (opcional pero recomendado)
        if archivo.startswith(".") or archivo == "desktop.ini":
            continue

        # Convertimos a minúsculas para comparar fácil (fc, Fc, FC da igual)
        nombre_lower = archivo.lower()
        destino = None

        # --- REGLAS DEFINIDAS ---
        
        # 0. EXCEL (Prioridad si queres agrupar todos los excel juntos)
        if nombre_lower.endswith(".xls") or nombre_lower.endswith(".xlsx"):
            destino = carpeta_excel

        # 1. Empieza con Fc o Fact -> Facturas
        elif nombre_lower.startswith("fc") or nombre_lower.startswith("fact"):
            destino = carpeta_facturas
            
        # 2. Empieza con Transferencia -> Transferencias
        elif nombre_lower.startswith("transferencia"):
            destino = carpeta_transferencias
            
        # 3. Empieza con Precomprobante -> Precomprobante
        elif nombre_lower.startswith("precomprobante"):
            destino = carpeta_precomprobantes

        # 4. Remitos (Empieza con RTO o REM)
        elif nombre_lower.startswith("rto") or nombre_lower.startswith("rem"):
            destino = carpeta_remitos
            
        # 5. Pagos a cuenta (Mínimo 3 guiones bajos '_')
        elif nombre_lower.count("-") >= 3:
            destino = carpeta_pagos
        
        # 6. TODO LO DEMÁS -> SE IGNORA (No se mueve)
        else:
            destino = None

        # --- MOVER EL ARCHIVO ---
        if destino:
            try:
                # Verificamos si ya existe el archivo en el destino para no sobrescribir
                ruta_final = os.path.join(destino, archivo)
                if os.path.exists(ruta_final):
                    # Si existe, le agregamos un prefijo "Copia_" para no perder nada
                    nuevo_nombre = f"Copia_{archivo}"
                    ruta_final = os.path.join(destino, nuevo_nombre)
                
                shutil.move(ruta_completa, ruta_final)
                print(f"Movido: {archivo}  --->  {os.path.basename(destino)}")
                movidos += 1
            except Exception as e:
                print(f"Error moviendo {archivo}: {e}")

    print("------------------------------------------------")
    print(f"Listo. Se han organizado {movidos} archivos.")
    
    # Pop-up notification
    try:
        ctypes.windll.user32.MessageBoxW(0, f"Se han organizado {movidos} archivos.", "Organización Completa", 0x40 | 0x1)
    except Exception:
        pass

if __name__ == "__main__":
    organizar_archivos()