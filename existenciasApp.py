import smtplib
import xmlrpc.client
import re
import os
import time 

import requests
import urllib.parse as urlparse
from urllib.parse import parse_qs, parse_qsl, unquote
from urllib.parse import urlparse, unquote
import smtplib
from email.message import EmailMessage
from datetime import datetime
import pdfplumber
import numpy as np

import pandas as pd 
from pandas._libs import index
from pandas.core.ops import invalid
from tqdm import tqdm

class existenciasOdoo():
    def __init__(self):
        # Pandas, analisis de datos 
        self.odooURL = ""
        self.odooDB = ""
        self.odooUser = ""
        self.odooPass = ""

        self.parkerPath = os.path.join("excel/Parker.xlsx")
        self.invCeaPath = os.path.join("excel/InventarioCEA.xlsx")
        self.stockPhoenixPath = os.path.join("excel/Phoenix.xlsx")
        self.optexPath = os.path.join("excel/existenciasOptex.xlsx")
        self.optexPDFPath = os.path.join("PDF/optex.pdf")
        self.pilzPath = os.path.join("excel/Pilz.xlsx")
    
        self.finder = pd.read_csv(os.path.join("excel/Finder.csv"), sep=";")
        self.finderPath = os.path.join("excel/Finder.xlsx")
        self.eatonPath = os.path.join("excel/Eaton.xlsx")
        self.eatonPDFPath = os.path.join("PDF/Eaton.pdf")

        self.parker = pd.read_excel(self.parkerPath)
        self.finder.to_excel(self.finderPath, index=False)  
        self.invCEA = pd.read_excel(self.invCeaPath, skiprows=6)
        self.stockPhoenix = pd.read_excel(self.stockPhoenixPath) 
        self.pilz = pd.read_excel(self.pilzPath)

        common = xmlrpc.client.ServerProxy(f"{self.odooURL}/xmlrpc/2/common")
        self.uid = common.authenticate(self.odooDB, self.odooUser, self.odooPass, {})
        self.models = xmlrpc.client.ServerProxy(f"{self.odooURL}/xmlrpc/2/object")
        #Convaertir de PDF a excel
        self.optex = self.pdfToExcel(pdfPath=self.optexPDFPath, excelPath= self.optexPath)
        self.eaton = self.pdfToExcel(pdfPath= self.eatonPDFPath, excelPath= self.eatonPath)
        #Para enviar correo 
        self.mailOrigen = "ceacontrol1405@gmail.com"
        self.mailDestinos = ["woopaco@gmail.com","apoyo.direccion@ceacontrol.com", "luis.vales@ceacontrol.com"]
        self.mailContraseña = "klsf sois xzpm gsyw"
        #Para creacion del excel
        self.activationData = {
        'ID en odoo':[],
        'Numero de parte':[]
        }

        
        
        #Para guardar las rutas de los archivos validos 
        self.arrPathsCEA = []
        # Calcular el total de tuplas de todo el inventario
        self.totalDF = 0
        self.countDF = 0      
        
        with open('ServiceList.txt', 'w', encoding='UTF-8') as f:
            f.write("-#-#-#-#-#-#-#-#-#-#-Lista de supuestos servicios en las listas de existencias-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#\n\n")
        


    def compProducts(self, column_code,location_ID, cuantity_code, df, filePath):
        
        df = df.dropna(how= 'all')
        
        mask = df.notna().sum(axis=1) == 1
        df = df[~mask]
        df = df[df.iloc[:,cuantity_code] != 0]
        df = df[df.iloc[:,cuantity_code] != '-']

        df.to_excel(filePath, index= False)
        
        
        self.countDF += 1
        validCodes = []
        invalidCodes = [] 
        productsIDs = []
        count = 0
        countVlid = 0
        countNonValid = 0
        prodCad = ""
        
        stockMessage = ""
        bar = tqdm(total= len(df), desc=f"Analizando {filePath}")
        
        for index, row  in df.iterrows():
            #Buscoando en Odoo si existe el producto on ese codigo    
            codigo = str(row.iloc[column_code]).strip()
            cantidad = str(row.iloc[cuantity_code]).strip()
            if count % 100 == 0 and count != 0:
                time.sleep(2)
            try:
                producto = self.models.execute_kw(
                    self.odooDB,
                    self.uid,
                    self.odooPass,
                    'product.product',
                    'search_read',
                    [[['name', 'ilike', codigo]]],
                    {'fields':['id','name','default_code', 'type'],'limit':1}
                    )
                count += 1
                0
                if os.name == 'nt':
                    os.system('cls')    
                else:                         
                    os.system('clear')
        
            #print(f"Analizando {filePath}...")
                
                bar.update(1)
                print(f"Productos analizados: {count}")
                print(f"Productos encontrados: {countVlid}")
                print(f"Productos no encontrados: {countNonValid}")
                print(prodCad)
                print(stockMessage)
                print(f"Archivos Analizados: {self.countDF}/{self.totalDF}")
                if producto:
                    id = producto[0]['id']
                    productsIDs.append(id)
                    validCodes.append(True)
                    invalidCodes.append(False)
                    countVlid += 1
                    stockMessage = self.createNewStockCuant(id, location_ID , cantidad, codigo)     
                    

            
                else:
                    validCodes.append(False)
                    invalidCodes.append(True)
                    countNonValid += 1
                    if filePath in [self.invCeaPath, self.parkerPath]:
                        descripcion = row.iloc[2]
                        descuento = '50%'
                        marca = 'Parker  Hannifin'
                    if filePath == self.pilzPath:
                        descripcion = row.iloc[1]
                        descuento = '25%'
                        marca = 'PILZ'
                    if filePath == self.finderPath:
                        descripcion = row.iloc[1]
                        descuento = '40%'
                        marca = 'FINDER'
                    if filePath == self.pilzPath:
                        descripcion = row.iloc[1]
                        descuento = '25%'
                        marca = 'PILZ'
                    if filePath == self.stockPhoenixPath:
                        descripcion = row.iloc[2]
                        descuento = '38.26%'
                        marca = 'PHOENIX CONTACT'
                    if filePath == self.eatonPath:
                        descripcion = row.iloc[1]
                        descuento = '65%'
                        marca = 'MOELLER ELECTRIC'
                    if filePath == self.optexPath:
                        descripcion = row.iloc[1]
                        descuento = '35%'
                        marca = 'PILZ'

                    productID = self.models.execute_kw(
                        self.odooDB,
                        self.uid,
                        self.odooPass,
                        'product.product',
                        'create',
                        [{
                            'name':codigo,
                            'default_code':descripcion,
                            'x_studio_marca_1':marca,
                            'x_studio_descuento_':descuento,
                            'x_studio_fecha_de_inicio': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                            'list_price':0,
                            'standard_price':0,
                            'categ_id':16,
                            'tracking':'none',
                            'type': 'consu'

                        }])
                    prodCad = f"[CREATE]: Nuevo producto creado para {codigo}"
                    self.activationData['ID en odoo'].append(productID)
                    self.activationData['Numero de parte'].append(codigo)

            except xmlrpc.client.ProtocolError as e:
                print(f"Error de protocolo! {e.url} codigo{e.errcode}")
                print(f"Mensaje de error: {e.errmsg}") 
            except xmlrpc.client.Fault as e:
                print(f"Error XMLRPC {e.faultString}")
                print(f"Codigo {e.faultCode}")
            except ConnectionRefusedError:
                print("Conexion Rechazada por el servidor")
            except TimeoutError:
                print(f"Tiempo de espera agotado para conectarse con el servidor")

            #filtrar solo los productos que se encuentran en la base de datos    
        df_validCodes = df[validCodes]
        df_invalidCodes = df[invalidCodes]

        #Guardar la ruta del excel de los no validos
        inValidPath = os.path.join("invalidCodes",f"invalidCodes_{filePath.replace("excel/","")}.xlsx")
        self.arrPathsCEA.append(inValidPath)
        #Guardar el contenido 
        df_validCodes.to_excel(os.path.join("validCodes", f"validCodes_{filePath.replace("excel/","")}.xlsx"), index=False)
        df_invalidCodes.to_excel(inValidPath, index=False)      
        print("Actualizando cache del servidor...")
        self.updateCache(productsIDs)
        # 3. Validar inventario (aplicar)
        
        if os.name == 'nt':
            os.system('cls')    
        else:                         
            os.system('clear')
        
        print("Proceso terminado")
    
    
    def createNewStockCuant(self, product_id, location_id, quantity, codigo):
        """
        Creates or updates a stock quant (inventory record) for a given product in Odoo.

        Parameters:
            product_id (int): The ID of the product to update or create the stock quant for.
            location_id (int): The ID of the location where the stock is stored.
            quantity (int or float): The quantity of the product to set in stock.

        Returns:
            str: A message indicating whether the stock quant was created or updated, or None if an error occurred.
        """
        try:
        # 1) Verificar si ya existe un stock.quant para ese producto en esa ubicación
            quant = self.models.execute_kw(
                self.odooDB, self.uid, self.odooPass, 
                'stock.quant', 
                'search_read',
                [[['product_id', '=', product_id], ['location_id', '=', location_id]]],
                {'fields': ['id', 'quantity']}
                )

            if quant:
                quant_id = quant[0]['id']

            # 2) Ajuste de inventario correctamente aplicado (se refleja en qty_available)
                self.models.execute_kw(
                    self.odooDB, self.uid, self.odooPass,
                    'stock.quant', 'write',
                    [[quant_id], {
                    'inventory_quantity': quantity,
                    'inventory_quantity_set': True,
                    'inventory_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    
                }])
                
                return f"[UPDATE] {product_id}: Existencia actualizada a {quantity}"

            else:
                # 3) Crear un nuevo quant en la ubicación correcta con cantidad aplicada
                self.models.execute_kw(
                    self.odooDB, 
                    self.uid, 
                    self.odooPass,
                    'stock.quant', 
                    'create',
                    [{
                        'product_id': product_id,
                        'location_id': location_id,
                        'inventory_quantity': quantity,
                        'inventory_quantity_set': True
                        }]
                        )
                
                return f"[CREATE] {product_id}: Nuevo stock.quant creado con {quantity}"
            
            
        except Exception as e:
        # 4) Registrar errores limpios
            with open('StockErrors.log', 'a', encoding='UTF-8') as f:
                f.write(f"Producto {codigo} -> {str(e)}\n{quantity}")
            self.activationData['ID en odoo'].append(product_id)
            self.activationData['Numero de parte'].append(codigo)
            return f"[ERROR] {product_id}: Revisar StockErrors.log {str(e)}"
            
        
            
                        
    def pdfToExcel(self, pdfPath, excelPath):
        
        data = ""
        procData = []
        rows = []

        if pdfPath == self.optexPDFPath:
            #Extrar una cadena de texto desde el pdf
            print(f"Analizando PDF: {pdfPath}")
            data = ""
            print(f"Analizando PDF: {pdfPath}")

            with pdfplumber.open(pdfPath) as pdf:
                for pagina in pdf.pages:
                    texto = pagina.extract_text()
                    if texto:
                        data += texto + "\n"

            if data == "":
                print("No se pudo extraer texto del PDF")
                return None

            # Eliminar encabezado
            lineas = data.split("\n")
            lineas = [l for l in lineas if not l.startswith("SKU DESCRIPCION LINEA EXISTENCIAS") and l.strip() != ""]

            for linea in lineas:
                partes = linea.split()

                if len(partes) < 3:
                    continue  # Saltar líneas dañadas

                sku = partes[0]
                existencias = partes[-1]

                # La palabra justo antes de existencias es la línea (OPTEX)
                linea_producto = partes[-2]

                # Todo lo demás (entre sku y linea_producto) es descripción
                descripcion = " ".join(partes[1:-2])

                rows.append([sku, descripcion, linea_producto, existencias])

            df = pd.DataFrame(rows, columns=["SKU", "Descripcion", "Linea", "Existencias"])
            df.to_excel(excelPath, index=False)

            print(f"PDF convertido correctamente a {excelPath}")

            if os.name == 'nt':
                os.system('cls')
            else:
                os.system('clear')

            return df
            #Crear el excel desde el texto
        else:    
            return self.pdfToExcelEaton(pdfPath, excelPath)
    def updateCache(self, product_ids):
        try:
            self.models.execute_kw(
                self.odooDB,
                self.uid,
                self.odooPass,
                'product.product',
                'write',
                [product_ids, {}]
            )
    
        except xmlrpc.client.ProtocolError as e:
                print(f"Error de protocolo! {e.url} codigo{e.errcode}")
                print(f"Mensaje de error: {e.errmsg}") 
        except xmlrpc.client.Fault as e:
                print(f"Error XMLRPC {e.faultString}")
                print(f"Codigo {e.faultCode}")
        except ConnectionRefusedError:
                print("Conexion Rechazada por el servidor")
        except TimeoutError:
                print(f"Tiempo de espera agotado para conectarse con el servidor")
        
    def deleteStockQuants(self, arrLoc):
        quants = []
        try:
            for a in arrLoc:
                quants += self.models.execute_kw(
                    self.odooDB,
                    self.uid,
                    self.odooPass,
                    'stock.quant',
                    'search',
                    [[('location_id', '=', a)]]
                    )
                
            print("Eliminando existencias...")
            if quants:
                self.models.execute_kw(
                    self.odooDB,
                    self.uid,
                    self.odooPass,
                    'stock.quant',
                    'write',
                    [quants, {'quantity': 0}]
                )
                print(f"Todadas las existencias han sido eliminadas")
                        
            else:
                print("No hay existencias que eliminar")
        except xmlrpc.client.ProtocolError as e:
                print(f"Error de protocolo! {e.url} codigo{e.errcode}")
                print(f"Mensaje de error: {e.errmsg}") 
        except xmlrpc.client.Fault as e:
                print(f"Error XMLRPC {e.faultString}")
                print(f"Codigo {e.faultCode}")
        except ConnectionRefusedError:
                print("Conexion Rechazada por el servidor")
        except TimeoutError:
                print(f"Tiempo de espera agotado para conectarse con el servidor")
        
        
    def pdfToExcelEaton(self,pdfPath, excelPath):
        print(f"Analizando PDF: {pdfPath}")


        dataframes = []

        try:
            with pdfplumber.open(pdfPath) as pdf:
                for i, page in enumerate(pdf.pages):
                    # Extrae tabla de cada página (si existe)
                    table = page.extract_table()
                    if table:
                        df = pd.DataFrame(table[1:], columns=table[0])
                        dataframes.append(df)
        except Exception as e:
            print(f"⚠️ Error leyendo el PDF: {e}")
            return pd.DataFrame()

        # Si no se encontraron tablas
        if not dataframes:
            print("⚠️ No se detectaron tablas en el PDF.")
            return pd.DataFrame()

        # Combinar todas las páginas en un solo DataFrame
        df_final = pd.concat(dataframes, ignore_index=True)

        # Guardar en Excel si se proporcionó la ruta
        if excelPath:
            df_final.to_excel(excelPath, index=False)
            print(f"PDF convertido correctamente a {excelPath}")
        if os.name == 'nt':
            os.system('cls')    
        else:                         
            os.system('clear')

        return df_final
    
    def sendEmail(self, listaRutas):
        mensaje = EmailMessage()
        mensaje['From'] = self.mailOrigen
        mensaje["To"] = ",".join(self.mailDestinos)
        mensaje['Subject'] = "PRODUCTOS DE LAS LISTAS DE EXISTENCIA QUE NO SE ENCONTRARON EN LA BASE DE DATOS"
        mensaje.set_content(f"""Hola, muy buenas tardes, espero y este correo los encuentre con bien.
Adjunto los siguientes archivos:
""")
        for ruta in listaRutas:
            with open(ruta, 'rb') as f:
                datos = f.read()
                nombre = os.path.basename(ruta)
                mensaje.add_attachment(
                    datos,
                    maintype="application",
                    subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    filename=nombre )
                try:
                    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                        smtp.login(self.mailOrigen, self.mailContraseña)
                        smtp.send_message(mensaje)
                    print("Correo enviado correctamente a todos los destinatarios.")
                except Exception as e:
                    print("Error al enviar el correo:", e)
    


if __name__ == "__main__":
    oExistencias = existenciasOdoo()
    """
    Locaciones por ID segun provedor

| Proveedor / Almacén       | ID  | Ubicación (`complete_name`)    
| ------------------------- | --- | ---------------------------------- |
| **CEA**                   | 252 | WH/Inventario CEA                  |
| **pilz**                  | 266 | pilz/Existencias                   |
| **PHOE (Phoenix)**        | 279 | PHOE/Existencias                   |
| **Eaton**                 | 285 | Eaton/Existencias                  |
| **P H (Parker Hannifin)** | 223 | P H/Existencias                    |
| **PAT (Patlite)**         | 235 | PAT/Existencias                    |
| **FINDE (Finder)**        | 241 | FINDE/Existencias                  |
| **OPTEX**                 | 247 | OPTEX/Existencias                  |


    """
    
    cea = False
    pilz = False
    eaton = False 
    parker = False 
    finder = False
    phinix = False
    optex = False 

    sendEmail = False
    
    arrLoc = []

    
    
    inp = input(f"""Escribe una secuencia de numero de acuerdo a los inventarios que deseas actualizar,
Despues, presiona ENTER:

[1]Inventario CEA (SIMAN)
[2]Pilz
[3]Eaton
[4]Parker
[5]Finder
[6]Phoenix
[7]Optex
""")
    for i in inp:
        if int(i) == 1:
            cea = True
            arrLoc.append(252)
            oExistencias.totalDF += 1
        elif int(i) == 2:
            pilz = True
            arrLoc.append(266)
            oExistencias.totalDF += 1
        elif int(i) == 3:
            eaton = True
            arrLoc.append(285)
            oExistencias.totalDF += 1
        elif int(i) == 4:
            parker = True
            arrLoc.append(223)
            oExistencias.totalDF += 1
        elif int(i) == 5:
            finder = True
            arrLoc.append(241)
            oExistencias.totalDF += 1
        elif int(i) == 6:
            phinix = True
            arrLoc.append(279)
            oExistencias.totalDF += 1
        elif int(i) == 7:
            optex = True
            arrLoc.append(247)
            oExistencias.totalDF += 1

    inp1 = input(f"""Se creara un excel por cada existencia donde no se encuentren productos en la base de datos.
Cada excel contiene el SKU o Numero de Parte de los productos en cuention.
¿Deasea enviar estos reportes por correo?
[s] si
[N] no 

""")
    if inp1.lower() == 's':
        email = True
    else:
        email = False


    #Aqui empieza el proceso de eliminacion y actualizacion de existencias
    if len(arrLoc) > 0:
        oExistencias.deleteStockQuants(arrLoc=arrLoc) 
    if cea:
        oExistencias.compProducts(0,252, 8,oExistencias.invCEA, oExistencias.invCeaPath )
    if pilz:
        oExistencias.compProducts(0,266,2, oExistencias.pilz, oExistencias.pilzPath)
    if eaton:
        oExistencias.compProducts(0,285, 2, oExistencias.eaton , oExistencias.eatonPath)
    if parker:
        oExistencias.compProducts(0, 223, 5, oExistencias.parker,oExistencias.parkerPath)
    if finder:
        oExistencias.compProducts(0,241,3, oExistencias.finder, oExistencias.finderPath)
    if phinix:
        oExistencias.compProducts(1,279, 5, oExistencias.stockPhoenix, oExistencias.stockPhoenixPath)
    if optex:
        oExistencias.compProducts(0,247,1, oExistencias.optex, oExistencias.optexPath) 
    if email:
        oExistencias.sendEmail(oExistencias.arrPathsCEA)
    print('Creando excel para activacion de nuevos productos')
    df = pd.DataFrame(oExistencias.activationData)
    df.to_excel(os.path.join('productos_para_activar.xlsx'), index=False)
    input('Proceso terminado. pulse cualquier tecla')
    
    

