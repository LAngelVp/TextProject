import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
import gspread
from google.oauth2.service_account import Credentials

class TestProject:
    def __init__(self, credenciales_path, hoja_nombre):
        self.credenciales_path = credenciales_path
        self.hoja_nombre = hoja_nombre
        self.cliente = self.autenticar()
        self.datos = self.obtener_datos()

    def autenticar(self):
        try:
            scopes = [
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"
            ]
            credenciales = Credentials.from_service_account_file(self.credenciales_path, scopes=scopes)
            return gspread.authorize(credenciales)
        except FileNotFoundError:
            print(f"Error: El archivo de credenciales no se encontró en la ruta: {self.credenciales_path}")
            raise
        except Exception as e:
            print(f"Error al autenticar: {e}")
            raise

    def obtener_datos(self):
        try:
            hoja = self.cliente.open(self.hoja_nombre).sheet1
            return hoja.get_all_records()
        except gspread.SpreadsheetNotFound:
            print(f"Error: No se encontró la hoja de cálculo '{self.hoja_nombre}'")
            raise
        except Exception as e:
            print(f"Error al obtener datos: {e}")
            raise

    def mostrar_datos(self):
        fecha_actual = datetime.now()
        conjunto_datos = pd.DataFrame(self.datos)
        conjunto_datos["Salario Anual"] = conjunto_datos["Salario Mensual"] * 12
        
        # Convertir la fecha de contratación
        conjunto_datos["Fecha de Contratacion"] = pd.to_datetime(conjunto_datos["Fecha de Contratacion"], dayfirst=True)
        
        # Calcular años de antigüedad
        conjunto_datos["Años en la Empresa"] = conjunto_datos.apply(
            lambda fila: relativedelta(fecha_actual, fila["Fecha de Contratacion"]).years, axis=1
        )
        
        # Convertir las fechas a formato de cadena
        conjunto_datos["Fecha de Contratacion"] = conjunto_datos["Fecha de Contratacion"].dt.strftime('%d/%m/%Y')

        # Crear una nueva hoja en la misma hoja de cálculo
        nueva_hoja = self.cliente.open(self.hoja_nombre).add_worksheet(title="Resumen Antigüedad", rows=str(len(conjunto_datos) + 1), cols=str(len(conjunto_datos.columns)))
    def cargar_datos(self, conjunto_datos):
        # Intentar acceder a la hoja "Salarios Anuales"
        hoja_existente = None
        try:
            hoja_existente = self.cliente.open(self.hoja_nombre).worksheet("Reporte de Salarios")
        except gspread.WorksheetNotFound:
            hoja_existente = None  # La hoja no existe

        # Si la hoja existe, eliminarla
        if hoja_existente:
            self.cliente.open(self.hoja_nombre).del_worksheet(hoja_existente)

        # Crear una nueva hoja
        hoja_resultado = self.cliente.open(self.hoja_nombre).add_worksheet(title="Reporte de Salarios", rows=str(len(conjunto_datos) + 1), cols=str(len(conjunto_datos.columns)))

        # Subir los datos a la nueva hoja
        hoja_resultado.update([conjunto_datos.columns.values.tolist()] + conjunto_datos.values.tolist())

if __name__ == "__main__":
    proyecto = TestProject("Credentials/CredencialesGS.json", "Hoja_Prueba")
    proyecto.mostrar_datos()
