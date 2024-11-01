from cryptography.fernet import Fernet
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
        print(self.datos)

if __name__ == "__main__":
    proyecto = TestProject("Credentials/CredencialesGS.json", "Hoja_Prueba")
    proyecto.mostrar_datos()
