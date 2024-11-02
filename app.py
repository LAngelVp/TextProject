import sys
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import mplcursors
import gspread
from google.oauth2.service_account import Credentials
from PyQt5.QtWidgets import *
from Test_Project import Ui_Formulario

class TestProject(QWidget):
    def __init__(self):
        super().__init__()
        self.ui = Ui_Formulario()
        self.ui.setupUi(self)
        self.credenciales_path = "Credentials/CredencialesGS.json"
        self.hoja_nombre = "Hoja_Prueba"
        self.cliente = None
        self.datos = None

        self.ui.btn_ejecutar.clicked.connect(self.ejecutar)
    def mostrar_mensaje(titulo, mensaje, tipo=QMessageBox.Information):
        cuadro_mensaje = QMessageBox()
        cuadro_mensaje.setWindowTitle(titulo)
        cuadro_mensaje.setIcon(tipo)
        cuadro_mensaje.setText(mensaje)
        cuadro_mensaje.exec_()

    def ejecutar(self):
        self.cliente = self.autenticar()
        if self.cliente is not None:
            # Intentar obtener datos de la hoja de cálculo
            self.datos = self.obtener_datos()
            if self.datos is not None:
                # Procesar los datos si se obtuvieron correctamente
                self.trabajar_datos()
            else:
                # Mostrar mensaje de alerta si no se encontraron datos
                self.mostrar_mensaje("Alerta", "No se encontraron datos en la hoja de cálculo.", QMessageBox.Warning)
        else:
            # Mostrar mensaje de alerta si la autenticación falló
            self.mostrar_mensaje("Error de Autenticación", "No se pudo autenticar el acceso a la hoja de cálculo.", QMessageBox.Critical)



    
#// FUNCION PARA AUTENTICARA GOOGLE SHEETS
    def autenticar(self):
        try:
            # COMMENT: PERMISOS PARA GOOGLE SHEETS Y GOOGLE DRIVE
            permisos = [
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"
            ]
            #COMMENT: CARGAR CREDENCIALES DESDE EL DOCUMENTO JSON
            credenciales = Credentials.from_service_account_file(self.credenciales_path, scopes=permisos)
            return gspread.authorize(credenciales) # Autenticamos la aplicacion con google sheets.
        except FileNotFoundError:
            print(f"Error: El archivo de credenciales no se encontró en la ruta: {self.credenciales_path}")
            raise
        except Exception as e:
            print(f"Error al autenticar: {e}")
            raise
#// FUNCION PARA OBTENER LOS DATOS
    def obtener_datos(self):
        try:
            #COMMENT: UTILIZAMOS EL CLIENTE AUTENTICADO PARA ACCEDER A LA HOJA DE SHEETS
                # USAMOS self.cliente QUE ES EL CLIENTE AUTENTICADO
                # USAMOR self.hoja_nombre QUE ES LA HOJA A LA QUE VAMOS A ACCEDER
                # USAMOS sheet1 QUE ES LA PRIMERA PESTAÑA DEL LIBRO
            hoja = self.cliente.open(self.hoja_nombre).sheet1
            #COMMENT: OBTENEMOS TODO EL CONTENIDO
            return hoja.get_all_records()
        except gspread.SpreadsheetNotFound:
            print(f"Error: No se encontró la hoja de cálculo '{self.hoja_nombre}'")
            raise
        except Exception as e:
            print(f"Error al obtener datos: {e}")
            raise
#// FUNCION PARA TRABAJAR CON LOS DATOS
    def trabajar_datos(self):
        fecha_actual = datetime.now() # Obtenemos la fecha actual.
        conjunto_datos = pd.DataFrame(self.datos)  # Convertimos los datos a un DataFrame.
        conjunto_datos["Salario Anual"] = conjunto_datos["Salario Mensual"] * 12   # Calculamos el salario anual.
        conjunto_datos["Fecha de Contratacion"] = pd.to_datetime(conjunto_datos["Fecha de Contratacion"], dayfirst=True) # Formateamos la fecha de contratación a tipo fecha, inidicando que se recibira el dia primero antes del mes.
        conjunto_datos["Años en la Empresa"] = conjunto_datos.apply(lambda fila: relativedelta(fecha_actual, fila["Fecha de Contratacion"]).years, axis=1) # Ejecutamos una funcion anonima por cada registro (fila) calculando la antigüedad. Relativedelta calcula la fecha exacta.
        conjunto_datos["Fecha de Contratacion"] = conjunto_datos["Fecha de Contratacion"].dt.strftime('%d/%m/%Y') # Convertimos las fechas a tipo String.
        # Creamos figura y ejes
        figura, variable_ejes = plt.subplots(figsize=(10, 6))  # Tamaño de la figura.
        
        # En esta parte, graficamos los empleados.
        barras = variable_ejes.bar(conjunto_datos["Nombre"], conjunto_datos["Salario Anual"], color='skyblue')  
        
        # Añadimos etiquetas y título.
        variable_ejes.set_xlabel("Empleados")  
        variable_ejes.set_ylabel("Salario Anual")  
        variable_ejes.set_title("Salario Anual por Empleado")  
        
        # Roto las etiquetas del eje X a 90 grados
        variable_ejes.set_xticklabels(conjunto_datos["Nombre"], rotation=90, ha='center')
        
        # Ajusto el layout para que no se corten las etiquetas.
        plt.tight_layout()

        # Añadir interactividad con mplcursors para ver detalles
        mplcursors.cursor(barras, hover=True).connect("add", lambda sel: sel.annotation.set_text(
            f"Empleado: {conjunto_datos['Nombre'][sel.index]}\nAños en la Empresa: {conjunto_datos['Años en la Empresa'][sel.index]} Años\nSalario Anual: {conjunto_datos['Salario Anual'][sel.index]:,.2f}"
        ))
        
        # Creo un canvas para el gráfico dentro de la misma ventana
        canvas = FigureCanvas(figura)
        layout = QVBoxLayout(self.ui.contenedor_grafico)  # Creo un layout dentro del QFrame que contendra la grafica
        self.ui.contenedor_grafico.setLayout(layout)  # Establezco el layout en el QFrame

        # Agrego el canvas al layout
        layout.addWidget(canvas)

        self.cargar_datos(conjunto_datos) # Finalmente agrego el dataframe a google sheets

#// FUNCION PARA CARGAR LOS DATOS A GOOGLE SHEETS
    def cargar_datos(self, conjunto_datos):
        # Intentar acceder a la hoja "Salarios Anuales"
        hoja_existente = None
        try:
            hoja_existente = self.cliente.open(self.hoja_nombre).worksheet("Reporte de Salarios")
        except gspread.WorksheetNotFound:
            hoja_existente = None  # Asignamos None si no existe la hoja
        # Si la hoja existe en el libro, la elimina
        if hoja_existente:
            self.cliente.open(self.hoja_nombre).del_worksheet(hoja_existente)
        # Creamos  una nueva hoja con el nombre "Reporte de Salarios"
            # Abrimos la hoja.
            # Añadimos una nueva pestaña (tittle = le damos un nombre, rows = establece el numero de filas + 1 (encabezados), cols = el numero de columnas)
        hoja_resultado = self.cliente.open(self.hoja_nombre).add_worksheet(title="Reporte de Salarios", rows=str(len(conjunto_datos) + 1), cols=str(len(conjunto_datos.columns)))

        # Obtenemos los nombres de las columnas.
        # conjunto_datos.columns.values.tolist convierte los nombres en listas para crear la fila. (para encabezados)
        # conjunto_datos.values.tolist obtiene los valores del dataframe sin los encabezados, convierte los valores en listas, donde cada sublista creada sera una fila en la sheet. (para contenido)
        hoja_resultado.update([conjunto_datos.columns.values.tolist()] + conjunto_datos.values.tolist())

if __name__ == "__main__":
    app = QApplication(sys.argv)
    ventana_principal = TestProject()
    ventana_principal.show()
    sys.exit(app.exec())
