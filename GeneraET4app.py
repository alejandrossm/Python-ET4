#importar openpyxl: pip install openpyxl
import openpyxl
#importar pandas: pip install pandas
import pandas as pd

# Carga un archivo CSV con los datos
resultados = pd.read_csv("carga.csv", sep=";", header = 0, index_col = False)
resultados

# Setea las celdas en donde irán los datos en la planilla 
# (esto quizás debe quedar en un archivo de configuración)
dict_celdas = {
  'nombre' : 'G8',
  'rut' : 'G10',
  'asignatura' : 'Y10',
  'sigla' : 'G12',
  'seccion' : 'S12',
  'jornada' : 'AC12',
  'sede' : 'AM12'    
}
# Fila en la que comienzan los indicadores
row = 30
# Nombre de las columnas en la planilla
columns = 'DFHJL'
# Posibles valores de los indicadores
valores = "abcde"
# Cantidad de indicadores que tiene el instrumento
n_indicadores = 8
# Nombre de la plantilla
wb_name = "PlantillaET4.xlsx"
# Nombre de la hoja
sheet_name = "ET4"

for indice in range(resultados.shape[0]):
  row = 30 # Reinicia la fila
  # Selecciona la plantilla que se debe completar
  wb = openpyxl.load_workbook(wb_name)
  # Selecciona la hoja que se debe completar
  sheet = wb[sheet_name]

  # Obtiene los datos generales del archivo
  for dato, celda in dict_celdas.items():
    sheet[celda] = resultados[dato.upper()][indice]

  # Obtiene los indicadores
  indicadores = dict()
  for i in range(1, n_indicadores+1):
    indicadores[str(i)] = resultados['I' + str(i)][indice]
  # Iteramos los datos para ir marcando el indicador
  for indicador, valor in indicadores.items():  
    col = columns[valores.index(str(valor).lower())] #corregido...version para PC
    name_col = col + str(row)
    sheet[name_col] = 'x'
    row += 2

  # Guarda el archivo usando el nombre del estudiante
  nombre = resultados['nombre'.upper()][indice]
  destino = "ET4-" + nombre + ".xlsx"
  wb.save(destino)    
  
  # Cierra la plantilla para poder abrirla con el siguiente estudiante
  wb.close()