import os
import requests
import openpyxl
from dotenv import load_dotenv

load_dotenv()


url = 'https://fintual.cl/api/goals'
user_email = os.environ.get('USER_EMAIL')
user_token = os.environ.get('USER_TOKEN')

response = requests.get(url, params={
    'user_email': user_email,
    'user_token': user_token
}, timeout=10)

data = response.json()['data']

# Creamos un nuevo libro de Excel
workbook = openpyxl.Workbook()

# Seleccionamos la hoja activa del libro
sheet = workbook.active

# Agregamos los encabezados de las columnas
sheet['A1'] = 'ID'
sheet['B1'] = 'Nombre'
sheet['C1'] = 'Goal Type'
sheet['D1'] = 'Valor Neto Actual'
sheet['E1'] = 'Depositado'
sheet['F1'] = 'Ganancia'

# Iteramos a través de los datos y los agregamos al libro de Excel
for i, goal in enumerate(data):
    sheet[f'A{i+2}'] = goal['id']
    sheet[f'B{i+2}'] = goal['attributes']['name']
    sheet[f'C{i+2}'] = goal['attributes']['goal_type']
    sheet[f'D{i+2}'] = goal['attributes']['nav']
    sheet[f'E{i+2}'] = goal['attributes']['deposited']
    sheet[f'F{i+2}'] = goal['attributes']['profit']

# Sumamos los valores de cada columna
valor_neto_actual_total = sum(goal['attributes']['nav'] for goal in data)
depositado_total = sum(goal['attributes']['deposited'] for goal in data)
ganancia_total = sum(goal['attributes']['profit'] for goal in data)

# Agregamos una fila al final con las sumas
sheet.append(['Total', '', '', valor_neto_actual_total,
             depositado_total, ganancia_total])

# Guardamos el libro de Excel
workbook.save('datos_fintual.xlsx')
