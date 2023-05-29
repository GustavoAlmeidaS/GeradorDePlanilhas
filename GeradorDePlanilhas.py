import random
from openpyxl import Workbook
from datetime import date, timedelta

print("Olá, bem vindo ao meu algoritimo, ele foi criado para gerar aleatoriamente uma planilha com nomes ficticios de possíveis crimes, incluindo informações sobre o caso. \nEsse projeto foi planejado devido a uma necessidade de gerar um banco de dados teste para um aplicativo em desenvolvimento na faculdade.\n\nUse como preferir! \n")
lines = int(input("How many lines values: "))
table = input("Name of the table (include .xlsx in the end): ")

# Generate 10 random Brazilian names
names = [random.choice([ "João", "Maria", "José", "Ana", "Pedro", "Sofia", "Lucas", "Rafaela", "Guilherme", "Luiza",
    "Renan", "Felipe", "Maicon", "Julia", "Thays", "Valdir", "Gustavo", "Mariana", "Fernando",
    "Larissa", "Diego", "Carolina", "Eduardo", "Isabela", "Matheus", "Letícia", "Vinicius",
    "Amanda", "Henrique", "Lívia", "Gabriel", "Marina", "Ricardo", "Bianca", "Fábio", "Natália",
    "Marcelo", "Camila", "André", "Jéssica", "Arthur", "Patrícia", "Roberto", "Raquel", "Carlos",
    "Laura", "Raphael", "Isabella", "Leandro", "Vitória", "Paulo", "Mariane", "Rodrigo"
]) for _ in range(lines)]

# Generate 10 random numbers with 11 digits
numbers = [random.randint(10000000000, 99999999999) for _ in range(lines)]

# Generate 10 random birth dates
start_date = date(1950, 1, 1)
end_date = date(2005, 12, 31)
birth_dates = [start_date + timedelta(days=random.randint(0, (end_date - start_date).days)) for _ in range(lines)]


tipos = [random.choice(["Roubo", "Furto", "Assédio", "Crime de ódio", "Feminicidio", "Homicídio", "Latrocínio", "Tráfico de Drogas", "Sequestro"]) for _ in range(lines)]

bens = [random.choice(["Celular", "Roupa", "Carteira", "Dinheiro", "Carro", "Bicicleta", "Notebook", "Joia", "Livro",
                       "Óculos", "Câmera", "Mochila"]) for _ in range(lines)]

data_inicio = date(2010, 1, 1)
data_fim = date(2023, 6, 30)
data_ocorrido = [data_inicio + timedelta(days=random.randint(0, (data_fim - data_inicio).days)) for _ in range(lines)]

# Create a Workbook object
workbook = Workbook()
sheet = workbook.active

# Set the column headers
headers = ["Nomes", "CPF Fictício", "Data de nascimento", "Tipo", "Bens", "Data do Ocorrido"]
sheet.append(headers)

# Add the data to the table
for name, number, birth_date, tipo, bens, data_ocorrido in zip(names, numbers, birth_dates, tipos, bens, data_ocorrido):
    row = [name, number, birth_date.strftime("%d/%m/%Y"), tipo, bens, data_ocorrido.strftime("%d/%m/%Y")]
    sheet.append(row)

# Save the workbook to a file
workbook.save(table)
