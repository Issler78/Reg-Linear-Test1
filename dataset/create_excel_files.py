import pandas as pd
from random import randint
from faker import Faker
import datetime
import calendar
import openpyxl
from random import choice, choices

fake = Faker("pt_BR")

# Definir alguns produtos ficticios no formato [ID, Produto, Categoria, Preço Unitário]
produtos = [
    [101, "Fone Bluetooth", "Eletônicos", 120.00],
    [102, "Smartphone", "Eletônicos", 1300.00],
    [201, "Máquina de Lavar", "Eletrodomésticos", 1500.00],
    [202, "Geladeira Duplex", "Eletrodomésticos", 2800.00],
    [504, "Calça Jeans", "Vestuário", 150.00],
    [505, "Camiseta Polo", "Vestuário", 70.00]
]

# Criar uma planilha para cada mes ate o mes dez
n_random = 0
for mes in range(1, 11):
    data = {
        "Data": [],
        "ID Produto": [],
        "Nome Produto": [],
        "Categoria": [],
        "Região": [],
        "Quantidade": [],
        "Preço Unitário (R$)": [],
    }
    
    
    # Variaveis para serem usadas no loop abaixo
    last_day = calendar.monthrange(year=2024, month=mes)[1]
    start_date = datetime.date(year=2024, month=mes, day=1)
    end_date = datetime.date(year=2024, month=mes, day=last_day)
    
    # Quantos dados terao em cada tabela | escolher entre 40 a 50 vendas no 1° mes e com no maximo 10 de diferenca nos meses posteriores comparado com o anterior
    if n_random == 0:
        n = randint(40, 50)
    else:
        n = randint(n_random-5, n_random+7)
    n_random = n
    
    for i in range(n_random):
        # Adicionar em data uma data entre o primeiro e o ultimo dia do mes em que o loop esta
        data["Data"].append(
            fake.date_between(start_date, end_date)
        )
        
        # adicionar em data as infos dos produtos
        indice_produto = choices([0, 1, 2, 3, 4, 5], weights=[2, 0.5, 0.5, 0.2, 3, 6], k=1)[0]
        data["ID Produto"].append(produtos[indice_produto][0])
        data["Nome Produto"].append(produtos[indice_produto][1])
        data["Categoria"].append(produtos[indice_produto][2])
        data["Preço Unitário (R$)"].append(produtos[indice_produto][3])
        data["Quantidade"].append(choices([1, 2, 3, 4], weights=[6, 3, 2, 0.5], k=1)[0])
        data["Região"].append(choice(["Sul", "Centro-Oeste", "Sudeste", "Nordeste", "Norte"]))
    # ordenar as datas
    data["Data"] = sorted(data["Data"])
    
    
    
    # criar dataframe com os dados
    df = pd.DataFrame(data)

    # criar a coluna de preço total
    df["Preço Total (R$)"] = df["Preço Unitário (R$)"] * df["Quantidade"]
    
    
    # transformar em excel
    df.to_excel(f"dataset/files/vendas{mes}.xlsx", sheet_name="Vendas", index=False)
    
    
    
    # formatar excel
    wb = openpyxl.load_workbook(f"dataset/files/vendas{mes}.xlsx")
    ws = wb["Vendas"]
    
    for row in ws.iter_rows(min_row=2):
        data = row[0]
        p_unit = row[6]
        p_tot = row[7]
        
        data.number_format = "dd/mm/yyyy"
        p_unit.number_format = "R$ #.##0,0"
        p_tot.number_format = "R$ #.##0,0"
    wb.save(f"dataset/files/vendas{mes}.xlsx")