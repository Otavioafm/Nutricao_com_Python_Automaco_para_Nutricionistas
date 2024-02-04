import pandas as pd
import os

pasta_destino = 'Planilhas Pacientes'
if not os.path.exists(pasta_destino):
    os.makedirs(pasta_destino)

def adicionar_dados_excel(arquivo, dados):
    if os.path.isfile(arquivo):
        df = pd.read_excel(arquivo)
    else:
        df = pd.DataFrame()
    dados_df = pd.DataFrame([dados])
    df = pd.concat([df, dados_df], ignore_index=True)
    df.to_excel(arquivo, index=False)

Pacientes_Info = []
Pacientes_Abaixo_do_Peso = []
Paciente_Normal = []
Pacientes_Acima_do_Peso = []
Paciente_Obeso_Grau_1 = []
Paciente_Obeso_Grau_2 = []
Paciente_Obeso_Grau_3 = []

#=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
def calcular_imc(peso, altura):
    imc = peso / (altura ** 2)
    return round(imc, 1)
#__________________________________________________________
def calcular_Mifflin_Masculino(peso,altura,idade):
    GEB_H=(10*peso)+(6.25*altura)-(5*idade)+5
    return round(GEB_H,1)

def calcular_Mifflin_Feminino(peso,altura,idade):
    GEB_F=(10*peso)+(6.25*altura)-(5*idade)-161
    return round(GEB_F,1)
#__________________________________________________________

def calcular_Tinsley(peso,altura,idade):
    altura_metros = altura / 100
    GEB_Tinsley=(22*peso)+(500*altura_metros)+(3*idade)+301
    return round(GEB_Tinsley,1)
#__________________________________________________________

def calcular_Harris_Benedict_Masculino(peso,altura,idade):
    GEB_Harris_Benedict_H=88.362+(13.3978*peso)+(4.799*altura)-(5.677*idade)
    return round(GEB_Harris_Benedict_H,1)

def calcular_Harris_Benedict_Feminino(peso,altura,idade):
    GEB_Harris_Benedict_F=447.593+(9.247*peso)+(3.098*altura)-(4.330*idade)
    return round(GEB_Harris_Benedict_F,1)
#__________________________________________________________

def calcular_Cunningham_comum(peso):
    MB_Comum=500+(22*peso)
    return round(MB_Comum,1)

def calcular_Cunningham_Atletas(peso):
    MB_Atleta=800+(22*peso)
    return round(MB_Atleta,1)

#=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

print(f"{'=-='*20}\nPor favor, Utilize Ponto Final!! e NÃO UTILIZE vírgula!\n{'=-='*20}")
Contador = int(input("Quantos Pacientes você deseja cadastrar: "))

for Cont in range(Contador):
    Nome = str(input(f"\nDigite o nome do seu paciente {Cont}: ")).capitalize()
    Sexo = str(input(f"Sexo do seu Paciente {Cont} [M/F]: ")).capitalize()
    Idade = int(input(f"Idade do seu paciente {Cont}: "))

    Peso = float(input(f"Peso do seu paciente {Cont}: "))
    Altura = float(input(f"Altura do seu paciente {Cont}: "))

    imc_resultado = calcular_imc(Peso, Altura)

    paciente_info = {"Nome": Nome,"Sexo": Sexo,"Idade": Idade,"Peso": Peso, "Altura": Altura, "IMC": imc_resultado}

    if Sexo == "M":
        GEB_Masculino_Resultado = calcular_Mifflin_Masculino(Peso, Altura, Idade)
        paciente_info["Mifflin"] = GEB_Masculino_Resultado

        GEB_tinsley_Resultado = calcular_Tinsley(Peso, Altura, Idade)
        paciente_info["Tinsley"] = GEB_tinsley_Resultado

        GEB_Harris_Benedict_Masculino_Resultado = calcular_Harris_Benedict_Masculino(Peso, Altura, Idade)
        paciente_info["Benedict"] = GEB_Harris_Benedict_Masculino_Resultado

        MB_Cunningham_comum = calcular_Cunningham_comum(Peso)
        paciente_info["Cunningham_C"] = MB_Cunningham_comum

        MB_Cunningham_Atleta = calcular_Cunningham_Atletas(Peso)
        paciente_info["Cunningham_A"] = MB_Cunningham_Atleta

    # =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

    elif Sexo == "F":
        GEB_Feminino_Resultado = calcular_Mifflin_Feminino(Peso, Altura, Idade)
        paciente_info["Mifflin"] = GEB_Feminino_Resultado

        GEB_tinsley_Resultado = calcular_Tinsley(Peso, Altura, Idade)
        paciente_info["Tinsley"] = GEB_tinsley_Resultado

        GEB_Harris_Benedict_Feminino_Resultado = calcular_Harris_Benedict_Feminino(Peso, Altura, Idade)
        paciente_info["Benedict"] = GEB_Harris_Benedict_Feminino_Resultado

        MB_Cunningham_comum = calcular_Cunningham_comum(Peso)
        paciente_info["Cunningham_C"] = MB_Cunningham_comum

        MB_Cunningham_Atleta = calcular_Cunningham_Atletas(Peso)
        paciente_info["Cunningham_A"] = MB_Cunningham_Atleta
    else:
        print("Sexo não reconhecido.")

    adicionar_dados_excel(os.path.join(pasta_destino, 'pacientes_info.xlsx'), paciente_info)
    if imc_resultado <= 18.5:
        adicionar_dados_excel(os.path.join(pasta_destino, 'pacientes_abaixo_do_peso.xlsx'), paciente_info)
    elif 18.5 < imc_resultado <= 25:
        adicionar_dados_excel(os.path.join(pasta_destino, 'paciente_normal.xlsx'), paciente_info)
    elif 25 < imc_resultado <= 30:
        adicionar_dados_excel(os.path.join(pasta_destino, 'pacientes_acima_do_peso.xlsx'), paciente_info)
    elif 30 < imc_resultado <= 35:
        adicionar_dados_excel(os.path.join(pasta_destino, 'paciente_obeso_grau_1.xlsx'), paciente_info)
    elif 35 < imc_resultado <= 40:
        adicionar_dados_excel(os.path.join(pasta_destino, 'paciente_obeso_grau_2.xlsx'), paciente_info)
    elif imc_resultado > 40:
        adicionar_dados_excel(os.path.join(pasta_destino, 'paciente_obeso_grau_3.xlsx'), paciente_info)

print("=-=" * 15)
print("Dados dos pacientes foram registrados com sucesso!")
print("=-=" * 15)
