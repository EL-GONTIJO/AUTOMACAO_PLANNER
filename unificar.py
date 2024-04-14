import os
import pandas as pd

# Carrega a base de dados com a relação entre ID do plano e Empresa
caminho_base_empresas = "C:\\CRONOGRAMA_PROJETOS_TREO\\Projetos\\Projetos_todo.xlsx"
# Assumindo que a coluna C é o ID do plano e a coluna D é a Empresa
base_empresas = pd.read_excel(caminho_base_empresas, usecols="C:D")
base_empresas.columns = [
    "ID do plano",
    "Empresa",
]  # Renomeia as colunas para corresponder às do DataFrame de destino

# Pasta onde estão localizados os arquivos Excel de entrada
pasta_entrada = "C:\\CRONOGRAMA_PROJETOS_TREO\\Arquivos"

# Lista para armazenar os DataFrames de cada arquivo
dataframes = []

# Percorra todos os arquivos na pasta de entrada
for arquivo in os.listdir(pasta_entrada):
    if arquivo.endswith(".xlsx") and arquivo.startswith("BACKLOG_"):
        caminho_completo = os.path.join(pasta_entrada, arquivo)

        try:
            info_plano = pd.read_excel(
                caminho_completo,
                sheet_name="Nome do plano",
                header=None,
                usecols="B",
                nrows=2,
            )
            nome_plano, id_plano = info_plano.iloc[0, 0], info_plano.iloc[1, 0]
        except ValueError:
            print(
                f"A sheet 'Nome Plano' não foi encontrada no arquivo {arquivo}. Adicionando valores padrão."
            )
            nome_plano, id_plano = "Desconhecido", "Desconhecido"

        df = pd.read_excel(caminho_completo)
        df_temp = pd.DataFrame(
            {"Nome do plano": nome_plano, "ID do plano": id_plano}, index=df.index
        )
        df_final = pd.concat([df_temp, df], axis=1)
        dataframes.append(df_final)

if dataframes:
    df_concatenado = pd.concat(dataframes, ignore_index=True)
    df_concatenado = df_concatenado.merge(base_empresas, how="left", on="ID do plano")

    # Certifica-se de que a coluna Empresa exista após o merge, caso contrário, adiciona-a com valores vazios
    if "Empresa" not in df_concatenado.columns:
        df_concatenado["Empresa"] = pd.NA

    colunas = ["Empresa"] + [col for col in df_concatenado.columns if col != "Empresa"]
    df_concatenado = df_concatenado[colunas]

    pasta_saida = "C:\\CRONOGRAMA_PROJETOS_TREO\\cronograma_planner"
    if not os.path.exists(pasta_saida):
        os.makedirs(pasta_saida)

    nome_arquivo_saida = os.path.join(pasta_saida, "Notificacoes_PreJur.xlsx")
    df_concatenado.to_excel(nome_arquivo_saida, index=False)
    print(f"Arquivo {nome_arquivo_saida} criado com sucesso na pasta de saída!")
else:
    print("Nenhum arquivo com o prefixo 'BACKLOG_' encontrado para processamento.")
