import time
import pandas as pd
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import os


class MyHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory:
            return None

        if event.src_path.endswith(".xlsx"):
            self.process_excel(event.src_path)

    def process_excel(self, file_path):
        # Lê as sheets do arquivo
        xls = pd.ExcelFile(file_path)
        tarefas_df = pd.read_excel(xls, "Tarefas")
        nome_plano_df = pd.read_excel(xls, "Nome do plano")

        # Supondo que 'Nome do plano' tenha as colunas 'Nome' e 'ID' para o nome e id do plano
        nome_do_plano = nome_plano_df["Nome"].iloc[0]
        id_do_plano = nome_plano_df["ID"].iloc[0]

        # Insere as informações nas duas primeiras colunas da sheet 'Tarefas'
        tarefas_df.insert(0, "UGIOvEq7v0W_S3fP2Uf03WUAEBmx", id_do_plano)
        tarefas_df.insert(1, "BACKLOG_DR_DOC_SPRINT 4", nome_do_plano)

        # Salva a sheet modificada em um novo arquivo
        nome_arquivo_saida = f"saida_{os.path.basename(file_path)}"
        tarefas_df.to_excel(nome_arquivo_saida, index=False)
        print(f"Arquivo processado e salvo como: {nome_arquivo_saida}")


if __name__ == "__main__":
    path = r"C:\\13_AUTOMACAO_PLANNER"
    event_handler = MyHandler()
    observer = Observer()
    observer.schedule(event_handler, path, recursive=False)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
