import os
import pandas as pd

def executar_automacao():
    arquivo_entrada = "notas_alunos.xlsx"
    arquivo_saida = "relatorio_processado.xlsx"

    print("--- Iniciando Processamento ---")

    # Verifica se o arquivo existe, se não, cria um de teste
    if not os.path.exists(arquivo_entrada):
        print(f"Aviso: '{arquivo_entrada}' não encontrado. Criando arquivo de teste...")
        dados_teste = {
            'Aluno': ['Alice', 'Bruno', 'Carla', 'David'],
            'Media': [8.5, 6.0, 9.0, 5.5],
            'Faltas': [2, 15, 5, 8]
        }
        pd.DataFrame(dados_teste).to_excel(arquivo_entrada, index=False)
        print(f"Arquivo '{arquivo_entrada}' criado com sucesso!")

    try:
        # Lendo o arquivo
        df = pd.read_excel(arquivo_entrada)

        if 'Media' not in df.columns or 'Faltas' not in df.columns:
            print("ERRO: As colunas 'Media' ou 'Faltas' não foram encontradas.")
            return

        print("Processando lógica de aprovação...")
        df['Status_Final'] = df.apply(
            lambda row: 'Aprovado' if row['Media'] >= 7 and row['Faltas'] <= 10 else 'Reprovado',
            axis=1
        )

        # Salvando o resultado
        df.to_excel(arquivo_saida, index=False)
        print(f"--- SUCESSO! Relatório gerado: {arquivo_saida} ---")

    except PermissionError:
        print("ERRO: Feche o arquivo Excel antes de rodar o código!")
    except Exception as e:
        print(f"Erro inesperado: {e}")

if __name__ == "__main__":
    executar_automacao()