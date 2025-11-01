import pandas as pd
from docxtpl import DocxTemplate
import os

# --- Configurações ---
NOME_TEMPLATE = 'CARTA PROPOSTA.docx'
NOME_PLANILHA = 'dados_documentos.xlsx'
PASTA_DADOS = 'dados'
PASTA_SAIDA = 'documentos_gerados'

def gerar_documentos():
    # 1. Carregar os dados
    caminho_planilha = os.path.join(PASTA_DADOS, NOME_PLANILHA)
    
    try:
        df = pd.read_excel(caminho_planilha).fillna('SEM TEXTO')
    except FileNotFoundError:
        print(f"ERRO: Arquivo {caminho_planilha} não encontrado.")
        return

    # 2. Carregar o template
    try:
        doc_template = DocxTemplate(NOME_TEMPLATE)
    except FileNotFoundError:
        print(f"ERRO: Template {NOME_TEMPLATE} não encontrado na raiz do projeto.")
        return

    os.makedirs(PASTA_SAIDA, exist_ok=True)
    
    contador = 0
    # 3. Processar cada linha de dados
    for dados_documentos in df.to_dict('records'):
        context = {k: v for k, v in dados_documentos.items()}
        doc_template.render(context)

        # 4. Salvar o novo documento
        nome_arquivo = f"Documento_{dados_documentos.get('CLIENTE', contador)}.docx"
        caminho_saida = os.path.join(PASTA_SAIDA, nome_arquivo)
        
        doc_template.save(caminho_saida)
        
        contador += 1
        print(f"✅ Documento gerado: {nome_arquivo}")

    print(f"\n--- Automação Concluída ---")
    print(f"{contador} documentos gerados na pasta '{PASTA_SAIDA}'")


if __name__ == "__main__":
    gerar_documentos()