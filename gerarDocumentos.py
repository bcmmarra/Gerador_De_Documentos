import pandas as pd
from docxtpl import DocxTemplate
import os
import sys
import re

# --- Configurações Principais ---
NOME_PLANILHA = 'dados_documentos.xlsx'
PASTA_DADOS = 'dados'
PASTA_TEMPLATES = 'modelos'
PASTA_SAIDA = 'documentos_gerados'
VALOR_PADRAO_VAZIO = 'N/A' 
COLUNA_TEMPLATE = 'NOME_DO_MODELO'      
COLUNA_NOME_CLIENTE = 'CLIENTE'         
COLUNA_NOME_DOCUMENTO = 'DOCUMENTO'   
COLUNA_NUMERO_PREGAO = 'NUMERO_PREGAO'  

def limpar_nome_arquivo(texto):
    """
    Remove caracteres que não são permitidos em nomes de arquivo e substitui por '_'.
    Mantém letras, números, espaços e hífens, substituindo / e outros.
    """
    texto = str(texto).strip()
    texto = texto.replace('/', '_').replace('\\', '_').replace('.', '_') 
    texto = re.sub(r'[\s_]+', '_', texto)
    
    return texto

def gerar_documentos():
    """
    Função principal que gerencia o carregamento de dados,
    a identificação do modelo, a renderização e o salvamento.
    """
    print("🚀 Iniciando a Automação de Geração de Documentos (Múltiplos Modelos)...")
    print("-" * 60)

    # 1. Preparação dos Caminhos e Pastas
    caminho_planilha = os.path.join(PASTA_DADOS, NOME_PLANILHA)
    os.makedirs(PASTA_SAIDA, exist_ok=True)

    # 2. Carregar e Limpar os Dados
    try:
        df = pd.read_excel(caminho_planilha).fillna(VALOR_PADRAO_VAZIO)
        df_limpo = df[df[COLUNA_TEMPLATE] != VALOR_PADRAO_VAZIO].copy()
        linhas_descartadas = len(df) - len(df_limpo)
        df = df_limpo 

        if linhas_descartadas > 0:
            print(f"🧹 Atenção: {linhas_descartadas} linhas vazias ou sem modelo foram descartadas.")

        if df.empty:
            print(f"AVISO: A planilha '{NOME_PLANILHA}' está vazia após a limpeza. Nenhuma ação será realizada.")
            return

        # Verifica se as colunas críticas existem
        colunas_criticas = [COLUNA_TEMPLATE, COLUNA_NOME_CLIENTE, COLUNA_NOME_DOCUMENTO, COLUNA_NUMERO_PREGAO]
        for col in colunas_criticas:
            if col not in df.columns:
                 print(f"❌ ERRO CRÍTICO: Coluna '{col}' não encontrada na planilha. Verifique a ortografia.")
                 sys.exit(1)


    except FileNotFoundError:
        print(f"❌ ERRO CRÍTICO: Arquivo de dados '{caminho_planilha}' não encontrado.")
        sys.exit(1)

    except Exception as e:
        print(f"❌ ERRO ao ler a planilha Excel: {e}")
        sys.exit(1)
        
    contador = 0

    # 3. Processar e Gerar Documentos
    for dados_documentos in df.to_dict('records'):
        
        nome_template_completo = str(dados_documentos.get(COLUNA_TEMPLATE))
        nome_cliente = str(dados_documentos.get(COLUNA_NOME_CLIENTE))
        nome_documento = str(dados_documentos.get(COLUNA_NOME_DOCUMENTO))
        numero_pregao = str(dados_documentos.get(COLUNA_NUMERO_PREGAO))
            
        caminho_template_completo = os.path.join(PASTA_TEMPLATES, nome_template_completo)
        
        try:
            # Carrega e renderiza
            doc_template = DocxTemplate(caminho_template_completo) 
            context = {k: v for k, v in dados_documentos.items()}
            doc_template.render(context)
            
            # --- 4. NOME DO ARQUIVO ---
            nome_documento_limpo = limpar_nome_arquivo(nome_documento)
            nome_cliente_limpo = limpar_nome_arquivo(nome_cliente)
            numero_pregao_limpo = limpar_nome_arquivo(numero_pregao) # Limpeza do Número do Pregão
            nome_arquivo = f"{nome_documento_limpo}_{nome_cliente_limpo}_{numero_pregao_limpo}.docx"
            
            if numero_pregao_limpo == VALOR_PADRAO_VAZIO:
                 nome_arquivo = f"{nome_documento_limpo}_{nome_cliente_limpo}.docx"
                        
            caminho_saida = os.path.join(PASTA_SAIDA, nome_arquivo)
            
            doc_template.save(caminho_saida)
            
            contador += 1
            print(f"✅ Gerado ({contador}): {nome_arquivo} (Modelo: {nome_template_completo})")

        except FileNotFoundError:
            print(f"❌ ERRO: Arquivo de modelo não encontrado! Caminho: '{caminho_template_completo}'.")
            print(f"   Verifique se o valor '{nome_template_completo}' na coluna '{COLUNA_TEMPLATE}' está correto. Pulando registro.")
        except Exception as e:
            print(f"⚠️ ERRO Geral ao processar o registro {contador+1} (Cliente: {nome_cliente}): {e}.")
            print("   Pode ser erro no placeholder no documento Word. Pulando registro.")

    # 5. Conclusão
    print("-" * 60)
    print(f"🎉 Automação Concluída!")
    print(f"{contador} documentos gerados com sucesso na pasta '{PASTA_SAIDA}'.")

if __name__ == "__main__":
    gerar_documentos()