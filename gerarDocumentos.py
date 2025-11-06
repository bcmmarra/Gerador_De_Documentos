# Importa a biblioteca pandas para manipulação de DataFrames (leitura do Excel)
import pandas as pd
# Importa DocxTemplate para trabalhar com modelos de documentos Word (.docx)
from docxtpl import DocxTemplate
# Importa 'os' para interações com o sistema operacional (caminhos de arquivos, criação de pastas)
import os
# Importa 'sys' para interações com o interpretador (como sair do script em caso de erro crítico)
import sys
# Importa 're' para usar Expressões Regulares (limpeza de nomes de arquivos)
import re
# Importa 'warnings' para gerenciar avisos
import warnings

# Filtra e ignora avisos específicos (UserWarning) que podem ser gerados por bibliotecas, 
# mantendo o console mais limpo.
warnings.filterwarnings("ignore", category=UserWarning)

# ==============================================================================================
# === FUNÇÃO AUXILIAR PARA COMPATIBILIDADE COM PYINSTALLER (Adicionada) ===
# ==============================================================================================

def resource_path(relative_path):
    """ 
    Obtém o caminho absoluto para o recurso empacotado, seja no ambiente de desenvolvimento 
    ou dentro do executável PyInstaller. 
    """
    if getattr(sys, 'frozen', False):
        # Estamos rodando como um executável PyInstaller. Caminho base é o diretório temporário.
        base_path = sys._MEIPASS
    else:
        # Estamos rodando em modo normal (Python). Caminho base é o diretório do script.
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, relative_path)

# ==============================================================================================
# === Variáveis de Configuração (Constantes) - Caminhos AJUSTADOS ===
# ==============================================================================================
# Nome do arquivo Excel que contém os dados a serem preenchidos nos documentos
NOME_PLANILHA = 'dados_documentos.xlsx'
# Nome da pasta onde o arquivo Excel de dados está localizado
# >>> ALTERAÇÃO 1: Usa resource_path para ler a cópia empacotada no EXE
PASTA_DADOS = resource_path('dados') 
# Nome da pasta onde os modelos (templates) de documentos Word estão localizados
# >>> ALTERAÇÃO 2: Usa resource_path para ler a cópia empacotada no EXE
PASTA_TEMPLATES = resource_path('modelos') 
# Nome da pasta onde os documentos finais serão salvos (DEVE ficar FORA do .exe)
# >>> ALTERAÇÃO 3: Fica como caminho relativo normal para salvar no disco.
PASTA_SAIDA = 'documentos_gerados'

# Valor padrão usado para preencher células vazias na planilha ou novas colunas
VALOR_PADRAO_VAZIO = 'N/A' 
# Nome da coluna no Excel que especifica qual arquivo de modelo Word deve ser usado
COLUNA_TEMPLATE = 'NOME_DO_MODELO'
# Nome da coluna que contém o nome do cliente (usado no nome do arquivo de saída)
COLUNA_NOME_CLIENTE = 'CLIENTE'
# Nome da coluna que contém a descrição do documento (usado no nome do arquivo de saída)
COLUNA_NOME_DOCUMENTO = 'DOCUMENTO' 
# Nome da coluna que contém o número do pregão (usado opcionalmente no nome do arquivo de saída)
COLUNA_NUMERO_PREGAO = 'NUMERO_PREGAO' 

# --------------------------------------------------------------------------------------------------

def limpar_nome_arquivo(texto):
    """
    Função para limpar e formatar strings para que possam ser usadas como nomes de arquivos.
    """
    # Converte para string e remove espaços em branco no início e fim
    texto = str(texto).strip()
    # Substitui caracteres problemáticos ('/', '\', '.') por '_'
    texto = texto.replace('/', '_').replace('\\', '_').replace('.', '_') 
    # Usa expressão regular para substituir um ou mais espaços ou underscores por um único underscore
    texto = re.sub(r'[\s_]+', '_', texto)
    
    return texto

# --------------------------------------------------------------------------------------------------

def extrair_variaveis_do_template(caminho_modelo):
    """
    Função que abre um modelo Word e extrai todos os placeholders (variáveis de contexto).
    """
    try:
        # Cria um objeto DocxTemplate
        doc = DocxTemplate(caminho_modelo)
        # Usa o método get_undeclared_template_variables() para encontrar todas as variáveis Jinja2
        context_placeholders = set(doc.get_undeclared_template_variables())
        
        # Filtra as variáveis para remover comandos Jinja2 (como 'tr', 'for', 'if', 'block')
        placeholders_filtrados = {
            var for var in context_placeholders 
            if not var.startswith(('tr', 'for', 'if', 'block'))
        }
        return placeholders_filtrados
    except Exception as e:
        # Em caso de erro (ex: arquivo corrompido ou inacessível)
        print(f"⚠️ ERRO ao extrair variáveis de {caminho_modelo}: {e}")
        return set()

# --------------------------------------------------------------------------------------------------

def checar_e_atualizar_colunas(df, caminho_planilha_disco):
    """
    Verifica se todas as variáveis encontradas em TODOS os modelos Word existem como colunas
    no DataFrame (planilha Excel). Se novas variáveis forem encontradas, elas são adicionadas
    ao DataFrame e a planilha é salva no CAMINHO DO DISCO (caminho_planilha_disco).
    """
    print("\n🔍 Iniciando checagem de variáveis dos templates vs. Planilha...")
    # Obtém um conjunto com todos os nomes de colunas atuais no DataFrame
    todas_colunas_excel = set(df.columns)
    # Conjunto para armazenar novas variáveis encontradas nos modelos, mas não no Excel
    novas_variaveis_encontradas = set()
    
    # Percorre recursivamente a pasta de modelos para encontrar todos os arquivos .docx
    for root, _, files in os.walk(PASTA_TEMPLATES):
        for file in files:
            if file.endswith('.docx'):
                caminho_modelo = os.path.join(root, file)
                # Extrai as variáveis do modelo atual
                variaveis_do_modelo = extrair_variaveis_do_template(caminho_modelo)
                
                # Compara cada variável extraída com as colunas existentes no Excel
                for var in variaveis_do_modelo:
                    if var not in todas_colunas_excel:
                        novas_variaveis_encontradas.add(var)

    # Lógica para atualização da planilha
    if novas_variaveis_encontradas:
        print("-" * 60)
        print(f"⚠️ **ATENÇÃO: NOVAS VARIÁVEIS ENCONTRADAS**")
        print("As seguintes variáveis foram encontradas nos templates, mas não existem como colunas na planilha:")
        print(f"{', '.join(sorted(novas_variaveis_encontradas))}")
        print("\n✅ As colunas serão adicionadas à planilha e preenchidas com 'N/A'.")
        print("-" * 60)
        
        # Adiciona as novas colunas ao DataFrame e preenche com o valor padrão
        for nova_coluna in novas_variaveis_encontradas:
            df[nova_coluna] = VALOR_PADRAO_VAZIO
            
        try:
            # Salva o DataFrame atualizado de volta no arquivo Excel NO DISCO
            df.to_excel(caminho_planilha_disco, index=False, engine='openpyxl')
            print(f"💾 Planilha '{NOME_PLANILHA}' atualizada com sucesso no disco.")
            return True # Retorna True indicando que a planilha foi modificada
        except Exception as e:
            # Em caso de erro ao salvar (ex: arquivo aberto por outro programa)
            print(f"❌ ERRO CRÍTICO ao salvar a planilha Excel no disco: {e}")
            print("Verifique se o arquivo Excel não está aberto por outro programa.")
            sys.exit(1) # Sai do programa
    else:
        print("✅ Planilha e templates estão em sincronia. Nenhuma coluna nova foi adicionada.")
        return False # Retorna False indicando que a planilha não foi modificada

# --------------------------------------------------------------------------------------------------

def gerar_documentos():
    """
    Função principal que coordena o fluxo de leitura, verificação e geração de documentos.
    """
    print("🚀 Iniciando a Automação de Geração de Documentos (Múltiplos Modelos)...")
    print("-" * 60)

    if getattr(sys, 'frozen', False):
        # Modo PyInstaller (EXE está em 'dist', precisa subir um nível para acessar 'dados')
        caminho_planilha_disco = os.path.join('..', 'dados', NOME_PLANILHA)
    else:
        # Modo de Desenvolvimento (Script está no raiz do projeto, a pasta 'dados' está ao lado)
        caminho_planilha_disco = os.path.join('dados', NOME_PLANILHA)

    # Cria a pasta de saída se ela não existir
    os.makedirs(PASTA_SAIDA, exist_ok=True)
    
    # --- Leitura e Preparação Inicial do DataFrame ---
    try:
        # Tenta LER O ARQUIVO NO DISCO (prioridade 1: versão editada pelo usuário)
        df = pd.read_excel(caminho_planilha_disco).fillna(VALOR_PADRAO_VAZIO)
        print(f"💾 Planilha lida do disco: '{caminho_planilha_disco}'.")

    except FileNotFoundError:
        # Se for a primeira execução e o arquivo não existe no disco:
        print(f"⚠️ A planilha de dados não foi encontrada no disco original. Tentando ler backup...")
        try:
            # Garante que a pasta 'dados' existe no disco
            os.makedirs(os.path.dirname(caminho_planilha_disco), exist_ok=True)
            
            # COPIA a estrutura lida do backup para o disco
            df.to_excel(caminho_planilha_disco, index=False, engine='openpyxl')
            print(f"✅ Backup copiado. O arquivo '{caminho_planilha_disco}' foi criado no disco para edição.")
        except Exception as e:
            # Trata o erro de backup ou erro ao salvar a cópia inicial
            print(f"❌ ERRO CRÍTICO ao ler ou copiar a planilha: {e}")
            sys.exit(1)
            
    except Exception as e:
        # Trata outros erros de leitura do Excel (ex: arquivo aberto)
        print(f"❌ ERRO ao ler a planilha Excel do disco: {e}")
        sys.exit(1)

    # --- Sincronização de Colunas ---
    # O df contém a versão mais atualizada (ou a cópia inicial). 
    # O caminho do disco é passado para que, se houver novas variáveis, a versão no disco seja sobrescrita.
    df_foi_modificado = checar_e_atualizar_colunas(df, caminho_planilha_disco)
    
    if df_foi_modificado:
        # Se a planilha foi modificada (novas colunas adicionadas), o script para
        print("\n🛑 POR FAVOR: Preencha os novos campos adicionados na planilha Excel ANTES de executar novamente.")
        input("\nPressione ENTER para fechar e continuar (depois de preencher o Excel).")
        return # Termina a execução da função principal
    
    # --- Limpeza do DataFrame ---
    # Cria uma cópia do DataFrame descartando linhas onde a coluna de modelo é o valor padrão
    df_limpo = df[df[COLUNA_TEMPLATE] != VALOR_PADRAO_VAZIO].copy()
    # ... (o restante da função é o mesmo) ...
    # ...
    # ...
    
    linhas_descartadas = len(df) - len(df_limpo)
    df = df_limpo # Atribui o DataFrame limpo de volta à variável principal
    
    if linhas_descartadas > 0:
        print(f"🧹 Atenção: {linhas_descartadas} linhas vazias ou sem modelo foram descartadas.")

    if df.empty:
        print(f"AVISO: A planilha '{NOME_PLANILHA}' está vazia após a limpeza. Nenhuma ação será realizada.")
        return

    # --- Checagem de Colunas Críticas ---
    # Verifica se as colunas essenciais para o funcionamento do script existem no DataFrame
    colunas_criticas = [COLUNA_TEMPLATE, COLUNA_NOME_CLIENTE, COLUNA_NOME_DOCUMENTO, COLUNA_NUMERO_PREGAO]
    for col in colunas_criticas:
        if col not in df.columns:
             print(f"❌ ERRO CRÍTICO: Coluna '{col}' não encontrada na planilha. Verifique a ortografia.")
             sys.exit(1)
             
    contador = 0 # Inicializa um contador de documentos gerados

    # --- Geração de Documentos (Loop Principal) ---
    # Itera sobre cada linha do DataFrame, tratando cada linha como um dicionário de dados (Contexto)
    for dados_documentos in df.to_dict('records'):
        
        # 1. Extrai dados das colunas críticas para o processamento
        nome_template_completo = str(dados_documentos.get(COLUNA_TEMPLATE))
        nome_cliente = str(dados_documentos.get(COLUNA_NOME_CLIENTE))
        nome_documento = str(dados_documentos.get(COLUNA_NOME_DOCUMENTO))
        numero_pregao = str(dados_documentos.get(COLUNA_NUMERO_PREGAO)) 
                
        # 2. Constrói o caminho completo para o modelo
        # Substitui barras (para suportar subpastas no nome do template) pelo separador de caminho do SO
        nome_template_tratado = nome_template_completo.replace('/', os.sep).replace('\\', os.sep)
        # PASTA_TEMPLATES já é o caminho PyInstaller/_MEIPASS/modelos
        caminho_template_completo = os.path.join(PASTA_TEMPLATES, nome_template_tratado)
        
        # 3. Processamento do Documento
        try:
            # Carrega o modelo Word
            doc_template = DocxTemplate(caminho_template_completo) 
            # O dicionário de dados da linha atual é o contexto completo para o render
            context = {k: v for k, v in dados_documentos.items()}
            # Preenche o modelo com os dados (renderiza)
            doc_template.render(context)
            
            # 4. Criação do Nome do Arquivo de Saída
            # Limpa os dados para garantir um nome de arquivo válido
            nome_documento_limpo = limpar_nome_arquivo(nome_documento)
            nome_cliente_limpo = limpar_nome_arquivo(nome_cliente)
            numero_pregao_limpo = limpar_nome_arquivo(numero_pregao)
            
            # Define o padrão do nome do arquivo (com ou sem o número do pregão)
            if numero_pregao_limpo == VALOR_PADRAO_VAZIO:
                nome_arquivo = f"{nome_documento_limpo}_{nome_cliente_limpo}.docx"
            else:
                nome_arquivo = f"{nome_documento_limpo}_{nome_cliente_limpo}_{numero_pregao_limpo}.docx"
            
            # Constrói o caminho de salvamento completo
            caminho_saida = os.path.join(PASTA_SAIDA, nome_arquivo)
            
            # 5. Salva o Documento Gerado
            doc_template.save(caminho_saida)
            
            contador += 1
            print(f"✅ Gerado ({contador}): {nome_arquivo} (Modelo: {nome_template_completo})")

        # 6. Tratamento de Erros
        except FileNotFoundError:
            # Erro específico para quando o arquivo de modelo não é encontrado
            print(f"❌ ERRO: Arquivo de modelo não encontrado! Caminho: '{caminho_template_completo}'.")
            print(f"   Verifique se o valor '{nome_template_completo}' na coluna '{COLUNA_TEMPLATE}' está correto e o arquivo existe. Pulando registro.")
        except Exception as e:
            # Tratamento genérico para outros erros (ex: erro no placeholder no .docx)
            print(f"⚠️ ERRO Geral ao processar o registro {contador+1} (Cliente: {nome_cliente}): {e}.")
            print("   Pode ser erro no placeholder no documento Word ou outro problema. Pulando registro.")

    # --- Conclusão ---
    print("-" * 60)
    print(f"🎉 Automação Concluída!")
    print(f"{contador} documentos gerados com sucesso na pasta '{PASTA_SAIDA}'.")
    
    input("\nPressione ENTER para fechar a janela.")
    
if __name__ == "__main__":
    # Chama a função principal
    gerar_documentos()