
# 📄 Automatizador de Documentos Word (DOCX) via Excel

Este projeto em Python automatiza a geração de múltiplos documentos `.docx` a partir de uma planilha Excel (`dados_documentos.xlsx`) e de diversos templates Word. Ele utiliza a biblioteca `docxtpl` para preencher as variáveis (*placeholders*) nos documentos Word com os dados fornecidos linha por linha na planilha.

## 🚀 Como Iniciar o Projeto

Siga estes passos para configurar e executar o automatizador.

### 1\. Pré-requisitos

Você precisa ter o Python instalado (versão 3.6 ou superior).

### 2\. Configuração do Ambiente Virtual

É **altamente recomendado** usar um ambiente virtual (`venv`) para isolar as dependências do projeto.

```bash
# 1. Crie o ambiente virtual
python -m venv .venv

# 2. Ative o ambiente virtual
# No Windows (PowerShell):
.venv\Scripts\Activate

# Em Linux/macOS:
source .venv/bin/activate
```

### 3\. Instalação das Dependências

As dependências necessárias estão listadas no arquivo `requirements.txt` no projeto.

Com o ambiente virtual ativado, instale as bibliotecas usando o `pip`:

```bash
pip install -r requirements.txt
```

### 4\. Estrutura do Projeto

O script espera que a estrutura de pastas do projeto seja organizada da seguinte forma:

```
AUTOMATIZADORDOCUMENTOS/
├── .venv/                              # Ambiente virtual (ignorar no Git)
├── dados/
│   └── dados_documentos.xlsx           # Fonte de dados
├── documentos_gerados/                 # Pasta de SAÍDA (Criada automaticamente)
├── modelos/                            # Pasta RAIZ que contém todos os templates
│   └── modelosEdital/                  # Exemplo de Subpasta
│   │   └── CARTA PROPOSTA.docx         # Template
│   └── modelosPadrao/                  # Exemplo de Subpasta
│       └── REPRESENTANTE LEGAL.docx    # Template
├── gerarDocumentos.py                  # Script principal
├── requirements.txt                    # Lista de dependências (já incluso)
└── README.md                           # Este arquivo
```

### 5\. Configuração da Planilha (`dados_documentos.xlsx`)

A primeira linha da planilha deve conter os cabeçalhos (`headers`) correspondentes aos *placeholders* nos seus templates Word (`{{NOME_DA_VARIAVEL}}`).

As seguintes colunas são **obrigatórias** e usadas para a lógica do script:

| Coluna Python | Nome no Excel | Função | Exemplo de Valor |
| :--- | :--- | :--- | :--- |
| `COLUNA_TEMPLATE` | **NOME\_DO\_MODELO** | **Caminho relativo** do template a ser usado, a partir da pasta `modelos/`. | `modelosPadrao/CARTA PROPOSTA.docx` |
| `COLUNA_NOME_CLIENTE` | **CLIENTE** | Nome do cliente/usuário (Parte do nome do arquivo final). | `Policia Militar de Minas Gerais` |
| `COLUNA_NOME_DOCUMENTO` | **DOCUMENTO** | Título do documento (Parte do nome do arquivo final). | `CARTA PROPOSTA` |
| `COLUNA_NUMERO_PREGAO` | **NUMERO\_PREGAO** | Número do Pregão (Parte do nome do arquivo final). | `9003/2025` |
| (Outras Colunas) | *qualquer nome* | Variáveis que preencherão os *placeholders* no Word. | `{{VALOR_DA_PROPOSTA}}` |

> **OBSERVAÇÃO SOBRE LIMPEZA:** O script limpa automaticamente as linhas que possuem o campo **`NOME_DO_MODELO`** vazio após a leitura, garantindo que apenas registros válidos sejam processados.

### 6\. Execução do Script

Execute o script principal diretamente do terminal (com o ambiente virtual ativado):

```bash
python gerarDocumentos.py
```

O script irá:

1.  Ler a planilha `dados/dados_documentos.xlsx`.
2.  Para cada linha válida, carregar o template Word especificado na coluna `NOME_DO_MODELO`.
3.  Preencher o template com todos os dados da linha.
4.  Salvar o documento gerado na pasta `documentos_gerados/`.

## ⚙️ Detalhes da Automação

### Nomenclatura do Arquivo de Saída

O nome do arquivo de saída é construído combinando três campos cruciais da planilha, garantindo organização:

**`<DOCUMENTO>_<CLIENTE>_<NUMERO_PREGAO>.docx`**

### Tratamento de Caracteres

A função `limpar_nome_arquivo()` é aplicada a cada parte do nome de arquivo (`DOCUMENTO`, `CLIENTE`, `NUMERO_PREGAO`). Ela substitui automaticamente caracteres problemáticos (como `/`, `\` e `.`) por *underscore* (`_`), garantindo nomes de arquivo válidos em qualquer sistema operacional.

**Regras Específicas:**

  * Se a coluna `NUMERO_PREGAO` estiver vazia na planilha, ela será omitida do nome do arquivo, evitando que o nome final fique com um *underscore* extra.