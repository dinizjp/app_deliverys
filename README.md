# app_deliverys

## Descrição
O projeto `app_deliverys` é uma aplicação de consolidação de dados de deliverys integrados com o sistema interno da empresa. Inclui funcionalidades para conectar-se ao banco de dados, extrair informações relevantes, processar diferentes tipos de planilhas de delivery, e consolidar esses dados com os registros do sistema, gerando uma planilha para download.

## Sumário
- [Dependências](#dependências)
- [Instalação](#instalação)
- [Uso](#uso)
- [Estrutura de Pastas](#estrutura-de-pastas)

## Dependências
O projeto utiliza as seguintes bibliotecas listadas no `requirements.txt`:
- streamlit==1.32.0
- pandas==2.2.2
- XlsxWriter==3.2.0
- openpyxl==3.1.2
- xlrd==2.0.1
- pyodbc==5.2.0

## Instalação
Para configurar o ambiente de desenvolvimento, execute os seguintes comandos:

```bash
pip install -r requirements.txt
```

Isso instalará todas as dependências necessárias para rodar a aplicação.

## Uso
Para executar a aplicação, utilize o comando:

```bash
streamlit run home.py
```

Ao abrir a interface no navegador, o usuário poderá:
- Selecionar a empresa (loja) e o delivery desejado na sidebar.
- Definir o intervalo de datas para consulta.
- Carregar a planilha do delivery.
- Visualizar os dados extraídos e a consolidação com os dados do sistema.
- Baixar a planilha gerada com os resultados.

## Estrutura de Pastas
```
app_deliverys/
│
├── home.py
├── utils.py
└── requirements.txt
```

- `home.py`: Script principal com interface utilizando Streamlit para a seleção de filtros, upload de planilhas e execução da consolidação.
- `utils.py`: Módulo contendo funções para conexão ao banco de dados, consulta do sistema, processamento específico de cada tipo de planilha de delivery e consolidação de dados.
- `requirements.txt`: Lista de dependências necessárias para a execução do projeto.

Este documento apresenta uma visão geral clara e completa da configuração e funcionamento do projeto `app_deliverys`.