# app_deliverys

## Descrição

O projeto *app_deliverys* é uma ferramenta destinada a consolidar e comparar planilhas de diferentes serviços de delivery. A aplicação foi desenvolvida para facilitar a análise dos pedidos, valores e discrepâncias entre várias bases de dados de delivery, permitindo que usuários identifiquem diferenças, validem suas bases de dados e gerem relatórios automaticamente.

## Sumário

- [Requisitos e Instalação](#requisitos-e-instalação)
- [Como Executar](#como-executar)
- [Estrutura de Pastas e Principais Arquivos](#estrutura-de-pastas-e-principais-arquivos)

## Requisitos e Instalação

Para executar o projeto, é necessário instalar as dependências listadas no arquivo `requirements.txt`. Utilize o seguinte comando:

```bash
pip install -r requirements.txt
```

As dependências incluem:

```
streamlit==1.32.0
pandas==2.2.2
XlsxWriter==3.2.0
openpyxl==3.1.2
xlrd==2.0.1
```

## Como Executar

Após instalar as dependências, execute o aplicativo com o comando:

```bash
streamlit run HOME.py
```

A página inicial será aberta no navegador, permitindo fazer upload das planilhas de cada serviço de delivery, realizar as comparações e baixar os relatórios gerados.

## Estrutura de Pastas e Principais Arquivos

A estrutura do projeto é a seguinte:

```
app_deliverys/
├── HOME.py
├── requirements.txt
├── pages_1_IFOOD.py
├── pages_2_TO_NO_LUCRO.py
├── pages_3_MAIS_DELIVERY.py
└── pages_4_AI_QUE_FOME.py
```

### Descrição dos principais arquivos:

- **HOME.py**: Página inicial com instruções gerais para o usuário.
- **requirements.txt**: Lista as dependências do projeto.
- **pages_1_IFOOD.py**: Script para comparação das planilhas do IFOOD.
- **pages_2_TO_NO_LUCRO.py**: Script para comparação das planilhas do Tonolucro.
- **pages_3_MAIS_DELIVERY.py**: Script para comparação das planilhas do Mais Delivery.
- **pages_4_AI_QUE_FOME.py**: Script para comparação das planilhas do AI QUE FOME.

---

Pronto para uso! Explore as páginas específicas para cada plataforma de delivery e otimize sua validação de dados com essa ferramenta.