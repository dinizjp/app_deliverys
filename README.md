Aqui está o README.md atualizado com base nas informações do projeto `app_deliverys` e em todos os arquivos disponíveis:

```markdown
# app_deliverys

## Resumo do projeto

O projeto *app_deliverys* é uma ferramenta destinada a consolidar e comparar planilhas de diferentes serviços de delivery. A aplicação foi desenvolvida para facilitar a análise dos pedidos, valores e discrepâncias entre várias bases de dados de delivery, permitindo que usuários identifiquem diferenças, validem suas bases de dados e gerem relatórios automaticamente.

## Objetivo

Automatizar a consolidação e comparação de planilhas provenientes de diferentes plataformas de delivery, como IFOOD, Tonolucro, AI QUE FOME e Mais Delivery. Os usuários poderão fazer upload das planilhas, realizarem as comparações e download dos relatórios gerados de forma eficiente.

## Dependências e Instalação

Para executar o projeto, você deve instalar as dependências listadas no arquivo `requirements.txt`. Utilize o seguinte comando:

```bash
pip install -r requirements.txt
```

O arquivo `requirements.txt` inclui as seguintes bibliotecas:

```
streamlit==1.32.0
pandas==2.2.2
XlsxWriter==3.2.0
openpyxl==3.1.2
xlrd==2.0.1
```

## Como executar e testar

Após instalar as dependências, você pode executar o aplicativo com o seguinte comando:

```bash
streamlit run HOME.py
```

A interface do aplicativo será aberta no seu navegador. Nele, você poderá fazer upload das planilhas correspondentes a cada serviço de delivery, realizar as comparações e baixar os relatórios gerados.

## Estrutura de pastas e descrição dos arquivos principais

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

### Descrição dos arquivos principais:

- **HOME.py**: Página inicial do aplicativo, onde os usuários podem acessar os diferentes módulos de comparação de planilhas.
  
- **requirements.txt**: Lista das dependências necessárias para a aplicação.

- **pages_1_IFOOD.py**: Script que contém a lógica de comparação para as planilhas do IFOOD.

- **pages_2_TO_NO_LUCRO.py**: Script que contém a lógica de comparação para as planilhas do Tonolucro.

- **pages_3_MAIS_DELIVERY.py**: Script que contém a lógica de comparação para as planilhas do Mais Delivery.

- **pages_4_AI_QUE_FOME.py**: Script que contém a lógica de comparação para as planilhas do AI QUE FOME.

---

Pronto para uso! Explore as páginas para cada plataforma de delivery e simplifique a validação de suas informações com esta ferramenta.
```

Esse README.md fornece uma visão clara do projeto, instruções sobre instalação e execução, além de uma descrição concisa dos arquivos no repositório. Se houver mais alguma informação ou detalhe que você gostaria de adicionar, sinta-se à vontade para me avisar!