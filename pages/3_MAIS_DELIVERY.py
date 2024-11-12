import streamlit as st
import pandas as pd
from io import BytesIO

st.title('Comparação de Planilhas Mais Delivery e Mais Delivery DB')

# Upload da planilha do Mais Delivery
uploaded_file_maisdelivery = st.file_uploader('Carregar planilha Mais Delivery', type=['xlsx'], key='maisdelivery')

# Upload da planilha Mais Delivery DB
uploaded_file_maisdeliverydb = st.file_uploader('Carregar planilha Mais Delivery DB', type=['xlsx'], key='maisdeliverydb')

# Verifica se ambos os arquivos foram carregados
if uploaded_file_maisdelivery is not None and uploaded_file_maisdeliverydb is not None:
    # Carrega as planilhas
    df_maisdelivery = pd.read_excel(uploaded_file_maisdelivery)
    df_maisdeliverydb = pd.read_excel(uploaded_file_maisdeliverydb)

    # Tratamento da planilha Mais Delivery
    # Manter apenas as colunas desejadas
    df_maisdelivery = df_maisdelivery[['Data Pedido', 'Número', 'Valor (R$)']]

    # Converter 'Data Pedido' para datetime com dayfirst=True
    df_maisdelivery['Data Pedido'] = pd.to_datetime(
        df_maisdelivery['Data Pedido'],
        dayfirst=True,
        errors='coerce'
    ).dt.date

    # Remover o símbolo 'R$' e converter 'Valor (R$)' para float
    df_maisdelivery['Valor (R$)'] = df_maisdelivery['Valor (R$)'].replace({'R\$': '', ',': '.', '\s+': ''}, regex=True)
    df_maisdelivery['Valor (R$)'] = pd.to_numeric(df_maisdelivery['Valor (R$)'], errors='coerce').fillna(0)

    # Tratamento da planilha Mais Delivery DB
    # Manter apenas as colunas desejadas
    df_maisdeliverydb = df_maisdeliverydb[['DATA', 'VALOR', 'ID PEDIDO']]

    # Converter 'DATA' para datetime com dayfirst=True
    df_maisdeliverydb['DATA'] = pd.to_datetime(
        df_maisdeliverydb['DATA'],
        dayfirst=True,
        errors='coerce'
    ).dt.date

    # Converter 'VALOR' para float
    df_maisdeliverydb['VALOR'] = pd.to_numeric(df_maisdeliverydb['VALOR'], errors='coerce').fillna(0)

    # Renomear colunas para evitar conflitos e facilitar o merge
    df_maisdelivery.rename(columns={'Data Pedido': 'Data', 'Valor (R$)': 'Valor Mais Delivery'}, inplace=True)
    df_maisdeliverydb.rename(columns={'DATA': 'Data', 'VALOR': 'Valor Mais Delivery DB'}, inplace=True)

    # Ordena os DataFrames por 'Data' e 'Valor' para consistência
    df_maisdelivery.sort_values(['Data', 'Valor Mais Delivery'], ascending=[True, True], inplace=True)
    df_maisdeliverydb.sort_values(['Data', 'Valor Mais Delivery DB'], ascending=[True, True], inplace=True)

    # Resetar os índices
    df_maisdelivery.reset_index(drop=True, inplace=True)
    df_maisdeliverydb.reset_index(drop=True, inplace=True)

    # Adicionar uma coluna para marcar se a linha foi correspondida
    df_maisdeliverydb['Correspondido'] = False

    # Lista para armazenar os resultados
    resultados = []

    # Correspondência por Data e Valor
    for i, row_maisdelivery in df_maisdelivery.iterrows():
        correspondencia = df_maisdeliverydb[
            (df_maisdeliverydb['Data'] == row_maisdelivery['Data']) &
            (df_maisdeliverydb['Valor Mais Delivery DB'] == row_maisdelivery['Valor Mais Delivery']) &
            (~df_maisdeliverydb['Correspondido'])
        ]

        if not correspondencia.empty:
            correspondencia = correspondencia.iloc[0]
            resultados.append((row_maisdelivery, correspondencia, 'Correspondente'))
            df_maisdeliverydb.at[correspondencia.name, 'Correspondido'] = True
        else:
            resultados.append((row_maisdelivery, pd.Series(), 'Diferença'))

    # Adicionar as linhas da planilha Mais Delivery DB que não foram correspondidas
    for j, row_maisdeliverydb in df_maisdeliverydb.iterrows():
        if not row_maisdeliverydb['Correspondido']:
            resultados.append((pd.Series(), row_maisdeliverydb, 'Diferença'))

    # Criar DataFrame final, mantendo todas as colunas das duas planilhas
    final_result = pd.DataFrame([{
        'Número do Pedido': row_maisdelivery['Número'] if 'Número' in row_maisdelivery else '',
        'Data Mais Delivery': row_maisdelivery['Data'].strftime('%d/%m/%Y') if 'Data' in row_maisdelivery and pd.notnull(row_maisdelivery['Data']) else '',
        'Valor Mais Delivery': row_maisdelivery['Valor Mais Delivery'] if 'Valor Mais Delivery' in row_maisdelivery else 0,
        'ID PEDIDO': row_maisdeliverydb['ID PEDIDO'] if 'ID PEDIDO' in row_maisdeliverydb else '',
        'Data Mais Delivery DB': row_maisdeliverydb['Data'].strftime('%d/%m/%Y') if 'Data' in row_maisdeliverydb and pd.notnull(row_maisdeliverydb['Data']) else '',
        'Valor Mais Delivery DB': row_maisdeliverydb['Valor Mais Delivery DB'] if 'Valor Mais Delivery DB' in row_maisdeliverydb else 0,
        'Discrepância Inicial': status
    } for row_maisdelivery, row_maisdeliverydb, status in resultados])

    # Converter as colunas de valores para numérico e substituir NaN por 0
    final_result['Valor Mais Delivery'] = pd.to_numeric(final_result['Valor Mais Delivery'], errors='coerce').fillna(0)
    final_result['Valor Mais Delivery DB'] = pd.to_numeric(final_result['Valor Mais Delivery DB'], errors='coerce').fillna(0)

    # Substituir NaN por vazio para melhor visualização nas outras colunas
    final_result.fillna('', inplace=True)

    # Converter o DataFrame final em um objeto BytesIO
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        final_result.to_excel(writer, index=False, sheet_name='Resultado')
        workbook = writer.book
        worksheet = writer.sheets['Resultado']

        # Definir a linha onde os totais serão escritos (após os dados)
        total_row = len(final_result) + 1  # +1 por causa do cabeçalho

        # Escrever 'Total' na coluna A na linha total
        worksheet.write(total_row, 0, 'Total')

        # Formatar as colunas de valores como moeda
        currency_format = workbook.add_format({'num_format': 'R$ #,##0.00'})

        # Definir as colunas que serão somadas
        colunas_soma = {
            'Valor Mais Delivery': final_result.columns.get_loc('Valor Mais Delivery'),
            'Valor Mais Delivery DB': final_result.columns.get_loc('Valor Mais Delivery DB'),
        }

        # Escrever as fórmulas de soma para as colunas de valores
        for label, col_index in colunas_soma.items():
            col_letter = chr(65 + col_index)
            formula = f"=SUM({col_letter}2:{col_letter}{total_row})"
            worksheet.write_formula(
                total_row,
                col_index,
                formula,
                currency_format
            )
            # Aplicar o formato de moeda às colunas
            worksheet.set_column(col_index, col_index, 15, currency_format)

        # Ajustar a largura das colunas para melhor visualização
        worksheet.set_column(0, len(final_result.columns) - 1, 18)

    # Obter o conteúdo do BytesIO
    processed_data = output.getvalue()

    # Botão para download da planilha consolidada
    st.download_button(
        label='Baixar planilha consolidada',
        data=processed_data,
        file_name='Resultado_Comparacao_MaisDelivery.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

    st.success("Comparação concluída e arquivo pronto para download.")