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

    # Converter 'Data Pedido' para datetime com dayfirst=True e formatar para dia/mês/ano
    df_maisdelivery['Data Pedido'] = pd.to_datetime(
        df_maisdelivery['Data Pedido'],
        dayfirst=True,
        errors='coerce'
    ).dt.strftime('%d/%m/%Y')

    # Remover o símbolo 'R$' e converter 'Valor (R$)' para float
    df_maisdelivery['Valor (R$)'] = df_maisdelivery['Valor (R$)'].replace({'R\$': '', ',': '.', '\s+': ''}, regex=True)
    df_maisdelivery['Valor (R$)'] = df_maisdelivery['Valor (R$)'].astype(float)

    # Tratamento da planilha Mais Delivery DB
    # Manter apenas as colunas desejadas
    df_maisdeliverydb = df_maisdeliverydb[['DATA', 'VALOR', 'ID PEDIDO']]

    # Converter 'DATA' para datetime com dayfirst=True e formatar para dia/mês/ano
    df_maisdeliverydb['DATA'] = pd.to_datetime(
        df_maisdeliverydb['DATA'],
        dayfirst=True,
        errors='coerce'
    ).dt.strftime('%d/%m/%Y')

    # Converter 'VALOR' para float
    df_maisdeliverydb['VALOR'] = df_maisdeliverydb['VALOR'].astype(float)

    # Renomear colunas para evitar conflitos e facilitar o merge
    df_maisdelivery.rename(columns={'Data Pedido': 'Data', 'Valor (R$)': 'Valor Mais Delivery'}, inplace=True)
    df_maisdeliverydb.rename(columns={'DATA': 'Data', 'VALOR': 'Valor Mais Delivery DB'}, inplace=True)
    # Ordena os DataFrames por 'Valor' para consistência
    df_maisdelivery.sort_values('Valor Mais Delivery', ascending=True, inplace=True)
    df_maisdeliverydb.sort_values('Valor Mais Delivery DB', ascending=True, inplace=True)

    # Resetar os índices
    df_maisdelivery = df_maisdelivery.reset_index(drop=True)
    df_maisdeliverydb = df_maisdeliverydb.reset_index(drop=True)

    # Criar listas para armazenar os índices já utilizados
    indices_utilizados_maisdelivery = []
    indices_utilizados_maisdeliverydb = []

    # Lista para armazenar os resultados
    resultados = []

    # Loop sobre cada linha da planilha Mais Delivery
    for i, row_maisdelivery in df_maisdelivery.iterrows():
        # Encontrar correspondência na planilha Mais Delivery DB
        correspondencia = df_maisdeliverydb[
            (df_maisdeliverydb['Data'] == row_maisdelivery['Data']) &
            (df_maisdeliverydb['Valor Mais Delivery DB'] == row_maisdelivery['Valor Mais Delivery']) &
            (~df_maisdeliverydb.index.isin(indices_utilizados_maisdeliverydb))
        ].head(1)

        if not correspondencia.empty:
            resultados.append((row_maisdelivery, correspondencia.iloc[0]))
            indices_utilizados_maisdelivery.append(i)
            indices_utilizados_maisdeliverydb.append(correspondencia.index[0])
        else:
            resultados.append((row_maisdelivery, pd.Series()))
            indices_utilizados_maisdelivery.append(i)

    # Adicionar as linhas da planilha Mais Delivery DB que não foram correspondidas
    for j, row_maisdeliverydb in df_maisdeliverydb.iterrows():
        if j not in indices_utilizados_maisdeliverydb:
            resultados.append((pd.Series(), row_maisdeliverydb))
            indices_utilizados_maisdeliverydb.append(j)

    # Criar DataFrame final, mantendo todas as colunas das duas planilhas
    final_result = pd.DataFrame([{
        **row_maisdelivery.to_dict(),
        **row_maisdeliverydb.to_dict()
    } for row_maisdelivery, row_maisdeliverydb in resultados])

    # Substituir NaN por vazio para melhor visualização
    final_result.fillna('', inplace=True)

    # Converter o DataFrame final em um objeto BytesIO
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        final_result.to_excel(writer, index=False, sheet_name='Resultado')
        workbook = writer.book
        worksheet = writer.sheets['Resultado']

        # Encontrar as posições das colunas de valores
        valor_col_maisdelivery = final_result.columns.get_loc('Valor Mais Delivery')
        valor_col_maisdeliverydb = final_result.columns.get_loc('Valor Mais Delivery DB')

        # Adicionar a coluna 'Diferença' após a última coluna
        diferenca_col = len(final_result.columns)
        worksheet.write(0, diferenca_col, 'Diferença')

        # Escrever a fórmula de diferença para cada linha
        for row_num in range(1, len(final_result) + 1):
            cell_formula = f"=IF(ISBLANK(${chr(65 + valor_col_maisdeliverydb)}{row_num + 1})," \
                           f"${chr(65 + valor_col_maisdelivery)}{row_num + 1}," \
                           f"${chr(65 + valor_col_maisdelivery)}{row_num + 1}-${chr(65 + valor_col_maisdeliverydb)}{row_num + 1})"
            worksheet.write_formula(row_num, diferenca_col, cell_formula)

        # Definir a linha onde os totais serão escritos (após os dados)
        total_row = len(final_result) + 1  # +1 por causa do cabeçalho

        # Escrever 'Total' na coluna A na linha total
        worksheet.write(total_row, 0, 'Total')

        # Formatar as colunas de valores como moeda
        currency_format = workbook.add_format({'num_format': 'R$ #,##0.00'})

        # Escrever as fórmulas de soma para as colunas de valores
        for col_index in [valor_col_maisdelivery, valor_col_maisdeliverydb]:
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

        # Soma dos valores absolutos para a coluna 'Diferença'
        col_letter = chr(65 + diferenca_col)
        formula = f"=SUMPRODUCT(ABS({col_letter}2:{col_letter}{total_row}))"
        worksheet.write_formula(
            total_row,
            diferenca_col,
            formula,
            currency_format
        )
        # Aplicar o formato de moeda à coluna 'Diferença'
        worksheet.set_column(diferenca_col, diferenca_col, 15, currency_format)

        # Ajustar a largura das colunas para melhor visualização
        worksheet.set_column(0, len(final_result.columns), 18)

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
