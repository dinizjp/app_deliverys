import streamlit as st
import pandas as pd
from io import BytesIO

st.title('Comparação de Planilhas IFOOD e IfoodDB')

# Upload da planilha do IFOOD
uploaded_file_ifood = st.file_uploader('Carregar planilha IFOOD', type=['xlsx'], key='ifood')

# Upload da planilha IfoodDB
uploaded_file_ifooddb = st.file_uploader('Carregar planilha IFOOD_DB', type=['xlsx'], key='ifooddb')

# Verifica se ambos os arquivos foram carregados
if uploaded_file_ifood is not None and uploaded_file_ifooddb is not None:
    # Carrega as planilhas
    df_ifood = pd.read_excel(uploaded_file_ifood)
    df_ifooddb = pd.read_excel(uploaded_file_ifooddb)

    # Seleciona apenas as colunas desejadas na planilha do IFOOD
    colunas_desejadas_ifood = ['N° PEDIDO', 'DATA', 'VALOR DOS ITENS']
    df_ifood = df_ifood[colunas_desejadas_ifood]

    # Formata a coluna 'DATA' para extrair apenas a data (dia/mês/ano)
    df_ifood['DATA'] = pd.to_datetime(df_ifood['DATA']).dt.strftime('%d/%m/%Y')

    # Formata a coluna 'DATA' na planilha IfoodDB
    df_ifooddb['DATA'] = pd.to_datetime(df_ifooddb['DATA']).dt.strftime('%d/%m/%Y')

    # Ordena os DataFrames por 'VALOR DOS ITENS' para consistência
    df_ifood.sort_values('VALOR DOS ITENS', ascending=True, inplace=True)
    df_ifooddb.sort_values('VALOR', ascending=True, inplace=True)

    # Resetar os índices
    df_ifood = df_ifood.reset_index(drop=True)
    df_ifooddb = df_ifooddb.reset_index(drop=True)

    # Criar listas para armazenar os índices já utilizados
    indices_utilizados_ifood = []
    indices_utilizados_ifooddb = []

    # Lista para armazenar os resultados
    resultados = []

    # Loop sobre cada linha da planilha do IFOOD
    for i, row_ifood in df_ifood.iterrows():
        # Encontrar correspondência na planilha IfoodDB
        correspondencia = df_ifooddb[
            (df_ifooddb['DATA'] == row_ifood['DATA']) &
            (df_ifooddb['VALOR'] == row_ifood['VALOR DOS ITENS']) &
            (~df_ifooddb.index.isin(indices_utilizados_ifooddb))
        ].head(1)

        if not correspondencia.empty:
            resultados.append((row_ifood, correspondencia.iloc[0]))
            indices_utilizados_ifood.append(i)
            indices_utilizados_ifooddb.append(correspondencia.index[0])
        else:
            resultados.append((row_ifood, pd.Series()))
            indices_utilizados_ifood.append(i)

    # Adicionar as linhas da planilha IfoodDB que não foram correspondidas
    for j, row_ifooddb in df_ifooddb.iterrows():
        if j not in indices_utilizados_ifooddb:
            resultados.append((pd.Series(), row_ifooddb))
            indices_utilizados_ifooddb.append(j)

    # Criar DataFrame final, mantendo todas as colunas das duas planilhas
    final_result = pd.DataFrame([{
        **row_ifood.to_dict(),
        **row_ifooddb.to_dict()
    } for row_ifood, row_ifooddb in resultados])

    # Substituir NaN por vazio para melhor visualização
    final_result.fillna('', inplace=True)

    # Garantir que a coluna 'DATA' esteja no formato correto no resultado final
    if 'DATA' in final_result.columns:
        final_result['DATA'] = final_result['DATA'].astype(str)

    # Converter o DataFrame final em um objeto BytesIO
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        final_result.to_excel(writer, index=False, sheet_name='Resultado')
        workbook = writer.book
        worksheet = writer.sheets['Resultado']

        # Encontrar as posições das colunas 'VALOR DOS ITENS' e 'VALOR'
        valor_itens_col = final_result.columns.get_loc('VALOR DOS ITENS')
        valor_db_col = final_result.columns.get_loc('VALOR')

        # Adicionar a coluna 'Diferença' após a última coluna
        diferenca_col = len(final_result.columns)
        worksheet.write(0, diferenca_col, 'Diferença')

        # Escrever a fórmula de diferença para cada linha
        for row_num in range(1, len(final_result) + 1):
            cell_formula = f"=IF(ISBLANK(${chr(65 + valor_db_col)}{row_num + 1})," \
                           f"${chr(65 + valor_itens_col)}{row_num + 1}," \
                           f"${chr(65 + valor_itens_col)}{row_num + 1}-${chr(65 + valor_db_col)}{row_num + 1})"
            worksheet.write_formula(row_num, diferenca_col, cell_formula)

        # Definir a linha onde os totais serão escritos (após os dados)
        total_row = len(final_result) + 1  # +1 por causa do cabeçalho

        # Escrever 'Total' na coluna A na linha total
        worksheet.write(total_row, 0, 'Total')

        # Formatar as colunas de valores como moeda
        currency_format = workbook.add_format({'num_format': 'R$ #,##0.00'})

        # Escrever as fórmulas de soma para 'VALOR DOS ITENS' e 'VALOR'
        for col_label, col_index in [('VALOR DOS ITENS', valor_itens_col), ('VALOR', valor_db_col)]:
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

        # Escrever a fórmula de soma dos valores absolutos da coluna 'Diferença'
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
        file_name='Resultado_Comparacao_IFOOD_IfoodDB.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

    st.success("Comparação concluída e arquivo pronto para download.")