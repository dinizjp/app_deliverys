import streamlit as st
import pandas as pd
from io import BytesIO

st.title('Comparação de Planilhas Tonolucro e Tonolucro DB')

# Upload da planilha do Tonolucro
uploaded_file_tonolucro = st.file_uploader('Carregar planilha Tonolucro', type=['xlsx'], key='tonolucro')

# Upload da planilha Tonolucro DB
uploaded_file_tonolucrodb = st.file_uploader('Carregar planilha Tonolucro DB', type=['xlsx'], key='tonolucrodb')

# Verifica se ambos os arquivos foram carregados
if uploaded_file_tonolucro is not None and uploaded_file_tonolucrodb is not None:
    # Carrega as planilhas
    df_tonolucro = pd.read_excel(uploaded_file_tonolucro)
    df_tonolucrodb = pd.read_excel(uploaded_file_tonolucrodb)

    # Tratamento da planilha Tonolucro
    # Formata a coluna 'Data' para datetime no formato dia/mês/ano
    df_tonolucro['Data'] = pd.to_datetime(df_tonolucro['Data']).dt.strftime('%d/%m/%Y')

    # Remove o símbolo 'R$' e converte a coluna 'Valor' para float
    df_tonolucro['Valor'] = df_tonolucro['Valor'].replace({'R\$': '', ',': '.', '\s+': ''}, regex=True)
    df_tonolucro['Valor'] = df_tonolucro['Valor'].astype(float)

    # Manter apenas as colunas desejadas
    df_tonolucro = df_tonolucro[['Data', 'Número do pedido', 'Valor']]

    # Formata a coluna 'DATA' na planilha Tonolucro DB
    df_tonolucrodb['DATA'] = pd.to_datetime(df_tonolucrodb['DATA']).dt.strftime('%d/%m/%Y')

    # Converte a coluna 'Valor' para float na planilha Tonolucro DB
    df_tonolucrodb['VALOR'] = df_tonolucrodb['VALOR'].astype(float)

    # Renomear as colunas 'Valor' para distinguir entre as planilhas
    df_tonolucro.rename(columns={'Valor': 'Valor Tonolucro'}, inplace=True)
    df_tonolucrodb.rename(columns={'VALOR': 'Valor Tonolucro DB', 'DATA': 'Data'}, inplace=True)

    # Ordena os DataFrames por 'Valor' para consistência
    df_tonolucro.sort_values('Valor Tonolucro', ascending=True, inplace=True)
    df_tonolucrodb.sort_values('Valor Tonolucro DB', ascending=True, inplace=True)

    # Resetar os índices
    df_tonolucro = df_tonolucro.reset_index(drop=True)
    df_tonolucrodb = df_tonolucrodb.reset_index(drop=True)

    # Criar listas para armazenar os índices já utilizados
    indices_utilizados_tonolucro = []
    indices_utilizados_tonolucrodb = []

    # Lista para armazenar os resultados
    resultados = []

    # Loop sobre cada linha da planilha do Tonolucro
    for i, row_tonolucro in df_tonolucro.iterrows():
        # Encontrar correspondência na planilha Tonolucro DB
        correspondencia = df_tonolucrodb[
            (df_tonolucrodb['Data'] == row_tonolucro['Data']) &
            (df_tonolucrodb['Valor Tonolucro DB'] == row_tonolucro['Valor Tonolucro']) &
            (~df_tonolucrodb.index.isin(indices_utilizados_tonolucrodb))
        ].head(1)

        if not correspondencia.empty:
            resultados.append((row_tonolucro, correspondencia.iloc[0]))
            indices_utilizados_tonolucro.append(i)
            indices_utilizados_tonolucrodb.append(correspondencia.index[0])
        else:
            resultados.append((row_tonolucro, pd.Series()))
            indices_utilizados_tonolucro.append(i)

    # Adicionar as linhas da planilha Tonolucro DB que não foram correspondidas
    for j, row_tonolucrodb in df_tonolucrodb.iterrows():
        if j not in indices_utilizados_tonolucrodb:
            resultados.append((pd.Series(), row_tonolucrodb))
            indices_utilizados_tonolucrodb.append(j)

    # Criar DataFrame final, mantendo todas as colunas das duas planilhas
    final_result = pd.DataFrame([{
        **row_tonolucro.to_dict(),
        **row_tonolucrodb.to_dict()
    } for row_tonolucro, row_tonolucrodb in resultados])

    # Substituir NaN por vazio para melhor visualização
    final_result.fillna('', inplace=True)

    # Converter o DataFrame final em um objeto BytesIO
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        final_result.to_excel(writer, index=False, sheet_name='Resultado')
        workbook = writer.book
        worksheet = writer.sheets['Resultado']

        # Encontrar as posições das colunas 'Valor Tonolucro' e 'Valor Tonolucro DB'
        valor_col_tonolucro = final_result.columns.get_loc('Valor Tonolucro')
        valor_col_tonolucrodb = final_result.columns.get_loc('Valor Tonolucro DB')

        # Adicionar a coluna 'Diferença' após a última coluna
        diferenca_col = len(final_result.columns)
        worksheet.write(0, diferenca_col, 'Diferença')

        # Escrever a fórmula de diferença para cada linha
        for row_num in range(1, len(final_result) + 1):
            cell_formula = f"=IF(ISBLANK(${chr(65 + valor_col_tonolucrodb)}{row_num + 1})," \
                           f"${chr(65 + valor_col_tonolucro)}{row_num + 1}," \
                           f"${chr(65 + valor_col_tonolucro)}{row_num + 1}-${chr(65 + valor_col_tonolucrodb)}{row_num + 1})"
            worksheet.write_formula(row_num, diferenca_col, cell_formula)

        # Definir a linha onde os totais serão escritos (após os dados)
        total_row = len(final_result) + 1  # +1 por causa do cabeçalho

        # Escrever 'Total' na coluna A na linha total
        worksheet.write(total_row, 0, 'Total')

        # Formatar as colunas de valores como moeda
        currency_format = workbook.add_format({'num_format': 'R$ #,##0.00'})

        # Escrever as fórmulas de soma para as colunas de 'Valor' e 'Diferença'
        for col_index in [valor_col_tonolucro, valor_col_tonolucrodb]:
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
        file_name='Resultado_Comparacao_Tonolucro.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

    st.success("Comparação concluída e arquivo pronto para download.")