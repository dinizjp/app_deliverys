import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

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

    # Função personalizada para parsing de datas
    def parse_date(date_str):
        for fmt in ('%d/%m/%Y', '%m/%d/%Y', '%Y-%m-%d'):
            try:
                return datetime.strptime(date_str, fmt)
            except (ValueError, TypeError):
                continue
        return pd.NaT

    # --- Tratamento da Planilha Tonolucro ---
    # Manter apenas as colunas desejadas
    df_tonolucro = df_tonolucro[['Data', 'Número do pedido', 'Valor']]

    # Remover o horário da coluna 'Data' se houver e converter para datetime
    df_tonolucro['Data'] = df_tonolucro['Data'].astype(str).str.split(' ').str[0].apply(parse_date)

    # Verificar e informar datas inválidas
    datas_invalidas_tonolucro = df_tonolucro[df_tonolucro['Data'].isna()]
    if not datas_invalidas_tonolucro.empty:
        st.warning("Existem datas inválidas na planilha Tonolucro que não puderam ser convertidas:")
        st.write(datas_invalidas_tonolucro)

    # Formatar para 'dd/mm/yyyy' onde possível
    df_tonolucro['Data'] = df_tonolucro['Data'].dt.strftime('%d/%m/%Y')

    # Remover símbolos de moeda e converter 'Valor' para float
    for col in ['Valor']:
        df_tonolucro[col] = df_tonolucro[col].replace({'R\$': '', ',': '.', '\s+': ''}, regex=True)
        df_tonolucro[col] = df_tonolucro[col].astype(float)

    # Manter apenas as colunas desejadas (já feito acima)

    # --- Tratamento da Planilha Tonolucro DB ---
    # Manter apenas as colunas desejadas
    df_tonolucrodb = df_tonolucrodb[['DATA', 'VALOR', 'ID PEDIDO']]

    # Remover o horário da coluna 'DATA' se houver e converter para datetime
    df_tonolucrodb['DATA'] = df_tonolucrodb['DATA'].astype(str).str.split(' ').str[0].apply(parse_date)

    # Verificar e informar datas inválidas
    datas_invalidas_tonolucrodb = df_tonolucrodb[df_tonolucrodb['DATA'].isna()]
    if not datas_invalidas_tonolucrodb.empty:
        st.warning("Existem datas inválidas na planilha Tonolucro DB que não puderam ser convertidas:")
        st.write(datas_invalidas_tonolucrodb)

    # Formatar para 'dd/mm/yyyy' onde possível
    df_tonolucrodb['DATA'] = df_tonolucrodb['DATA'].dt.strftime('%d/%m/%Y')

    # Converter 'VALOR' para float
    df_tonolucrodb['VALOR'] = df_tonolucrodb['VALOR'].astype(float)

    # Renomear colunas para evitar conflitos e facilitar o merge
    # Renomeia 'DATA' para 'Data DB' para manter ambas as datas no resultado final
    df_tonolucrodb.rename(columns={'DATA': 'Data DB', 'VALOR': 'Valor Tonolucro DB'}, inplace=True)

    # Criar uma coluna 'Data' duplicada para uso no merge
    df_tonolucrodb['Data'] = df_tonolucrodb['Data DB']

    # Ordena os DataFrames por 'Valor' para consistência
    df_tonolucro.sort_values('Valor', ascending=True, inplace=True)
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
            (df_tonolucrodb['Valor Tonolucro DB'] == row_tonolucro['Valor']) &
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
        valor_col_tonolucro = final_result.columns.get_loc('Valor')
        valor_col_tonolucrodb = final_result.columns.get_loc('Valor Tonolucro DB')

        # Adicionar a coluna 'Diferença' após a última coluna
        diferenca_col = len(final_result.columns)
        worksheet.write(0, diferenca_col, 'Diferença')

        # Função para converter índice de coluna para letra(s) do Excel
        def col_letter(col_index):
            letter = ''
            while col_index >= 0:
                letter = chr(col_index % 26 + 65) + letter
                col_index = col_index // 26 - 1
            return letter

        # Escrever a fórmula de diferença para cada linha
        for row_num in range(1, len(final_result) + 1):
            col_tonolucrodb_letter = col_letter(valor_col_tonolucrodb)
            col_tonolucro_letter = col_letter(valor_col_tonolucro)
            cell_formula = f"=IF(ISBLANK({col_tonolucrodb_letter}{row_num + 1})," \
                           f"{col_tonolucro_letter}{row_num + 1}," \
                           f"{col_tonolucro_letter}{row_num + 1}-{col_tonolucrodb_letter}{row_num + 1})"
            worksheet.write_formula(row_num, diferenca_col, cell_formula)

        # Definir a linha onde os totais serão escritos (após os dados)
        total_row = len(final_result) + 1  # +1 por causa do cabeçalho

        # Escrever 'Total' na coluna A na linha total
        worksheet.write(total_row, 0, 'Total')

        # Formatar as colunas de valores como moeda
        currency_format = workbook.add_format({'num_format': 'R$ #,##0.00'})

        # Escrever as fórmulas de soma para as colunas de valores
        for col_index in [valor_col_tonolucro, valor_col_tonolucrodb]:
            col_letter_formula = col_letter(col_index)
            formula = f"=SUM({col_letter_formula}2:{col_letter_formula}{total_row})"
            worksheet.write_formula(
                total_row,
                col_index,
                formula,
                currency_format
            )
            # Aplicar o formato de moeda às colunas
            worksheet.set_column(col_index, col_index, 15, currency_format)

        # Soma dos valores absolutos para a coluna 'Diferença'
        col_letter_diferenca = col_letter(diferenca_col)
        formula_diferenca = f"=SUMPRODUCT(ABS({col_letter_diferenca}2:{col_letter_diferenca}{total_row}))"
        worksheet.write_formula(
            total_row,
            diferenca_col,
            formula_diferenca,
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
