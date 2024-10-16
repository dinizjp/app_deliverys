import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.title('Comparação de Planilhas AI QUE FOME e AI QUE FOME DB')

# Upload da planilha do AI QUE FOME
uploaded_file_aiquefome = st.file_uploader('Carregar planilha AI QUE FOME', type=['xls'], key='aiquefome')

# Upload da planilha AI QUE FOME DB
uploaded_file_aiquefomedb = st.file_uploader('Carregar planilha AI QUE FOME DB', type=['xlsx'], key='aiquefomedb')

# Verifica se ambos os arquivos foram carregados
if uploaded_file_aiquefome is not None and uploaded_file_aiquefomedb is not None:
    # Carrega as planilhas
    df_aiquefome = pd.read_excel(uploaded_file_aiquefome)
    df_aiquefomedb = pd.read_excel(uploaded_file_aiquefomedb)

    # Função personalizada para parsing de datas
    def parse_date(date_str):
        for fmt in ('%d/%m/%Y', '%m/%d/%Y', '%Y-%m-%d'):
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue
        return pd.NaT

    # --- Tratamento da Planilha AI QUE FOME ---
    # Manter apenas as colunas desejadas
    df_aiquefome = df_aiquefome[['Nro. Pedido', 'Data', 'Total (R$)', 'Desconto (R$)']]

    # Remover o horário da coluna 'Data' se houver e converter para datetime
    df_aiquefome['Data'] = df_aiquefome['Data'].astype(str).str.split(' ').str[0].apply(parse_date)

    # Verificar e informar datas inválidas
    datas_invalidas_aiquefome = df_aiquefome[df_aiquefome['Data'].isna()]
    if not datas_invalidas_aiquefome.empty:
        st.warning("Existem datas inválidas na planilha AI QUE FOME que não puderam ser convertidas:")
        st.write(datas_invalidas_aiquefome)

    # Formatar para 'dd/mm/yyyy' onde possível
    df_aiquefome['Data'] = df_aiquefome['Data'].dt.strftime('%d/%m/%Y')

    # Remover símbolos de moeda e converter 'Total (R$)' e 'Desconto (R$)' para float
    for col in ['Total (R$)', 'Desconto (R$)']:
        df_aiquefome[col] = df_aiquefome[col].replace({'R\$': '', ',': '.', '\s+': ''}, regex=True)
        df_aiquefome[col] = df_aiquefome[col].astype(float)

    # Substituir NaN em 'Desconto (R$)' por 0
    df_aiquefome['Desconto (R$)'] = df_aiquefome['Desconto (R$)'].fillna(0)

    # Criar a coluna 'Valor AI QUE FOME' somando 'Total (R$)' e 'Desconto (R$)'
    df_aiquefome['Valor AI QUE FOME'] = df_aiquefome['Total (R$)'] + df_aiquefome['Desconto (R$)']

    # --- Tratamento da Planilha AI QUE FOME DB ---
    # Manter apenas as colunas desejadas
    df_aiquefomedb = df_aiquefomedb[['DATA', 'VALOR', 'ID PEDIDO']]

    # Remover o horário da coluna 'DATA' se houver e converter para datetime
    df_aiquefomedb['DATA'] = df_aiquefomedb['DATA'].astype(str).str.split(' ').str[0].apply(parse_date)

    # Verificar e informar datas inválidas
    datas_invalidas_aiquefomedb = df_aiquefomedb[df_aiquefomedb['DATA'].isna()]
    if not datas_invalidas_aiquefomedb.empty:
        st.warning("Existem datas inválidas na planilha AI QUE FOME DB que não puderam ser convertidas:")
        st.write(datas_invalidas_aiquefomedb)

    # Formatar para 'dd/mm/yyyy' onde possível
    df_aiquefomedb['DATA'] = df_aiquefomedb['DATA'].dt.strftime('%d/%m/%Y')

    # Converter 'VALOR' para float
    df_aiquefomedb['VALOR'] = df_aiquefomedb['VALOR'].astype(float)

    # Renomear colunas para evitar conflitos e facilitar o merge
    # Renomeia 'DATA' para 'Data DB' para manter ambas as datas no resultado final
    df_aiquefomedb.rename(columns={'DATA': 'Data DB', 'VALOR': 'Valor AI QUE FOME DB'}, inplace=True)

    # Criar uma coluna 'Data' duplicada para uso no merge
    df_aiquefomedb['Data'] = df_aiquefomedb['Data DB']

    # Ordena os DataFrames por 'Valor' para consistência
    df_aiquefome.sort_values('Valor AI QUE FOME', ascending=True, inplace=True)
    df_aiquefomedb.sort_values('Valor AI QUE FOME DB', ascending=True, inplace=True)

    # Resetar os índices
    df_aiquefome = df_aiquefome.reset_index(drop=True)
    df_aiquefomedb = df_aiquefomedb.reset_index(drop=True)

    # Criar listas para armazenar os índices já utilizados
    indices_utilizados_aiquefome = []
    indices_utilizados_aiquefomedb = []

    # Lista para armazenar os resultados
    resultados = []

    # Loop sobre cada linha da planilha AI QUE FOME
    for i, row_aiquefome in df_aiquefome.iterrows():
        # Encontrar correspondência na planilha AI QUE FOME DB
        correspondencia = df_aiquefomedb[
            (df_aiquefomedb['Data'] == row_aiquefome['Data']) &
            (df_aiquefomedb['Valor AI QUE FOME DB'] == row_aiquefome['Valor AI QUE FOME']) &
            (~df_aiquefomedb.index.isin(indices_utilizados_aiquefomedb))
        ].head(1)

        if not correspondencia.empty:
            resultados.append((row_aiquefome, correspondencia.iloc[0]))
            indices_utilizados_aiquefome.append(i)
            indices_utilizados_aiquefomedb.append(correspondencia.index[0])
        else:
            resultados.append((row_aiquefome, pd.Series()))
            indices_utilizados_aiquefome.append(i)

    # Adicionar as linhas da planilha AI QUE FOME DB que não foram correspondidas
    for j, row_aiquefomedb in df_aiquefomedb.iterrows():
        if j not in indices_utilizados_aiquefomedb:
            resultados.append((pd.Series(), row_aiquefomedb))
            indices_utilizados_aiquefomedb.append(j)

    # Criar DataFrame final, mantendo todas as colunas das duas planilhas
    final_result = pd.DataFrame([{
        **row_aiquefome.to_dict(),
        **row_aiquefomedb.to_dict()
    } for row_aiquefome, row_aiquefomedb in resultados])

    # Substituir NaN por vazio para melhor visualização
    final_result.fillna('', inplace=True)

    # Converter o DataFrame final em um objeto BytesIO
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        final_result.to_excel(writer, index=False, sheet_name='Resultado')
        workbook = writer.book
        worksheet = writer.sheets['Resultado']

        # Encontrar as posições das colunas de valores
        valor_col_aiquefome = final_result.columns.get_loc('Valor AI QUE FOME')
        valor_col_aiquefomedb = final_result.columns.get_loc('Valor AI QUE FOME DB')

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
            col_aiquefomedb_letter = col_letter(valor_col_aiquefomedb)
            col_aiquefome_letter = col_letter(valor_col_aiquefome)
            cell_formula = f"=IF(ISBLANK({col_aiquefomedb_letter}{row_num + 1})," \
                           f"{col_aiquefome_letter}{row_num + 1}," \
                           f"{col_aiquefome_letter}{row_num + 1}-{col_aiquefomedb_letter}{row_num + 1})"
            worksheet.write_formula(row_num, diferenca_col, cell_formula)

        # Definir a linha onde os totais serão escritos (após os dados)
        total_row = len(final_result) + 1  # +1 por causa do cabeçalho

        # Escrever 'Total' na coluna A na linha total
        worksheet.write(total_row, 0, 'Total')

        # Formatar as colunas de valores como moeda
        currency_format = workbook.add_format({'num_format': 'R$ #,##0.00'})

        # Escrever as fórmulas de soma para as colunas de valores
        for col_index in [valor_col_aiquefome, valor_col_aiquefomedb]:
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
        file_name='Resultado_Comparacao_AI_QUE_FOME.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

    st.success("Comparação concluída e arquivo pronto para download.")
