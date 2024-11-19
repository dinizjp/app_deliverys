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

    # Remover símbolos de moeda e converter 'Valor' para float
    for col in ['Valor']:
        df_tonolucro[col] = df_tonolucro[col].replace({'R\$': '', ',': '.', '\s+': ''}, regex=True)
        df_tonolucro[col] = pd.to_numeric(df_tonolucro[col], errors='coerce').fillna(0)

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

    # Converter 'VALOR' para float
    df_tonolucrodb['VALOR'] = pd.to_numeric(df_tonolucrodb['VALOR'], errors='coerce').fillna(0)

    # Renomear colunas para evitar conflitos e facilitar o merge
    df_tonolucrodb.rename(columns={'DATA': 'Data DB', 'VALOR': 'Valor Tonolucro DB'}, inplace=True)

    # Criar uma coluna 'Data' duplicada para uso no merge e ordenação
    df_tonolucrodb['Data'] = df_tonolucrodb['Data DB']

    # --- Ordenação dos DataFrames por Data e Valor ---
    df_tonolucro.sort_values(['Data', 'Valor'], ascending=[True, True], inplace=True)
    df_tonolucrodb.sort_values(['Data', 'Valor Tonolucro DB'], ascending=[True, True], inplace=True)

    # Resetar os índices
    df_tonolucro = df_tonolucro.reset_index(drop=True)
    df_tonolucrodb = df_tonolucrodb.reset_index(drop=True)

    # Adicionar uma coluna para marcar se a linha foi correspondida
    df_tonolucrodb['Correspondido'] = False

    # Lista para armazenar os resultados
    resultados = []

    # Correspondência por Data e Valor
    for i, row_tonolucro in df_tonolucro.iterrows():
        correspondencia = df_tonolucrodb[
            (df_tonolucrodb['Data'] == row_tonolucro['Data']) &
            (df_tonolucrodb['Valor Tonolucro DB'] == row_tonolucro['Valor']) &
            (~df_tonolucrodb['Correspondido'])
        ]

        if not correspondencia.empty:
            correspondencia = correspondencia.iloc[0]
            resultados.append((row_tonolucro, correspondencia, 'Correspondente'))
            df_tonolucrodb.at[correspondencia.name, 'Correspondido'] = True
        else:
            resultados.append((row_tonolucro, pd.Series(), 'Diferença'))

    # Adicionar as linhas da planilha Tonolucro DB que não foram correspondidas
    for j, row_tonolucrodb in df_tonolucrodb.iterrows():
        if not row_tonolucrodb['Correspondido']:
            resultados.append((pd.Series(), row_tonolucrodb, 'Diferença'))

    # Criar DataFrame final, mantendo todas as colunas das duas planilhas
    final_result = pd.DataFrame([{
        'Número do pedido': row_tonolucro['Número do pedido'] if 'Número do pedido' in row_tonolucro else '',
        'Data Tonolucro': row_tonolucro['Data'] if 'Data' in row_tonolucro and pd.notnull(row_tonolucro['Data']) else '',
        'Valor Tonolucro': row_tonolucro['Valor'] if 'Valor' in row_tonolucro else 0,
        'ID PEDIDO': row_tonolucrodb['ID PEDIDO'] if 'ID PEDIDO' in row_tonolucrodb else '',
        'Data Tonolucro DB': row_tonolucrodb['Data DB'] if 'Data DB' in row_tonolucrodb and pd.notnull(row_tonolucrodb['Data DB']) else '',
        'Valor Tonolucro DB': row_tonolucrodb['Valor Tonolucro DB'] if 'Valor Tonolucro DB' in row_tonolucrodb else 0,
        'Discrepância Inicial': status
    } for row_tonolucro, row_tonolucrodb, status in resultados])

    # **Criar coluna auxiliar para ordenação por data**
    final_result['Data Ordenacao'] = final_result.apply(
        lambda row: row['Data Tonolucro'] if row['Data Tonolucro'] != '' else row['Data Tonolucro DB'],
        axis=1
    )

    # Converter 'Data Ordenacao' para datetime
    final_result['Data Ordenacao'] = pd.to_datetime(final_result['Data Ordenacao'], errors='coerce')

    # Ordenar o DataFrame final pela 'Data Ordenacao'
    final_result.sort_values('Data Ordenacao', inplace=True)
    final_result.reset_index(drop=True, inplace=True)

    # Converter as colunas de valores para numérico e substituir NaN por 0
    final_result['Valor Tonolucro'] = pd.to_numeric(final_result['Valor Tonolucro'], errors='coerce').fillna(0)
    final_result['Valor Tonolucro DB'] = pd.to_numeric(final_result['Valor Tonolucro DB'], errors='coerce').fillna(0)

    # **Formatar as datas para string no formato 'dd/mm/yyyy'**
    final_result['Data Tonolucro'] = final_result['Data Tonolucro'].apply(
        lambda x: x.strftime('%d/%m/%Y') if pd.notnull(x) and x != '' else ''
    )
    final_result['Data Tonolucro DB'] = final_result['Data Tonolucro DB'].apply(
        lambda x: x.strftime('%d/%m/%Y') if pd.notnull(x) and x != '' else ''
    )

    # Substituir NaN por vazio para melhor visualização nas outras colunas
    final_result.fillna('', inplace=True)

    # Remover a coluna auxiliar 'Data Ordenacao'
    final_result.drop(columns=['Data Ordenacao'], inplace=True)

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
            'Valor Tonolucro': final_result.columns.get_loc('Valor Tonolucro'),
            'Valor Tonolucro DB': final_result.columns.get_loc('Valor Tonolucro DB'),
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
        file_name='Resultado_Comparacao_Tonolucro.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

    st.success("Comparação concluída e arquivo pronto para download.")