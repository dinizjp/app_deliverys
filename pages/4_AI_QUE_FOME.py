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
            except (ValueError, TypeError):
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
    df_aiquefome['Data'] = df_aiquefome['Data'].dt.date

    # Remover símbolos de moeda e converter 'Total (R$)' e 'Desconto (R$)' para float
    for col in ['Total (R$)', 'Desconto (R$)']:
        df_aiquefome[col] = df_aiquefome[col].replace({'R\$': '', ',': '.', '\s+': ''}, regex=True)
        df_aiquefome[col] = pd.to_numeric(df_aiquefome[col], errors='coerce').fillna(0)

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
    df_aiquefomedb['DATA'] = df_aiquefomedb['DATA'].dt.date

    # Converter 'VALOR' para float
    df_aiquefomedb['VALOR'] = pd.to_numeric(df_aiquefomedb['VALOR'], errors='coerce').fillna(0)

    # Renomear colunas para evitar conflitos e facilitar o merge
    df_aiquefomedb.rename(columns={'DATA': 'Data DB', 'VALOR': 'Valor AI QUE FOME DB'}, inplace=True)

    # Criar uma coluna 'Data' duplicada para uso no merge
    df_aiquefomedb['Data'] = df_aiquefomedb['Data DB']

    # Ordena os DataFrames por 'Data' e 'Valor' para consistência
    df_aiquefome.sort_values(['Data', 'Valor AI QUE FOME'], ascending=[True, True], inplace=True)
    df_aiquefomedb.sort_values(['Data', 'Valor AI QUE FOME DB'], ascending=[True, True], inplace=True)

    # Resetar os índices
    df_aiquefome = df_aiquefome.reset_index(drop=True)
    df_aiquefomedb = df_aiquefomedb.reset_index(drop=True)

    # Adicionar uma coluna para marcar se a linha foi correspondida
    df_aiquefomedb['Correspondido'] = False

    # Lista para armazenar os resultados
    resultados = []

    # Correspondência por Data e Valor
    for i, row_aiquefome in df_aiquefome.iterrows():
        correspondencia = df_aiquefomedb[
            (df_aiquefomedb['Data'] == row_aiquefome['Data']) &
            (df_aiquefomedb['Valor AI QUE FOME DB'] == row_aiquefome['Valor AI QUE FOME']) &
            (~df_aiquefomedb['Correspondido'])
        ]

        if not correspondencia.empty:
            correspondencia = correspondencia.iloc[0]
            resultados.append((row_aiquefome, correspondencia, 'Correspondente'))
            df_aiquefomedb.at[correspondencia.name, 'Correspondido'] = True
        else:
            resultados.append((row_aiquefome, pd.Series(), 'Diferença'))

    # Adicionar as linhas da planilha AI QUE FOME DB que não foram correspondidas
    for j, row_aiquefomedb in df_aiquefomedb.iterrows():
        if not row_aiquefomedb['Correspondido']:
            resultados.append((pd.Series(), row_aiquefomedb, 'Diferença'))

    # Criar DataFrame final, mantendo todas as colunas das duas planilhas
    final_result = pd.DataFrame([{
        'Nro. Pedido': row_aiquefome['Nro. Pedido'] if 'Nro. Pedido' in row_aiquefome else '',
        'Data AI QUE FOME': row_aiquefome['Data'].strftime('%d/%m/%Y') if 'Data' in row_aiquefome and pd.notnull(row_aiquefome['Data']) else '',
        'Valor AI QUE FOME': row_aiquefome['Valor AI QUE FOME'] if 'Valor AI QUE FOME' in row_aiquefome else 0,
        'ID PEDIDO': row_aiquefomedb['ID PEDIDO'] if 'ID PEDIDO' in row_aiquefomedb else '',
        'Data AI QUE FOME DB': row_aiquefomedb['Data DB'].strftime('%d/%m/%Y') if 'Data DB' in row_aiquefomedb and pd.notnull(row_aiquefomedb['Data DB']) else '',
        'Valor AI QUE FOME DB': row_aiquefomedb['Valor AI QUE FOME DB'] if 'Valor AI QUE FOME DB' in row_aiquefomedb else 0,
        'Discrepância Inicial': status
    } for row_aiquefome, row_aiquefomedb, status in resultados])

    # Converter as colunas de valores para numérico e substituir NaN por 0
    final_result['Valor AI QUE FOME'] = pd.to_numeric(final_result['Valor AI QUE FOME'], errors='coerce').fillna(0)
    final_result['Valor AI QUE FOME DB'] = pd.to_numeric(final_result['Valor AI QUE FOME DB'], errors='coerce').fillna(0)

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
            'Valor AI QUE FOME': final_result.columns.get_loc('Valor AI QUE FOME'),
            'Valor AI QUE FOME DB': final_result.columns.get_loc('Valor AI QUE FOME DB'),
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
        file_name='Resultado_Comparacao_AI_QUE_FOME.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

    st.success("Comparação concluída e arquivo pronto para download.")