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

    # Converter a coluna 'DATA' para datetime e normalizar (remover hora)
    df_ifood['DATA'] = pd.to_datetime(df_ifood['DATA'], dayfirst=True).dt.date
    df_ifooddb['DATA'] = pd.to_datetime(df_ifooddb['DATA'], dayfirst=True).dt.date

    # Renomear as colunas para facilitar a mesclagem
    df_ifood.rename(columns={'VALOR DOS ITENS': 'VALOR_IFOOD'}, inplace=True)
    df_ifooddb.rename(columns={'VALOR': 'VALOR_IfoodDB'}, inplace=True)

    # Ordena os DataFrames por 'DATA' e 'VALOR' para consistência
    df_ifood.sort_values(['DATA', 'VALOR_IFOOD'], ascending=[True, True], inplace=True)
    df_ifooddb.sort_values(['DATA', 'VALOR_IfoodDB'], ascending=[True, True], inplace=True)

    # Resetar os índices
    df_ifood.reset_index(drop=True, inplace=True)
    df_ifooddb.reset_index(drop=True, inplace=True)

    # Adicionar uma coluna para marcar se a linha foi correspondida
    df_ifooddb['Correspondido'] = False

    # Lista para armazenar os resultados
    resultados = []

    # Correspondência por Data e Valor
    for i, row_ifood in df_ifood.iterrows():
        correspondencia = df_ifooddb[
            (df_ifooddb['DATA'] == row_ifood['DATA']) &
            (df_ifooddb['VALOR_IfoodDB'] == row_ifood['VALOR_IFOOD']) &
            (~df_ifooddb['Correspondido'])
        ]

        if not correspondencia.empty:
            correspondencia = correspondencia.iloc[0]
            resultados.append((row_ifood, correspondencia))
            df_ifooddb.at[correspondencia.name, 'Correspondido'] = True
        else:
            resultados.append((row_ifood, pd.Series()))

    # Adicionar as linhas da planilha IfoodDB que não foram correspondidas
    for j, row_ifooddb in df_ifooddb.iterrows():
        if not row_ifooddb['Correspondido']:
            resultados.append((pd.Series(), row_ifooddb))

    # Criar DataFrame final, mantendo todas as colunas das duas planilhas
    final_result = pd.DataFrame([{
        'N° PEDIDO IFOOD': row_ifood['N° PEDIDO'] if 'N° PEDIDO' in row_ifood else '',
        'DATA IFOOD': row_ifood['DATA'].strftime('%d/%m/%Y') if 'DATA' in row_ifood and pd.notnull(row_ifood['DATA']) else '',
        'VALOR IFOOD': row_ifood['VALOR_IFOOD'] if 'VALOR_IFOOD' in row_ifood else 0,
        'N° PEDIDO IfoodDB': row_ifooddb['N° PEDIDO'] if 'N° PEDIDO' in row_ifooddb else '',
        'DATA IfoodDB': row_ifooddb['DATA'].strftime('%d/%m/%Y') if 'DATA' in row_ifooddb and pd.notnull(row_ifooddb['DATA']) else '',
        'VALOR IfoodDB': row_ifooddb['VALOR_IfoodDB'] if 'VALOR_IfoodDB' in row_ifooddb else 0,
    } for row_ifood, row_ifooddb in resultados])

    # Converter as colunas de valores para numérico e substituir NaN por 0
    final_result['VALOR IFOOD'] = pd.to_numeric(final_result['VALOR IFOOD'], errors='coerce').fillna(0)
    final_result['VALOR IfoodDB'] = pd.to_numeric(final_result['VALOR IfoodDB'], errors='coerce').fillna(0)

    # Adicionar a coluna 'Discrepância Inicial' com base na lógica fornecida
    final_result['Discrepância Inicial'] = final_result.apply(
        lambda row: 'Diferença' if row['VALOR IFOOD'] == 0 or row['VALOR IfoodDB'] == 0 else 'Correspondente',
        axis=1
    )

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

        # Definir as colunas que serão somadas
        colunas_soma = {
            'VALOR IFOOD': final_result.columns.get_loc('VALOR IFOOD'),
            'VALOR IfoodDB': final_result.columns.get_loc('VALOR IfoodDB'),
        }

        # Formatar as colunas de valores como moeda
        currency_format = workbook.add_format({'num_format': 'R$ #,##0.00'})

        # Escrever as fórmulas de soma
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