import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from utils import (
    get_db_connection, run_system_query, process_ifood, process_mais_delivery,
    process_aiquefome, process_goomer, process_edvania, consolidate_data
)

st.title("Consolidação de Deliverys e Sistema")

# Mapeamento de empresas (lojas)
id_empresa_mapping = {
    58: 'Araguaína II',
    53: 'Imperatriz II', 
    50: 'Balsas I', 
    56: 'Gurupi I',
    46: 'Formosa I',
    59: 'Guaraí'
}

# Mapeamento dos deliverys e regras de negócio
mapping_deliverys = {
    1032: "IFood.COM AG REST ONLINE S.A.",
    1035: "MAIS DELIVERY ARAGUAINA LTDA",
    1231: "AI QUE FOME",
    709:  "GOOMER",
    202:  "IFood.COM",
    1205: "EDVANIA SOBRINHO DA SILVA"
}

# Relação de deliverys permitidos por loja
allowed_deliveries = {
    58: [1032, 1035, 1231, 709],
    53: [202],
    50: [1032],
    56: [1032, 1231],
    46: [1032],
    59: [1205]
}

# Regras para filtrar os dados do sistema (para cada delivery)
delivery_rules = {
    1032: {"client_ids": [1032], "id_forma": [17]}, 
    1035: {"client_ids": [1035], "id_forma": [18]},
    1231: {"client_ids": [1231], "id_forma": [37]},
    709:  {"client_ids": [709], "id_forma": [1,31,5,6]},
    202:  {"client_ids": [202], "id_forma": [17]},
    1205: {"client_ids": [1205], "id_forma": [1,31,5,6]}
}

# Interface de filtros na sidebar
st.sidebar.header("Filtros de Consolidação")
selected_company_id = st.sidebar.selectbox("Selecione a Empresa", 
                                             options=list(id_empresa_mapping.keys()),
                                             format_func=lambda x: id_empresa_mapping[x])
allowed = allowed_deliveries.get(selected_company_id, [])
delivery_options = {d: mapping_deliverys[d] for d in allowed}
selected_delivery = st.sidebar.selectbox("Selecione o Delivery", 
                                           options=list(delivery_options.keys()),
                                           format_func=lambda x: delivery_options[x])
start_date = st.sidebar.date_input("Data Inicial", value=datetime.today())
end_date = st.sidebar.date_input("Data Final", value=datetime.today())

st.write(f"**Empresa selecionada:** {id_empresa_mapping[selected_company_id]}")
st.write(f"**Delivery selecionado:** {mapping_deliverys[selected_delivery]}")

# Upload da planilha do delivery
uploaded_file = st.file_uploader("Carregar planilha do delivery", type=["xlsx", "xls"])

if st.button("Consolidar"):
    if not uploaded_file:
        st.error("Carregue a planilha do delivery.")
    else:
        st.info("Executando consulta no sistema...")
        # Obter os dados do sistema conforme os filtros e regras do delivery
        drules = delivery_rules[selected_delivery]
        system_df = run_system_query(start_date, end_date, selected_company_id, drules["client_ids"], drules["id_forma"])
        st.write("**Dados do sistema obtidos:**")
        st.dataframe(system_df.head())
        
        # Processar a planilha do delivery conforme o tipo selecionado
        if selected_delivery in [1032, 202]:
            delivery_df = process_ifood(uploaded_file)
        elif selected_delivery == 1035:
            delivery_df = process_mais_delivery(uploaded_file)
        elif selected_delivery == 1231:
            delivery_df = process_aiquefome(uploaded_file)
        elif selected_delivery == 709:
            delivery_df = process_goomer(uploaded_file)
        elif selected_delivery == 1205:
            delivery_df = process_edvania(uploaded_file)
        else:
            st.error("Delivery não suportado.")
            st.stop()
        
        st.write("**Dados do delivery:**")
        st.dataframe(delivery_df.head())
        
        # Definir os mapeamentos para as colunas de consolidação (delivery x sistema)
        if selected_delivery in [1032, 202]:
            order_delivery_cols = {
                'Pedido Delivery': 'N° PEDIDO IFOOD',
                'Data Delivery': 'DATA IFOOD',
                'Valor Delivery': 'VALOR IFOOD'
            }
        elif selected_delivery == 1035:
            order_delivery_cols = {
                'Pedido Delivery': 'N° PEDIDO MAIS DELIVERY',
                'Data Delivery': 'DATA MAIS DELIVERY',
                'Valor Delivery': 'VALOR MAIS DELIVERY'
            }
        elif selected_delivery == 1231:
            order_delivery_cols = {
                'Pedido Delivery': 'N° PEDIDO AI QUE FOME',
                'Data Delivery': 'DATA AI QUE FOME',
                'Valor Delivery': 'Valor AI QUE FOME'
            }
        elif selected_delivery == 709:
            order_delivery_cols = {
                'Pedido Delivery': 'N° PEDIDO GOOMER',
                'Data Delivery': 'DATA GOOMER',
                'Valor Delivery': 'VALOR GOOMER'
            }
        elif selected_delivery == 1205:
            order_delivery_cols = {
                'Pedido Delivery': 'N° PEDIDO MAIS DELIVERY',
                'Data Delivery': 'DATA MAIS DELIVERY',
                'Valor Delivery': 'VALOR MAIS DELIVERY'
            }
        else:
            st.error("Mapping de colunas não definido para este delivery.")
            st.stop()
        
        order_system_cols = {
            'ID Venda Sistema': 'ID_Venda',
            'Data Sistema': 'Data_Faturamento',
            'Valor Sistema': 'Valor Bruto'
        }
        
        # Definir as colunas de data e valor para a consolidação
        if selected_delivery in [1032, 202]:
            date_col_delivery = 'DATA IFOOD'
            value_col_delivery = 'VALOR IFOOD'
        elif selected_delivery == 1035:
            date_col_delivery = 'DATA MAIS DELIVERY'
            value_col_delivery = 'VALOR MAIS DELIVERY'
        elif selected_delivery == 1231:
            date_col_delivery = 'DATA AI QUE FOME'
            value_col_delivery = 'Valor AI QUE FOME'
        elif selected_delivery == 709:
            date_col_delivery = 'DATA GOOMER'
            value_col_delivery = 'VALOR GOOMER'
        elif selected_delivery == 1205:
            date_col_delivery = 'DATA MAIS DELIVERY'
            value_col_delivery = 'VALOR MAIS DELIVERY'
        else:
            st.error("Não foi possível definir as colunas para consolidação.")
            st.stop()
        date_col_system = 'Data_Faturamento'
        value_col_system = 'Valor Bruto'
        
        consolidated_df = consolidate_data(delivery_df.copy(), system_df.copy(),
                                             date_col_delivery, value_col_delivery,
                                             date_col_system, value_col_system,
                                             order_delivery_cols, order_system_cols)
        
        st.write("**Resultado da Consolidação:**")
        st.dataframe(consolidated_df.head())
        
        # Gerar arquivo Excel para download
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            consolidated_df.to_excel(writer, index=False, sheet_name='Consolidado')
        processed_data = output.getvalue()
        st.download_button(
            label='Baixar planilha consolidada',
            data=processed_data,
            file_name=f"Consolidado_{mapping_deliverys[selected_delivery]}.xlsx",
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
