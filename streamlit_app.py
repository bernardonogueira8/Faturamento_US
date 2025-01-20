import pandas as pd
import streamlit as st
import openpyxl
from io import BytesIO

# Função para carregar a planilha
def carregar_planilha(file, sep=';', skiprows=0):
    try:
        planilha = pd.read_csv(file, sep=sep, header=None, skiprows=skiprows)
        return planilha
    except Exception as e:
        st.error(f"Erro ao carregar o arquivo: {e}")
        return None

# Função para transformar colunas em grupos de 7 e reorganizar em linhas
def transformar_colunas_em_linhas(df):
    colunas = df.values.flatten()  # Obtém todas as colunas em uma única lista
    grupos = [colunas[i:i + 7] for i in range(0, len(colunas), 7)]  # Divide em grupos de 7
    novo_df = pd.DataFrame(grupos)
    return novo_df

# Função para converter DataFrame em bytes para download
def to_excel_bytes(df):
    try:
        output = BytesIO()  # Criar um buffer de memória
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, header=False)  # Gravar o DataFrame no buffer
        output.seek(0)  # Retornar ao início do buffer
        return output.getvalue()  # Retornar os bytes do arquivo
    except Exception as e:
        st.error(f"Erro ao gerar arquivo Excel: {e}")
        return None


def main():
    st.title("Formatador Relatório")

    st.write("""
        Este aplicativo permite formatar arquivo texto para Excel.
    """)

    st.write("#### Selecione a planilha:")
    file = st.file_uploader("Clique para selecionar o arquivo de texto (somente .txt)", type=["txt"])

    if file:
        st.write(f"Arquivo carregado: {file.name}")
        
        # Carregar a planilha
        df = carregar_planilha(file)
        
        if df is not None:
            # Transformar as colunas em grupos de 7 e reorganizar como linhas
            resultado_df = transformar_colunas_em_linhas(df)
            
            # Exibir o DataFrame resultante
            st.write("Dados formatados:", resultado_df)
            
            # Gerar bytes do arquivo Excel
            excel_bytes = to_excel_bytes(resultado_df)
            
            if excel_bytes:
                # Botão para download do arquivo
                st.download_button(
                    label="Baixar Arquivo Resultante",
                    data=excel_bytes,
                    file_name="Faturamento_US.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == "__main__":
    main()
