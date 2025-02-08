import streamlit as st
import time
from leitura_arquivo import ler_planilha, personalizar_planilha

st.title('Planilha Terra Util ST')

def main():
    planilha = st.file_uploader('Importar Planilha', type='xlsx')

    if planilha is not None:
        if st.button('Importar'):
            with st.status("Realizando importação da Planilha ...", expanded=True) as status:

                time.sleep(1)
                st.write('Lendo a planilha ...')

                time.sleep(1)
                st.write('Cálculo Base ICMS...')
                
                time.sleep(1)
                st.write('Verificando convênio e protocolo de icms ...')
                
                time.sleep(1)
                st.write('Verificando substituto tributario operações internas ...')
                
                time.sleep(1)
                st.write('Verificando cálculo esta correto ...')
                
                time.sleep(1)
                st.write('Verificando MVA ...')
                
                time.sleep(1)
                st.write('Verificando CFOP ...')
                
                time.sleep(1)
                st.write('Realizando Análise ...')

                df = ler_planilha(planilha=planilha)
                
                time.sleep(1)
                st.write('Aplicando a formatação na planilha ...')
                excel_bytes = personalizar_planilha(df)

                time.sleep(2)


            status.update(
                label="Verificação Concluida!", state="complete", expanded=False
            )
        
            st.download_button(
                    label="Baixar Excel",
                    data=excel_bytes,
                    file_name=f"planilha_st.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    icon=":material/download:"
                )





if __name__ == '__main__':
    main()