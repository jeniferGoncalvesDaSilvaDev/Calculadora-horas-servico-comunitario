import streamlit as st
import pandas as pd
import numpy as np
import requests
from PIL import Image
import re
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO
import xlsxwriter

OCR_SPACE_API_KEY = ''  # Se tiver, coloque aqui. Se não, deixe vazio.

# Layout e título cyberpunk
st.set_page_config(layout="wide", page_title=" Cálculo Horas", page_icon="🕶️")

# Estilo cyberpunk (usando HTML e CSS simples)
st.markdown("""
    <style>
    html, body, [class*="css"] {
        background-color: #0f0f0f !important;
        color: #00ffea !important;
        font-family: 'Courier New', monospace;
    }
    .stButton>button {
        background-color: #440099;
        color: white;
        border-radius: 10px;
    }
    .stDownloadButton>button {
        background-color: #008080;
        color: white;
    }
    .css-1offfwp {
        background-color: #1a1a1a !important;
    }
    </style>
""", unsafe_allow_html=True)

def ocr_space_api(image_bytes):
    """Envia a imagem para OCR.space e retorna o texto extraído."""
    if not OCR_SPACE_API_KEY:
        # Retorna dados de exemplo para demonstração
        return """
        04/12/2024
        08:00 12:00 13:00 17:00
        05/12/2024  
        08:30 12:00 13:00 17:30
        06/12/2024
        09:00 12:00 13:00 18:00
        """
    
    try:
        payload = {
            'isOverlayRequired': False,
            'apikey': OCR_SPACE_API_KEY,
            'language': 'por',
        }
        files = {
            'file': ('image.png', image_bytes),
        }
        response = requests.post('https://api.ocr.space/parse/image', data=payload, files=files)
        result = response.json()
        if result['IsErroredOnProcessing']:
            return ''
        parsed_results = result.get("ParsedResults")
        if parsed_results:
            return parsed_results[0]['ParsedText']
        return ''
    except Exception as e:
        st.error(f"Erro no OCR: {str(e)}")
        return ''

def extract_data_from_image(image):
    img_byte_arr = BytesIO()
    image.save(img_byte_arr, format='PNG')
    img_bytes = img_byte_arr.getvalue()
    text = ocr_space_api(img_bytes)
    return text

def parse_time_data(text):
    date_pattern = r'\d{1,2}/\d{1,2}/\d{4}'
    time_pattern = r'\d{1,2}:\d{2}'

    dates = re.findall(date_pattern, text)
    times = re.findall(time_pattern, text)

    data = []
    for i in range(0, len(dates)):
        if i * 4 + 3 < len(times):
            entry = {
                'Data': dates[i],
                'Entrada': times[i*4],
                'Início Intervalo': times[i*4+1],
                'Fim Intervalo': times[i*4+2],
                'Saída': times[i*4+3]
            }
            data.append(entry)
    return data

def calculate_hours(row):
    try:
        entrada = pd.to_datetime(row['Entrada'], format='%H:%M', errors='coerce')
        saida = pd.to_datetime(row['Saída'], format='%H:%M', errors='coerce')
        ini_intervalo = pd.to_datetime(row['Início Intervalo'], format='%H:%M', errors='coerce')
        fim_intervalo = pd.to_datetime(row['Fim Intervalo'], format='%H:%M', errors='coerce')

        if pd.isna(entrada) or pd.isna(saida):
            return 0

        # Horas trabalhadas = (Saída1 - Entrada1) + (Saída2 - Entrada2)
        # Saída1 = Início do Intervalo, Entrada1 = Entrada
        # Saída2 = Saída, Entrada2 = Fim do Intervalo
        
        if pd.notna(ini_intervalo) and pd.notna(fim_intervalo):
            periodo1 = (ini_intervalo - entrada).total_seconds() / 3600  # Manhã
            periodo2 = (saida - fim_intervalo).total_seconds() / 3600    # Tarde
            horas_trabalhadas = periodo1 + periodo2
        else:
            # Se não há intervalo, calcula normalmente
            horas_trabalhadas = (saida - entrada).total_seconds() / 3600

        return max(0, round(horas_trabalhadas, 2))
    except:
        return 0

def main():
    st.title("Controle de Horas")
    st.subheader("Análise OCR e Visualização de Jornadas")

    st.markdown("### 📁 Upload de Arquivos")
    st.info("💡 Dica: Você pode selecionar múltiplos arquivos de uma vez!")
    uploaded_files = st.file_uploader(
        "📸 Selecione uma ou mais imagens do formulário PSC", 
        type=['jpg', 'jpeg', 'png'], 
        accept_multiple_files=True,
        help="Formatos aceitos: JPG, JPEG, PNG. Você pode selecionar múltiplos arquivos."
    )
    
    if uploaded_files:
        st.success(f"✅ {len(uploaded_files)} arquivo(s) carregado(s) com sucesso!")

    if uploaded_files:
        all_data = []
        total_geral = 0
        
        # Barra de progresso para múltiplos arquivos
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for i, uploaded_file in enumerate(uploaded_files):
            # Atualizar progresso
            progress = (i + 1) / len(uploaded_files)
            progress_bar.progress(progress)
            status_text.text(f"Processando arquivo {i+1} de {len(uploaded_files)}: {uploaded_file.name}")
            image = Image.open(uploaded_file)
            st.image(image, caption=uploaded_file.name, width=300)
            text = extract_data_from_image(image)
            raw_data = parse_time_data(text)

            if not raw_data:
                st.warning(f"Nenhum dado detectado em {uploaded_file.name}")
                continue

            df = pd.DataFrame(raw_data)
            df['Horas/Dia'] = df.apply(calculate_hours, axis=1)

            total = df['Horas/Dia'].sum()
            total_geral += total
            all_data.append(df)

            st.write(f"**{uploaded_file.name}** - Total: `{total:.2f}` horas 🕒")
            st.dataframe(df)
        
        # Finalizar barra de progresso
        progress_bar.progress(1.0)
        status_text.text(f"✅ Processamento completo! {len(uploaded_files)} arquivo(s) processado(s).")

        if all_data:
            # Processar dados automaticamente
            st.success("✅ Processamento concluído! Preparando download do Excel...")
            full_df = pd.concat(all_data)
            full_df['Data'] = pd.to_datetime(full_df['Data'], format='%d/%m/%Y')
            full_df.sort_values('Data', inplace=True)

            # Calcular estatísticas
            media = full_df['Horas/Dia'].mean()
            desvio = full_df['Horas/Dia'].std()
            
            # Calcular total mensal
            full_df['Mes_Ano'] = full_df['Data'].dt.to_period('M')
            monthly_totals = full_df.groupby('Mes_Ano')['Horas/Dia'].sum().reset_index()
            monthly_totals['Mes_Ano'] = monthly_totals['Mes_Ano'].astype(str)
            monthly_totals_display = full_df.groupby('Mes_Ano')['Horas/Dia'].sum()
            
            # Gerar Excel automaticamente
            try:
                output = BytesIO()
                
                # Preparar DataFrame para Excel
                excel_df = full_df.copy()
                excel_df['Data'] = excel_df['Data'].dt.strftime('%d/%m/%Y')
                excel_df = excel_df.drop('Mes_Ano', axis=1, errors='ignore')
                
                # Usar pandas para criar Excel
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    # Aba Detalhes
                    excel_df.to_excel(writer, sheet_name='Detalhes', index=False)
                    
                    # Aba Resumo
                    stats_df = pd.DataFrame([
                        ['Total Geral', f"{total_geral:.2f}"],
                        ['Média Diária', f"{media:.2f}"],
                        ['Desvio Padrão', f"{desvio:.2f}"]
                    ], columns=['Estatística', 'Valor'])
                    stats_df.to_excel(writer, sheet_name='Resumo', index=False)
                    
                    # Aba Totais Mensais
                    monthly_totals_df = monthly_totals.copy()
                    monthly_totals_df.columns = ['Mês/Ano', 'Total Horas']
                    monthly_totals_df['Total Horas'] = monthly_totals_df['Total Horas'].round(2)
                    monthly_totals_df.to_excel(writer, sheet_name='Totais Mensais', index=False)
                
                output.seek(0)
                
            except Exception as e:
                st.error(f"Erro ao gerar Excel: {str(e)}")
                output = None
            
            # Download automático
            if output:
                st.download_button(
                    label="📊 DOWNLOAD AUTOMÁTICO - Excel Consolidado",
                    data=output,
                    file_name="horas_psc_analise.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

            st.subheader("📊 Visualizações")
            col1, col2 = st.columns(2)

            with col1:
                st.write("📅 **Horas Trabalhadas por Dia**")
                try:
                    plt.figure(figsize=(10, 4))
                    plt.style.use('dark_background')
                    dates_str = full_df['Data'].dt.strftime('%d/%m')
                    plt.bar(dates_str, full_df['Horas/Dia'], color='#00ffea')
                    plt.axhline(full_df['Horas/Dia'].mean(), color='magenta', linestyle='--', label='Média')
                    plt.xticks(rotation=45)
                    plt.ylabel("Horas")
                    plt.title("Horas por Dia")
                    plt.legend()
                    plt.tight_layout()
                    st.pyplot(plt)
                except Exception as e:
                    st.error(f"Erro no gráfico: {str(e)}")

            with col2:
                st.write("📈 **Distribuição das Horas**")
                try:
                    plt.figure(figsize=(10, 4))
                    plt.style.use('dark_background')
                    plt.hist(full_df['Horas/Dia'], bins=10, alpha=0.7, color='lime', edgecolor='white')
                    plt.axvline(full_df['Horas/Dia'].mean(), color='cyan', linestyle='--', label='Média')
                    plt.axvline(full_df['Horas/Dia'].std() + full_df['Horas/Dia'].mean(), color='orange', linestyle=':', label='1 Desvio Padrão')
                    plt.xlabel("Horas por Dia")
                    plt.ylabel("Frequência")
                    plt.title("Distribuição das Horas")
                    plt.legend()
                    plt.tight_layout()
                    st.pyplot(plt)
                except Exception as e:
                    st.error(f"Erro no gráfico: {str(e)}")

            st.subheader("📚 Estatísticas Explicadas")

            st.markdown(f"""
            - **Total Geral:** `{total_geral:.2f}` horas
            - **Média diária:** `{media:.2f}` horas  
              A média é o valor médio de horas que você trabalhou por dia.
            - **Desvio padrão:** `{desvio:.2f}` horas  
              O desvio padrão mostra quanto os valores de horas variam em relação à média.
              Quanto maior, mais instável está sua jornada diária.
            """)
            
            st.subheader("📅 Totais por Mês")
            for mes, total_mes in monthly_totals_display.items():
                st.write(f"**{mes}:** `{total_mes:.2f}` horas")

if __name__ == "__main__":
    main()