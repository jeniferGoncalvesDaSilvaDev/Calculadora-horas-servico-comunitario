import streamlit as st
import pandas as pd
import numpy as np
import pytesseract
from PIL import Image
import re
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO

# Caminho do Tesseract para Windows (ajuste se estiver usando outro SO)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Layout e título cyberpunk
st.set_page_config(layout="wide", page_title="💀 Cyberpunk PSC Horas", page_icon="🕶️")

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

def extract_data_from_image(image):
    text = pytesseract.image_to_string(image, lang='por')
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

        horas_trabalhadas = (saida - entrada).total_seconds() / 3600

        if pd.notna(ini_intervalo) and pd.notna(fim_intervalo):
            intervalo = (fim_intervalo - ini_intervalo).total_seconds() / 3600
            horas_trabalhadas -= intervalo

        return max(0, round(horas_trabalhadas, 2))
    except:
        return 0

def main():
    st.title("💀 PSC Cyberpunk - Controle de Horas")
    st.subheader("Análise OCR e Visualização de Jornadas")

    uploaded_files = st.file_uploader("📸 Envie imagens do formulário PSC", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True)

    if uploaded_files:
        all_data = []
        total_geral = 0

        for uploaded_file in uploaded_files:
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

        if all_data:
            full_df = pd.concat(all_data)
            full_df['Data'] = pd.to_datetime(full_df['Data'], format='%d/%m/%Y')
            full_df.sort_values('Data', inplace=True)

            st.subheader("📊 Visualizações Cyberpunk")
            col1, col2 = st.columns(2)

            with col1:
                st.write("📅 **Horas Trabalhadas por Dia**")
                plt.figure(figsize=(10, 4))
                sns.set_theme(style="darkgrid")
                sns.barplot(x=full_df['Data'].dt.strftime('%d/%m'), y=full_df['Horas/Dia'], color='#00ffea')
                plt.axhline(full_df['Horas/Dia'].mean(), color='magenta', linestyle='--', label='Média')
                plt.xticks(rotation=45)
                plt.ylabel("Horas")
                plt.title("Horas por Dia")
                plt.legend()
                st.pyplot(plt)

            with col2:
                st.write("📈 **Distribuição das Horas**")
                plt.figure(figsize=(10, 4))
                sns.histplot(full_df['Horas/Dia'], bins=10, kde=True, color='lime')
                plt.axvline(full_df['Horas/Dia'].mean(), color='cyan', linestyle='--', label='Média')
                plt.axvline(full_df['Horas/Dia'].std() + full_df['Horas/Dia'].mean(), color='orange', linestyle=':', label='1 Desvio Padrão')
                plt.xlabel("Horas por Dia")
                plt.title("Distribuição das Horas")
                plt.legend()
                st.pyplot(plt)

            st.subheader("📚 Estatísticas Explicadas")
            media = full_df['Horas/Dia'].mean()
            desvio = full_df['Horas/Dia'].std()

            st.markdown(f"""
            - **Total Geral:** `{total_geral:.2f}` horas
            - **Média diária:** `{media:.2f}` horas  
              A média é o valor médio de horas que você trabalhou por dia.
            - **Desvio padrão:** `{desvio:.2f}` horas  
              O desvio padrão mostra quanto os valores de horas variam em relação à média.
              Quanto maior, mais instável está sua jornada diária.
            """)

            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                full_df.to_excel(writer, sheet_name='Detalhes', index=False)
                stats_df = pd.DataFrame({
                    'Total Geral': [total_geral],
                    'Média': [media],
                    'Desvio Padrão': [desvio]
                })
                stats_df.to_excel(writer, sheet_name='Resumo', index=False)

            output.seek(0)
            st.download_button(
                label="⬇️ Baixar Excel Consolidado",
                data=output,
                file_name="horas_psc_analise.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()