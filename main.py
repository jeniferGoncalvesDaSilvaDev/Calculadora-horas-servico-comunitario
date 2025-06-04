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
import openpyxl
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from scipy import stats

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

        # Horas trabalhadas = (Horário de saída1 - Horário de entrada1) + (Horário de saída2 - Horário de entrada2) - (hora fim do intervalo - hora inicio do intervalo)
        # Saída1 = Início do Intervalo, Entrada1 = Entrada
        # Saída2 = Saída, Entrada2 = Fim do Intervalo
        
        if pd.notna(ini_intervalo) and pd.notna(fim_intervalo):
            periodo1 = (ini_intervalo - entrada).total_seconds() / 3600  # Manhã
            periodo2 = (saida - fim_intervalo).total_seconds() / 3600    # Tarde
            tempo_intervalo = (fim_intervalo - ini_intervalo).total_seconds() / 3600  # Intervalo
            horas_trabalhadas = periodo1 + periodo2 - tempo_intervalo
        else:
            # Se não há intervalo, calcula normalmente
            horas_trabalhadas = (saida - entrada).total_seconds() / 3600

        return max(0, round(horas_trabalhadas, 2))
    except:
        return 0

def process_excel_file(excel_file):
    """Processa um arquivo Excel e extrai os dados de horário"""
    try:
        # Tenta ler o arquivo Excel
        df = pd.read_excel(excel_file)
        
        # Busca por colunas que possam conter os dados necessários
        data_processed = []
        
        # Mapear possíveis nomes de colunas
        col_mapping = {
            'data': ['data', 'date', 'dia'],
            'entrada1': ['horário de entrada 1', 'entrada 1', 'entrada1', 'entry1'],
            'saida1': ['horário de saída 1', 'saida 1', 'saida1', 'exit1'],
            'entrada2': ['horário de entrada 2', 'entrada 2', 'entrada2', 'entry2'],
            'saida2': ['horário de saída 2', 'saida 2', 'saida2', 'exit2'],
            'entrada': ['entrada', 'entry', 'inicio', 'start'],
            'saida': ['saida', 'exit', 'fim', 'end'],
            'inicio_intervalo': ['inicio_intervalo', 'inicio intervalo', 'break_start', 'intervalo_inicio'],
            'fim_intervalo': ['fim_intervalo', 'fim intervalo', 'break_end', 'intervalo_fim']
        }
        
        # Encontrar as colunas corretas
        found_cols = {}
        for key, possible_names in col_mapping.items():
            for col in df.columns:
                if any(name.lower() in col.lower() for name in possible_names):
                    found_cols[key] = col
                    break
        
        # Verificar se é o formato com 4 colunas (entrada1, saida1, entrada2, saida2)
        if 'entrada1' in found_cols and 'saida2' in found_cols:
            # Formato com horários separados
            for i, row in df.iterrows():
                try:
                    entrada1 = str(row[found_cols['entrada1']]).strip() if 'entrada1' in found_cols else ''
                    saida1 = str(row[found_cols.get('saida1', '')]).strip() if 'saida1' in found_cols else ''
                    entrada2 = str(row[found_cols.get('entrada2', '')]).strip() if 'entrada2' in found_cols else ''
                    saida2 = str(row[found_cols['saida2']]).strip() if 'saida2' in found_cols else ''
                    
                    # Verificar se tem dados válidos
                    if entrada1 and saida2 and entrada1 != 'nan' and saida2 != 'nan':
                        data_entry = {
                            'Data': f"{i+1:02d}/12/2024",  # Data fictícia
                            'Entrada': entrada1,
                            'Saída': saida2,
                            'Início Intervalo': saida1 if saida1 and saida1 != 'nan' else '',
                            'Fim Intervalo': entrada2 if entrada2 and entrada2 != 'nan' else ''
                        }
                        data_processed.append(data_entry)
                        
                except Exception as e:
                    continue
        else:
            # Formato tradicional
            # Se não encontrar as colunas esperadas, assumir ordem padrão
            if len(found_cols) < 3:  # Pelo menos Data, Entrada e Saída
                cols = df.columns.tolist()
                if len(cols) >= 3:
                    found_cols = {
                        'data': cols[0] if len(cols) > 4 else None,
                        'entrada': cols[0] if len(cols) <= 4 else cols[1],
                        'saida': cols[-1]  # Última coluna como saída
                    }
                    if len(cols) >= 5:
                        found_cols['inicio_intervalo'] = cols[2]
                        found_cols['fim_intervalo'] = cols[3]
            
            # Processar os dados
            for i, row in df.iterrows():
                try:
                    if 'data' in found_cols and found_cols['data']:
                        data_val = str(row[found_cols['data']])
                    else:
                        data_val = f"{i+1:02d}/12/2024"  # Data fictícia se não houver coluna de data
                    
                    data_entry = {
                        'Data': data_val,
                        'Entrada': str(row[found_cols['entrada']]) if 'entrada' in found_cols else '',
                        'Saída': str(row[found_cols['saida']]) if 'saida' in found_cols else '',
                        'Início Intervalo': str(row[found_cols.get('inicio_intervalo', '')]) if 'inicio_intervalo' in found_cols else '',
                        'Fim Intervalo': str(row[found_cols.get('fim_intervalo', '')]) if 'fim_intervalo' in found_cols else ''
                    }
                    
                    # Validar se a linha tem dados válidos
                    if data_entry['Entrada'] and data_entry['Saída'] and data_entry['Entrada'] != 'nan' and data_entry['Saída'] != 'nan':
                        data_processed.append(data_entry)
                        
                except Exception as e:
                    continue
                
        return data_processed
        
    except Exception as e:
        st.error(f"Erro ao processar arquivo Excel: {str(e)}")
        return []

def main():
    st.title("Controle de Horas")
    st.subheader("Análise OCR e Visualização de Jornadas")

    st.markdown("### 📁 Upload de Arquivos")
    
    # Opção para escolher tipo de arquivo
    file_type = st.radio(
        "🔧 Escolha o tipo de arquivo:",
        ("📸 Imagens (OCR)", "📊 Excel"),
        help="Selecione o tipo de arquivo que você deseja enviar"
    )
    
    if file_type == "📸 Imagens (OCR)":
        st.info("💡 Dica: Você pode selecionar múltiplas imagens de uma vez!")
        uploaded_files = st.file_uploader(
            "📸 Selecione uma ou mais imagens do formulário PSC", 
            type=['jpg', 'jpeg', 'png'], 
            accept_multiple_files=True,
            help="Formatos aceitos: JPG, JPEG, PNG. Você pode selecionar múltiplos arquivos."
        )
        file_processing_type = "image"
    else:
        st.info("💡 Dica: Você pode selecionar múltiplos arquivos Excel de uma vez!")
        uploaded_files = st.file_uploader(
            "📊 Selecione um ou mais arquivos Excel com dados de horário", 
            type=['xlsx', 'xls'], 
            accept_multiple_files=True,
            help="Formatos aceitos: XLSX, XLS. Os dados devem ter colunas para Data, Entrada, Saída e opcionalmente intervalos."
        )
        file_processing_type = "excel"
    
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
            
            if file_processing_type == "image":
                # Processamento de imagem (OCR)
                image = Image.open(uploaded_file)
                st.image(image, caption=uploaded_file.name, width=300)
                text = extract_data_from_image(image)
                raw_data = parse_time_data(text)
            else:
                # Processamento de Excel
                raw_data = process_excel_file(uploaded_file)
                st.write(f"📊 **{uploaded_file.name}** processado")

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
            
            # Converter data com tratamento de erro
            try:
                full_df['Data'] = pd.to_datetime(full_df['Data'], format='%d/%m/%Y', errors='coerce')
                # Remove linhas com datas inválidas
                full_df = full_df.dropna(subset=['Data'])
                full_df.sort_values('Data', inplace=True)
            except Exception as e:
                st.warning(f"Erro na conversão de datas: {str(e)}. Usando dados sem ordenação por data.")
                # Se não conseguir converter, manter como está

            # Calcular estatísticas DIÁRIAS
            media_diaria = full_df['Horas/Dia'].mean()
            desvio_diario = full_df['Horas/Dia'].std()
            # Calcular moda (valor mais frequente)
            moda_diaria = stats.mode(full_df['Horas/Dia'], keepdims=True)[0][0] if len(full_df['Horas/Dia']) > 0 else 0
            
            # Calcular estatísticas SEMANAIS
            full_df['Semana'] = full_df['Data'].dt.to_period('W')
            weekly_totals = full_df.groupby('Semana')['Horas/Dia'].sum()
            media_semanal = weekly_totals.mean()
            desvio_semanal = weekly_totals.std()
            
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
                        ['Média Diária', f"{media_diaria:.2f}"],
                        ['Moda Diária', f"{moda_diaria:.2f}"],
                        ['Desvio Padrão Diário', f"{desvio_diario:.2f}"],
                        ['Média Semanal', f"{media_semanal:.2f}"],
                        ['Desvio Padrão Semanal', f"{desvio_semanal:.2f}"]
                    ], columns=['Estatística', 'Valor'])
                    stats_df.to_excel(writer, sheet_name='Resumo', index=False)
                    
                    # Aba Totais Semanais
                    weekly_totals_df = weekly_totals.reset_index()
                    weekly_totals_df['Semana'] = weekly_totals_df['Semana'].astype(str)
                    weekly_totals_df.columns = ['Semana', 'Total Horas']
                    weekly_totals_df['Total Horas'] = weekly_totals_df['Total Horas'].round(2)
                    weekly_totals_df.to_excel(writer, sheet_name='Totais Semanais', index=False)
                    
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

            st.subheader("📊 Visualizações Interativas")
            
            # Gráfico 1: Horas por Dia (Gráfico de Barras Interativo)
            st.write("📅 **Horas Trabalhadas por Dia**")
            st.write("*💡 Passe o mouse sobre as barras para ver detalhes. Este gráfico mostra sua jornada diária, permitindo identificar dias com mais ou menos trabalho.*")
            try:
                dates_str = full_df['Data'].dt.strftime('%d/%m/%Y')
                fig1 = go.Figure()
                
                # Barras das horas trabalhadas
                fig1.add_trace(go.Bar(
                    x=dates_str, 
                    y=full_df['Horas/Dia'],
                    name='Horas Trabalhadas',
                    marker_color='#00ffea',
                    hovertemplate='<b>%{x}</b><br>Horas: %{y:.2f}h<extra></extra>'
                ))
                
                # Linha da média
                fig1.add_trace(go.Scatter(
                    x=dates_str,
                    y=[media_diaria] * len(dates_str),
                    mode='lines',
                    name=f'Média Diária ({media_diaria:.2f}h)',
                    line=dict(color='magenta', dash='dash', width=2),
                    hovertemplate='Média: %{y:.2f}h<extra></extra>'
                ))
                
                # Linha da moda
                fig1.add_trace(go.Scatter(
                    x=dates_str,
                    y=[moda_diaria] * len(dates_str),
                    mode='lines',
                    name=f'Moda ({moda_diaria:.2f}h)',
                    line=dict(color='yellow', dash='dot', width=2),
                    hovertemplate='Moda: %{y:.2f}h<extra></extra>'
                ))
                
                fig1.update_layout(
                    title="Horas Trabalhadas por Dia",
                    xaxis_title="Data",
                    yaxis_title="Horas",
                    template="plotly_dark",
                    height=500,
                    showlegend=True
                )
                st.plotly_chart(fig1, use_container_width=True)
            except Exception as e:
                st.error(f"Erro no gráfico: {str(e)}")

            # Gráfico 2: Distribuição das Horas (Histograma Interativo)
            st.write("📈 **Distribuição das Horas Diárias**")
            st.write("*💡 Este histograma mostra com que frequência você trabalha determinadas quantidades de horas. Picos indicam suas jornadas mais comuns.*")
            try:
                fig2 = go.Figure()
                
                # Histograma
                fig2.add_trace(go.Histogram(
                    x=full_df['Horas/Dia'],
                    nbinsx=15,
                    name='Frequência',
                    marker_color='lime',
                    opacity=0.7,
                    hovertemplate='Horas: %{x:.1f}-%{x:.1f}<br>Frequência: %{y}<extra></extra>'
                ))
                
                # Linha da média
                fig2.add_vline(
                    x=media_diaria, 
                    line_dash="dash", 
                    line_color="cyan",
                    annotation_text=f"Média: {media_diaria:.2f}h",
                    annotation_position="top"
                )
                
                # Linha da moda
                fig2.add_vline(
                    x=moda_diaria, 
                    line_dash="dot", 
                    line_color="yellow",
                    annotation_text=f"Moda: {moda_diaria:.2f}h",
                    annotation_position="bottom"
                )
                
                # Área do desvio padrão
                fig2.add_vrect(
                    x0=media_diaria - desvio_diario,
                    x1=media_diaria + desvio_diario,
                    fillcolor="orange",
                    opacity=0.2,
                    annotation_text="±1 Desvio Padrão",
                    annotation_position="top right"
                )
                
                fig2.update_layout(
                    title="Distribuição das Horas Trabalhadas",
                    xaxis_title="Horas por Dia",
                    yaxis_title="Frequência (Quantos Dias)",
                    template="plotly_dark",
                    height=500,
                    showlegend=False
                )
                st.plotly_chart(fig2, use_container_width=True)
            except Exception as e:
                st.error(f"Erro no gráfico: {str(e)}")

            # Gráfico 3: Totais Semanais (Gráfico de Barras Interativo)
            st.write("📅 **Totais Semanais**")
            st.write("*💡 Visualize suas horas semanais totais. Ajuda a identificar semanas mais intensas e padrões semanais de trabalho.*")
            try:
                week_labels = [str(w) for w in weekly_totals.index]
                fig3 = go.Figure()
                
                # Barras semanais
                fig3.add_trace(go.Bar(
                    x=week_labels,
                    y=weekly_totals.values,
                    name='Horas Semanais',
                    marker_color='#ffaa00',
                    hovertemplate='<b>Semana %{x}</b><br>Total: %{y:.2f}h<extra></extra>'
                ))
                
                # Linha da média semanal
                fig3.add_trace(go.Scatter(
                    x=week_labels,
                    y=[media_semanal] * len(week_labels),
                    mode='lines',
                    name=f'Média Semanal ({media_semanal:.2f}h)',
                    line=dict(color='cyan', dash='dash', width=2),
                    hovertemplate='Média Semanal: %{y:.2f}h<extra></extra>'
                ))
                
                fig3.update_layout(
                    title="Horas Trabalhadas por Semana",
                    xaxis_title="Semana",
                    yaxis_title="Total de Horas",
                    template="plotly_dark",
                    height=500,
                    showlegend=True
                )
                st.plotly_chart(fig3, use_container_width=True)
            except Exception as e:
                st.error(f"Erro no gráfico semanal: {str(e)}")

            st.subheader("📚 Estatísticas Explicadas")

            st.markdown(f"""
            ### 📊 Estatísticas Diárias:
            - **Total Geral:** `{total_geral:.2f}` horas
            - **Média diária:** `{media_diaria:.2f}` horas  
              📈 A média é o valor médio de horas que você trabalhou por dia. É a soma de todas as horas dividida pelo número de dias.
            - **Moda diária:** `{moda_diaria:.2f}` horas  
              🎯 A moda é o valor de horas que você mais trabalhou (mais frequente). Indica sua jornada típica.
            - **Desvio padrão diário:** `{desvio_diario:.2f}` horas  
              📊 O desvio padrão mostra a variabilidade das suas horas diárias. Quanto menor, mais consistente é sua rotina.
              - Se ≤ 1h: Rotina muito consistente 🟢
              - Se 1-2h: Rotina moderadamente variável 🟡  
              - Se > 2h: Rotina muito variável 🔴
              
            ### 📅 Estatísticas Semanais:
            - **Média semanal:** `{media_semanal:.2f}` horas  
              📈 A média de horas trabalhadas por semana completa.
            - **Desvio padrão semanal:** `{desvio_semanal:.2f}` horas  
              📊 Mostra a variação das horas semanais. Quanto maior, mais irregular é sua carga de trabalho semanal.
              
            ### 🔍 Como Interpretar os Gráficos:
            - **Gráfico de Barras Diário:** Cada barra representa um dia. Barras altas = dias intensos.
            - **Histograma:** Mostra quantos dias você trabalhou X horas. Picos = suas jornadas mais comuns.
            - **Gráfico Semanal:** Compare semanas inteiras. Útil para identificar períodos mais intensos.
            """)
            
            st.subheader("📅 Totais por Semana")
            for semana, total_semana in weekly_totals.items():
                st.write(f"**Semana {semana}:** `{total_semana:.2f}` horas")
                
            st.subheader("📅 Totais por Mês")
            for mes, total_mes in monthly_totals_display.items():
                st.write(f"**{mes}:** `{total_mes:.2f}` horas")

if __name__ == "__main__":
    main()