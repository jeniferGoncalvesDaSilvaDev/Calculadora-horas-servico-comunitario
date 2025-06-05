
import pandas as pd
from datetime import datetime
import openpyxl

def criar_excel_dados_fotos():
    """
    Extrai dados das fotos do formulário PSC e cria arquivo Excel
    """
    
    # Dados extraídos das fotos (leitura manual das imagens)
    dados_formulario = [
        # Dados da primeira foto
        {'Data': '05/05/2025', 'Entrada': '08:08', 'Início Intervalo': '09:15', 'Fim Intervalo': '15:02', 'Saída': '18:52'},
        {'Data': '06/05/2025', 'Entrada': '07:06', 'Início Intervalo': '09:06', 'Fim Intervalo': '14:36', 'Saída': '19:00'},
        {'Data': '07/05/2025', 'Entrada': '08:20', 'Início Intervalo': '09:10', 'Fim Intervalo': '15:02', 'Saída': '18:30'},
        {'Data': '08/05/2025', 'Entrada': '08:10', 'Início Intervalo': '09:20', 'Fim Intervalo': '15:10', 'Saída': '18:50'},
        {'Data': '09/05/2025', 'Entrada': '08:00', 'Início Intervalo': '09:30', 'Fim Intervalo': '14:10', 'Saída': '18:30'},
        {'Data': '12/05/2025', 'Entrada': '08:20', 'Início Intervalo': '09:50', 'Fim Intervalo': '15:05', 'Saída': '18:30'},
        {'Data': '13/05/2025', 'Entrada': '08:15', 'Início Intervalo': '09:15', 'Fim Intervalo': '15:02', 'Saída': '18:00'},
        {'Data': '14/05/2025', 'Entrada': 'Feriado', 'Início Intervalo': '', 'Fim Intervalo': '', 'Saída': ''},
        {'Data': '15/05/2025', 'Entrada': '08:04', 'Início Intervalo': '09:30', 'Fim Intervalo': '14:30', 'Saída': '18:00'},
        {'Data': '16/05/2025', 'Entrada': '08:10', 'Início Intervalo': '09:50', 'Fim Intervalo': '15:05', 'Saída': '18:00'},
        {'Data': '19/05/2025', 'Entrada': '08:05', 'Início Intervalo': '09:20', 'Fim Intervalo': '14:20', 'Saída': '18:30'},
        {'Data': '20/05/2025', 'Entrada': '08:08', 'Início Intervalo': '09:25', 'Fim Intervalo': '15:08', 'Saída': '18:30'},
        {'Data': '21/05/2025', 'Entrada': '08:01', 'Início Intervalo': '09:10', 'Fim Intervalo': '15:05', 'Saída': '18:00'},
        {'Data': '22/05/2025', 'Entrada': '08:28', 'Início Intervalo': '09:17', 'Fim Intervalo': '15:05', 'Saída': '18:50'},
        {'Data': '23/05/2025', 'Entrada': '08:09', 'Início Intervalo': '09:30', 'Fim Intervalo': '14:28', 'Saída': '18:50'},
        {'Data': '26/05/2025', 'Entrada': '08:09', 'Início Intervalo': '09:10', 'Fim Intervalo': '15:05', 'Saída': '18:30'},
        {'Data': '27/05/2025', 'Entrada': '08:14', 'Início Intervalo': '09:20', 'Fim Intervalo': '14:55', 'Saída': '18:30'},
        {'Data': '28/05/2025', 'Entrada': '08:30', 'Início Intervalo': '10:07', 'Fim Intervalo': '15:16', 'Saída': '18:50'},
        
        # Dados da segunda foto (complementares)
        {'Data': '30/05/2025', 'Entrada': '08:10', 'Início Intervalo': '09:10', 'Fim Intervalo': '15:00', 'Saída': '18:00'},
        {'Data': '02/06/2025', 'Entrada': '08:16', 'Início Intervalo': '09:06', 'Fim Intervalo': '14:40', 'Saída': '18:00'},
    ]
    
    # Criar DataFrame
    df = pd.DataFrame(dados_formulario)
    
    # Função para calcular horas trabalhadas
    def calcular_horas_psc(row):
        try:
            if row['Entrada'] == 'Feriado' or row['Entrada'] == '':
                return 0
                
            entrada = pd.to_datetime(row['Entrada'], format='%H:%M', errors='coerce')
            saida = pd.to_datetime(row['Saída'], format='%H:%M', errors='coerce')
            ini_intervalo = pd.to_datetime(row['Início Intervalo'], format='%H:%M', errors='coerce')
            fim_intervalo = pd.to_datetime(row['Fim Intervalo'], format='%H:%M', errors='coerce')

            if pd.isna(entrada) or pd.isna(saida):
                return 0

            # Calcular períodos de trabalho
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
    
    # Calcular horas trabalhadas
    df['Horas/Dia'] = df.apply(calcular_horas_psc, axis=1)
    
    # Converter datas
    df['Data'] = pd.to_datetime(df['Data'], format='%d/%m/%Y', errors='coerce')
    df_limpo = df.dropna(subset=['Data']).sort_values('Data')
    df_limpo['Data'] = df_limpo['Data'].dt.strftime('%d/%m/%Y')
    
    # Calcular estatísticas
    total_horas = df_limpo['Horas/Dia'].sum()
    media_diaria = df_limpo['Horas/Dia'].mean()
    
    # Calcular totais semanais
    df_limpo['Data_dt'] = pd.to_datetime(df_limpo['Data'], format='%d/%m/%Y')
    df_limpo['Semana'] = df_limpo['Data_dt'].dt.to_period('W')
    weekly_totals = df_limpo.groupby('Semana')['Horas/Dia'].sum().reset_index()
    weekly_totals['Semana'] = weekly_totals['Semana'].astype(str)
    
    # Calcular totais mensais
    df_limpo['Mes_Ano'] = df_limpo['Data_dt'].dt.to_period('M')
    monthly_totals = df_limpo.groupby('Mes_Ano')['Horas/Dia'].sum().reset_index()
    monthly_totals['Mes_Ano'] = monthly_totals['Mes_Ano'].astype(str)
    
    # Preparar dados para Excel (sem coluna de data auxiliar)
    df_excel = df_limpo[['Data', 'Entrada', 'Início Intervalo', 'Fim Intervalo', 'Saída', 'Horas/Dia']].copy()
    
    # Criar arquivo Excel
    nome_arquivo = 'dados_formulario_psc_extraidos.xlsx'
    
    with pd.ExcelWriter(nome_arquivo, engine='xlsxwriter') as writer:
        # Aba principal com dados
        df_excel.to_excel(writer, sheet_name='Dados PSC', index=False)
        
        # Aba de resumo estatístico
        resumo_dados = [
            ['Total de Horas', f'{total_horas:.2f}'],
            ['Média Diária', f'{media_diaria:.2f}'],
            ['Dias Trabalhados', len(df_limpo[df_limpo['Horas/Dia'] > 0])],
            ['Período', f"{df_limpo['Data'].min()} a {df_limpo['Data'].max()}"]
        ]
        
        resumo_df = pd.DataFrame(resumo_dados, columns=['Estatística', 'Valor'])
        resumo_df.to_excel(writer, sheet_name='Resumo', index=False)
        
        # Aba de totais semanais
        weekly_totals.columns = ['Semana', 'Total Horas']
        weekly_totals.to_excel(writer, sheet_name='Totais Semanais', index=False)
        
        # Aba de totais mensais
        monthly_totals.columns = ['Mês/Ano', 'Total Horas']
        monthly_totals.to_excel(writer, sheet_name='Totais Mensais', index=False)
    
    print(f"✅ Arquivo '{nome_arquivo}' criado com sucesso!")
    print(f"📊 Dados extraídos das fotos do formulário PSC")
    print(f"📅 Período: {df_limpo['Data'].min()} a {df_limpo['Data'].max()}")
    print(f"⏰ Total de horas: {total_horas:.2f}")
    print(f"📈 Média diária: {media_diaria:.2f}")
    print(f"🗓️ Dias trabalhados: {len(df_limpo[df_limpo['Horas/Dia'] > 0])}")
    
    # Mostrar amostra dos dados
    print(f"\n📋 Primeiras linhas dos dados extraídos:")
    print(df_excel.head(10).to_string(index=False))
    
    return df_excel

if __name__ == "__main__":
    dados_extraidos = criar_excel_dados_fotos()
