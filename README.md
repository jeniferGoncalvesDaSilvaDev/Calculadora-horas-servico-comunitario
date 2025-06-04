
# 🕶️ Controle de Horas - Análise OCR e Visualização de Jornadas

Um sistema moderno e interativo para análise de horas trabalhadas com interface cyberpunk, desenvolvido em Python com Streamlit.

## 📋 Funcionalidades

### 🔍 Análise de Dados
- **OCR de Imagens**: Extrai dados de horários diretamente de imagens/fotos de formulários
- **Importação Excel**: Suporte para múltiplos formatos de planilhas Excel
- **Upload Múltiplo**: Processa vários arquivos de uma vez

### 📊 Visualizações Interativas
- **Gráficos de Barras**: Horas trabalhadas por dia com médias e modas
- **Histogramas**: Distribuição das jornadas diárias
- **Análise Semanal**: Totais e médias por semana
- **Gráficos Responsivos**: Interface interativa com Plotly

### 📈 Estatísticas Avançadas
- **Métricas Diárias**: Média, moda e desvio padrão
- **Métricas Semanais**: Totais e variabilidade semanal
- **Métricas Mensais**: Consolidação por mês
- **Análise de Consistência**: Avaliação da regularidade da rotina

### 💾 Exportação
- **Excel Automático**: Download instantâneo após processamento
- **Múltiplas Abas**: Detalhes, resumo, totais semanais e mensais
- **Formatação Profissional**: Dados organizados e prontos para uso

## 🛠️ Tecnologias Utilizadas

- **Python 3.11**
- **Streamlit** - Interface web
- **Pandas** - Manipulação de dados
- **Plotly** - Visualizações interativas
- **PIL (Pillow)** - Processamento de imagens
- **OpenPyXL** - Manipulação de Excel
- **SciPy** - Cálculos estatísticos
- **OCR.space API** - Reconhecimento ótico de caracteres

## 📁 Estrutura de Arquivos Suportados

### 🖼️ Imagens (OCR)
- Formatos: JPG, JPEG, PNG
- Deve conter: Data, horários de entrada/saída e intervalos

### 📊 Excel
Suporta dois formatos principais:

**Formato 1 - Tradicional:**
- Data
- Horário de Entrada
- Início do Intervalo
- Fim do Intervalo  
- Horário de Saída

**Formato 2 - Separado:**
- Horário de entrada 1
- Horário de saída 1
- Horário de entrada 2
- Horário de saída 2

## 🚀 Como Usar

1. **Acesse a aplicação** clicando no botão "Run"
2. **Escolha o tipo de arquivo**: Imagens (OCR) ou Excel
3. **Faça upload** de um ou múltiplos arquivos
4. **Visualize os resultados** automaticamente
5. **Baixe o Excel** com análise completa

## 📊 Cálculo de Horas

O sistema usa a seguinte fórmula:

```
Horas Trabalhadas = (Saída1 - Entrada1) + (Saída2 - Entrada2) - (Fim Intervalo - Início Intervalo)
```

Onde:
- **Período 1**: Manhã (Entrada até Início do Intervalo)
- **Período 2**: Tarde (Fim do Intervalo até Saída)
- **Intervalo**: Deduzido automaticamente

## 📈 Interpretação das Estatísticas

### 📊 Métricas Diárias
- **Média**: Horas médias trabalhadas por dia
- **Moda**: Jornada mais frequente
- **Desvio Padrão**: 
  - ≤ 1h: Rotina consistente 🟢
  - 1-2h: Moderadamente variável 🟡
  - > 2h: Muito variável 🔴

### 📅 Análise Semanal
- **Total Semanal**: Soma das horas da semana
- **Média Semanal**: Média de horas por semana
- **Variabilidade**: Consistência da carga semanal

## 🎨 Interface

- **Tema Cyberpunk**: Design moderno com cores neon
- **Responsiva**: Adapta-se a diferentes tamanhos de tela
- **Interativa**: Gráficos com hover e zoom
- **Intuitiva**: Interface simples e fácil de usar

## 📦 Arquivos de Exemplo

O projeto inclui arquivos de teste:
- `teste_horas.xlsx` - Formato tradicional
- `exemplo_horarios_separados.xlsx` - Formato com períodos separados
- `horarios_4_colunas.xlsx` - Formato simplificado

## 🔧 Configuração

### OCR.space API (Opcional)
Para usar OCR real, adicione sua chave da API:
```python
OCR_SPACE_API_KEY = 'sua_chave_aqui'
```

Sem a chave, o sistema usa dados de demonstração.

## 🚀 Deploy

O projeto está configurado para deploy automático no Replit:
- Porta: 8501 (mapeada para 80 em produção)
- Comando: `streamlit run main.py`
- Target: Google Cloud Run

## 📄 Licença

Este projeto é de uso livre para fins educacionais e profissionais.

---

**Desenvolvido com ❤️ em Python | Interface Cyberpunk 🕶️**
