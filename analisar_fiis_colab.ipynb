# ============================================================
# NOME DO SCRIPT: analisar_fiis_colab.py
# OBJETIVO: Coletar dados de FIIs do Fundamentus, aplicar filtros
#           de qualidade previdenciária e exportar planilha Excel
#           formatada com ranking multicritério
# BIBLIOTECAS NECESSÁRIAS: requests, pandas, openpyxl, lxml
# COMO EXECUTAR: Google Colab — rodar célula por célula
# ============================================================

# ===========================================================
# CÉLULA 1 — Instalação das bibliotecas
# ===========================================================
# No Colab, algumas bibliotecas não vêm pré-instaladas.
# Esta célula garante que tudo está disponível antes de começar.
# Execute esta célula primeiro e aguarde a conclusão.

# !pip install requests pandas openpyxl lxml --quiet


# ===========================================================
# CÉLULA 2 — Importações
# ===========================================================

# --- Bibliotecas padrão do Python (já vêm instaladas) ---
import time          # Para pausas entre requisições (boa prática de scraping)
import re            # Para expressões regulares (limpeza de texto)
from datetime import datetime  # Para registrar data/hora da coleta
import io            # Para manipular dados em memória (download no Colab)

# --- Bibliotecas de dados ---
import requests      # Para fazer requisições HTTP ao Fundamentus
import pandas as pd  # Para manipular tabelas de dados (DataFrames)

# --- Biblioteca para gerar o arquivo Excel ---
from openpyxl import Workbook                          # Criar planilha
from openpyxl.styles import (
    Font,            # Controlar fonte (negrito, cor, tamanho)
    PatternFill,     # Cor de fundo das células
    Alignment,       # Alinhamento do texto
    Border,          # Bordas das células
    Side             # Estilo das linhas de borda
)
from openpyxl.utils import get_column_letter  # Converte número de coluna → letra (1 → 'A')

# --- Para download automático no Colab ---
try:
    # Esta função só existe no ambiente do Google Colab
    from google.colab import files
    RODANDO_NO_COLAB = True
except ImportError:
    # Se não estiver no Colab, salva localmente sem erro
    RODANDO_NO_COLAB = False

print("✅ Bibliotecas importadas com sucesso!")
print(f"📅 Data/hora da análise: {datetime.now().strftime('%d/%m/%Y às %H:%M')}")


# ===========================================================
# CÉLULA 3 — Configurações (ajuste aqui conforme necessário)
# ===========================================================

# --- Filtros de qualidade mínima ---
# FIIs que não atingirem esses critérios serão excluídos da planilha.
# Você pode afrouxar ou apertar esses valores conforme sua estratégia.

FILTRO_LIQUIDEZ_MINIMA    = 500_000   # R$ 500 mil de volume médio diário
                                       # Garante que você consegue comprar/vender sem impactar o preço

FILTRO_PVP_MINIMO         = 0.50      # P/VP mínimo
FILTRO_PVP_MAXIMO         = 1.30      # P/VP máximo
                                       # Faixa razoável: nem muito barato (sinal de problema)
                                       # nem muito caro (prêmio excessivo sobre o patrimônio)

FILTRO_DY_MINIMO          = 0.06      # DY mínimo de 6% ao ano
                                       # Abaixo disso, o rendimento não compensa o risco

FILTRO_VACANCIA_MAXIMA    = 0.25      # Vacância máxima de 25%
                                       # Acima disso, muitos imóveis vazios = risco de corte de rendimento

# --- Pesos do ranking multicritério ---
# A soma dos pesos deve ser 1.0 (100%).
# Você pode redistribuir conforme sua preferência.
PESO_DY                   = 0.30      # 30% — Rendimento é o foco previdenciário
PESO_PVP                  = 0.25      # 25% — Valuation: quanto estou pagando pelo patrimônio
PESO_VACANCIA             = 0.25      # 25% — Qualidade dos imóveis (vacância baixa = imóveis ocupados)
PESO_LIQUIDEZ             = 0.20      # 20% — Facilidade de operar sem impactar o preço

# --- Nome do arquivo de saída ---
NOME_ARQUIVO = f"analise_fiis_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

print("✅ Configurações definidas!")
print(f"\n📋 Filtros ativos:")
print(f"   Liquidez mínima:  R$ {FILTRO_LIQUIDEZ_MINIMA:,.0f}")
print(f"   P/VP:             entre {FILTRO_PVP_MINIMO} e {FILTRO_PVP_MAXIMO}")
print(f"   DY mínimo:        {FILTRO_DY_MINIMO*100:.0f}% ao ano")
print(f"   Vacância máxima:  {FILTRO_VACANCIA_MAXIMA*100:.0f}%")


# ===========================================================
# CÉLULA 4 — Funções auxiliares
# ===========================================================

def converter_numero_br(valor):
    """
    Converte strings numéricas no formato brasileiro para float.

    O Fundamentus usa formatação brasileira com vírgulas e pontos,
    além de uma codificação implícita em percentuais.

    Exemplos de conversão:
        "1.234,56"  → 1234.56   (número com separadores BR)
        "12,34%"    → 0.1234    (percentual → decimal)
        "011"       → 0.11      (percentual implícito de 2 casas)
        ""          → None      (vazio → nulo)
        "N/D"       → None      (não disponível → nulo)

    Parâmetros:
        valor: qualquer tipo — será convertido para string antes do processamento

    Retorna:
        float ou None
    """

    # Se o valor for nulo (None ou NaN do pandas), retorna None diretamente
    if pd.isna(valor):
        return None

    # Converte para string e remove espaços nas bordas
    texto = str(valor).strip()

    # Se estiver vazio ou for um marcador de "não disponível", retorna None
    if texto in ('', '-', 'N/D', 'N/A', '#N/D'):
        return None

    # Verifica se é um percentual explícito com o símbolo %
    tem_percentual = '%' in texto

    # Remove o símbolo de % e quaisquer espaços
    texto = texto.replace('%', '').strip()

    # Converte separadores do formato brasileiro para o formato Python:
    # No Brasil:  1.234,56  (ponto = milhar, vírgula = decimal)
    # No Python:  1234.56   (sem ponto de milhar, ponto = decimal)
    texto = texto.replace('.', '').replace(',', '.')

    try:
        numero = float(texto)
    except ValueError:
        # Se ainda assim não conseguir converter, retorna None
        return None

    # Se era um percentual explícito ("12,34%"), divide por 100
    if tem_percentual:
        return numero / 100

    return numero


def coletar_fiis_fundamentus():
    """
    Coleta a tabela completa de FIIs do site Fundamentus.

    O Fundamentus tem uma página que lista todos os FIIs com seus
    principais indicadores em uma única tabela HTML. Fazemos scraping
    dessa tabela usando pandas + requests.

    Retorna:
        DataFrame com todos os FIIs e indicadores brutos,
        ou None em caso de falha na coleta.
    """

    # URL da página de FIIs do Fundamentus
    url = "https://www.fundamentus.com.br/fii_resultado.php"

    # Headers simulando um navegador real.
    # Sem isso, o Fundamentus pode bloquear a requisição (erro 403).
    headers = {
        'User-Agent': (
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
            'AppleWebKit/537.36 (KHTML, like Gecko) '
            'Chrome/120.0.0.0 Safari/537.36'
        ),
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7',
        'Referer': 'https://www.fundamentus.com.br/'
    }

    print("🌐 Conectando ao Fundamentus...")

    try:
        # Faz a requisição HTTP com timeout de 30 segundos
        resposta = requests.get(url, headers=headers, timeout=30)

        # Verifica se a requisição foi bem-sucedida (código 200)
        # raise_for_status() lança exceção automaticamente para códigos 4xx e 5xx
        resposta.raise_for_status()

        print("✅ Conexão bem-sucedida! Extraindo tabela...")

        # O pandas consegue ler tabelas HTML diretamente do conteúdo da página.
        # read_html() retorna uma lista de DataFrames (um para cada tabela encontrada).
        # O Fundamentus tem apenas uma tabela relevante na página.
        tabelas = pd.read_html(
            io.StringIO(resposta.text),  # Conteúdo HTML como string
            decimal=',',                  # Decimal brasileiro usa vírgula
            thousands='.'                 # Milhar brasileiro usa ponto
        )

        # Pega a primeira (e única) tabela da página
        df = tabelas[0]

        print(f"📊 {len(df)} FIIs encontrados no Fundamentus.")
        return df

    except requests.exceptions.ConnectionError:
        print("❌ Erro: Sem conexão com a internet. Verifique sua conexão.")
        return None

    except requests.exceptions.Timeout:
        print("❌ Erro: O Fundamentus demorou demais para responder. Tente novamente.")
        return None

    except requests.exceptions.HTTPError as e:
        print(f"❌ Erro HTTP ao acessar o Fundamentus: {e}")
        return None

    except Exception as e:
        print(f"❌ Erro inesperado na coleta: {e}")
        return None


def limpar_e_padronizar(df_bruto):
    """
    Limpa e padroniza o DataFrame bruto do Fundamentus.

    O Fundamentus retorna os dados em formato de exibição (para humanos),
    não em formato numérico (para cálculos). Esta função:
    1. Renomeia as colunas para nomes em português padronizados
    2. Converte strings para float usando converter_numero_br()
    3. Remove linhas com dados críticos faltando

    Parâmetros:
        df_bruto: DataFrame retornado diretamente pelo read_html()

    Retorna:
        DataFrame limpo e padronizado
    """

    print("\n🔧 Limpando e padronizando os dados...")

    # --- Passo 1: Verificar e renomear colunas ---
    # O Fundamentus pode mudar os nomes das colunas sem aviso.
    # Mapeamos os nomes originais para nomes padronizados em português.
    print(f"   Colunas encontradas: {list(df_bruto.columns)}")

    # Mapeamento: nome original → nome padronizado
    # Ajuste aqui se o Fundamentus mudar os nomes
    mapa_colunas = {
        'Papel':          'ticker',
        'Segmento':       'segmento',
        'Cotação':        'cotacao',
        'FFO Yield':      'ffo_yield',
        'Dividend Yield': 'dy_anual',
        'P/VP':           'pvp',
        'Valor de Mercado': 'valor_mercado',
        'Liquidez':       'liquidez_diaria',
        'Qtd de imóveis': 'qtd_imoveis',
        'Vacância Média': 'vacancia'
    }

    # Renomeia apenas as colunas que existem no DataFrame
    # (evita erro se alguma coluna não existir)
    colunas_presentes = {k: v for k, v in mapa_colunas.items() if k in df_bruto.columns}
    df = df_bruto.rename(columns=colunas_presentes)

    # --- Passo 2: Converter colunas numéricas ---
    # Cada coluna numérica precisa passar pela função de conversão BR
    colunas_numericas = ['cotacao', 'ffo_yield', 'dy_anual', 'pvp',
                         'valor_mercado', 'liquidez_diaria', 'qtd_imoveis', 'vacancia']

    for coluna in colunas_numericas:
        if coluna in df.columns:
            # Aplica a função de conversão em cada célula da coluna
            df[coluna] = df[coluna].apply(converter_numero_br)

    # --- Passo 3: Remover linhas sem dados críticos ---
    # Sem ticker, DY, P/VP e liquidez, não podemos analisar o FII
    colunas_obrigatorias = ['ticker', 'dy_anual', 'pvp', 'liquidez_diaria']
    colunas_obrigatorias_presentes = [c for c in colunas_obrigatorias if c in df.columns]

    linhas_antes = len(df)
    df = df.dropna(subset=colunas_obrigatorias_presentes)
    linhas_removidas = linhas_antes - len(df)

    print(f"   {linhas_removidas} FIIs removidos por dados faltantes.")
    print(f"   {len(df)} FIIs com dados completos.")

    return df


def aplicar_filtros_qualidade(df):
    """
    Aplica os filtros de qualidade mínima definidos nas configurações.

    Cada filtro remove FIIs que não atendem aos critérios previdenciários.
    O resultado é uma lista menor, mas de maior qualidade.

    Parâmetros:
        df: DataFrame limpo e padronizado

    Retorna:
        DataFrame filtrado com apenas FIIs que passaram em todos os critérios
    """

    print("\n🔍 Aplicando filtros de qualidade...")
    total_inicial = len(df)

    # --- Filtro 1: Liquidez mínima ---
    # FIIs com pouco volume são difíceis de comprar/vender
    if 'liquidez_diaria' in df.columns:
        antes = len(df)
        df = df[df['liquidez_diaria'] >= FILTRO_LIQUIDEZ_MINIMA]
        removidos = antes - len(df)
        print(f"   Liquidez < R$ {FILTRO_LIQUIDEZ_MINIMA:,.0f}: {removidos} FIIs removidos")

    # --- Filtro 2: P/VP dentro da faixa aceitável ---
    # Muito baixo = possível problema estrutural
    # Muito alto = pagando prêmio excessivo sobre o patrimônio
    if 'pvp' in df.columns:
        antes = len(df)
        df = df[(df['pvp'] >= FILTRO_PVP_MINIMO) & (df['pvp'] <= FILTRO_PVP_MAXIMO)]
        removidos = antes - len(df)
        print(f"   P/VP fora de [{FILTRO_PVP_MINIMO}, {FILTRO_PVP_MAXIMO}]: {removidos} FIIs removidos")

    # --- Filtro 3: DY mínimo ---
    # Rendimento abaixo do mínimo não justifica o risco
    if 'dy_anual' in df.columns:
        antes = len(df)
        df = df[df['dy_anual'] >= FILTRO_DY_MINIMO]
        removidos = antes - len(df)
        print(f"   DY < {FILTRO_DY_MINIMO*100:.0f}%: {removidos} FIIs removidos")

    # --- Filtro 4: Vacância máxima ---
    # Vacância alta = muitos imóveis vazios = risco de corte de rendimento
    # Nota: alguns FIIs (Papel, FOF) não têm vacância → mantemos esses
    if 'vacancia' in df.columns:
        antes = len(df)
        # A condição mantém FIIs sem vacância (NaN) E FIIs dentro do limite
        df = df[(df['vacancia'].isna()) | (df['vacancia'] <= FILTRO_VACANCIA_MAXIMA)]
        removidos = antes - len(df)
        print(f"   Vacância > {FILTRO_VACANCIA_MAXIMA*100:.0f}%: {removidos} FIIs removidos")

    total_final = len(df)
    print(f"\n✅ {total_final} FIIs aprovados nos filtros (de {total_inicial} analisados)")

    return df.reset_index(drop=True)


def calcular_score(df):
    """
    Calcula o score multicritério de cada FII aprovado nos filtros.

    O score funciona normalizando cada indicador para uma escala de 0 a 10
    e depois combinando com os pesos definidos nas configurações.

    Normalização min-max:
        score_indicador = (valor - min) / (max - min) × 10
        - O pior valor do grupo recebe 0
        - O melhor valor do grupo recebe 10
        - Os demais ficam proporcionalmente entre 0 e 10

    Para indicadores onde "menor é melhor" (P/VP, vacância),
    invertemos a escala para que o menor valor receba 10.

    Parâmetros:
        df: DataFrame filtrado

    Retorna:
        DataFrame com colunas de score por indicador e score_final
    """

    print("\n📐 Calculando scores do ranking...")

    df = df.copy()  # Evita modificar o DataFrame original

    def normalizar(serie, inverter=False):
        """
        Normaliza uma série para escala 0–10.
        inverter=True: menor valor recebe pontuação maior (ex: vacância)
        """
        minimo = serie.min()
        maximo = serie.max()

        # Evita divisão por zero se todos os valores forem iguais
        if maximo == minimo:
            return pd.Series([5.0] * len(serie), index=serie.index)

        # Normalização min-max
        normalizado = (serie - minimo) / (maximo - minimo) * 10

        # Inverte se necessário (para indicadores onde menor = melhor)
        if inverter:
            normalizado = 10 - normalizado

        return normalizado

    # --- Score de DY: maior DY → maior score ---
    if 'dy_anual' in df.columns:
        df['score_dy'] = normalizar(df['dy_anual'], inverter=False)
    else:
        df['score_dy'] = 5.0  # Score neutro se coluna ausente

    # --- Score de P/VP: mais próximo de 1,0 → maior score ---
    # Distância ao valor 1,0 — quanto menor, melhor
    if 'pvp' in df.columns:
        distancia_pvp = (df['pvp'] - 1.0).abs()
        df['score_pvp'] = normalizar(distancia_pvp, inverter=True)
    else:
        df['score_pvp'] = 5.0

    # --- Score de Vacância: menor vacância → maior score ---
    # FIIs sem vacância (Papel/FOF) recebem score máximo (10)
    if 'vacancia' in df.columns:
        vacancia_preenchida = df['vacancia'].fillna(0)  # NaN = 0% de vacância
        df['score_vacancia'] = normalizar(vacancia_preenchida, inverter=True)
    else:
        df['score_vacancia'] = 5.0

    # --- Score de Liquidez: maior liquidez → maior score ---
    if 'liquidez_diaria' in df.columns:
        df['score_liquidez'] = normalizar(df['liquidez_diaria'], inverter=False)
    else:
        df['score_liquidez'] = 5.0

    # --- Score Final: média ponderada dos scores individuais ---
    df['score_final'] = (
        df['score_dy']       * PESO_DY       +
        df['score_pvp']      * PESO_PVP      +
        df['score_vacancia'] * PESO_VACANCIA +
        df['score_liquidez'] * PESO_LIQUIDEZ
    ).round(2)

    # Ordena do maior score para o menor
    df = df.sort_values('score_final', ascending=False).reset_index(drop=True)

    # Cria a coluna de posição no ranking (1º, 2º, 3º...)
    df.insert(0, 'ranking', range(1, len(df) + 1))

    print(f"✅ Score calculado para {len(df)} FIIs.")
    return df


# ===========================================================
# CÉLULA 5 — Funções de formatação da planilha Excel
# ===========================================================

def criar_estilo_cabecalho():
    """Retorna as configurações de estilo para o cabeçalho da tabela."""

    return {
        'font':      Font(name='Arial', bold=True, color='FFFFFF', size=11),
        'fill':      PatternFill('solid', start_color='1F4E79'),  # Azul escuro
        'alignment': Alignment(horizontal='center', vertical='center', wrap_text=True),
        'border':    Border(
            left=Side(style='thin', color='FFFFFF'),
            right=Side(style='thin', color='FFFFFF'),
            bottom=Side(style='medium', color='FFFFFF')
        )
    }


def criar_estilo_linha(posicao):
    """
    Retorna o estilo para linhas de dados, alternando cores.
    Linhas pares recebem fundo levemente colorido (zebra striping).
    """

    # Cor de fundo alternada para facilitar a leitura
    cor_fundo = 'DCE6F1' if posicao % 2 == 0 else 'FFFFFF'  # Azul claro ou branco

    return {
        'fill':      PatternFill('solid', start_color=cor_fundo),
        'alignment': Alignment(horizontal='center', vertical='center'),
        'font':      Font(name='Arial', size=10),
        'border':    Border(
            left=Side(style='thin', color='BDD7EE'),
            right=Side(style='thin', color='BDD7EE'),
            bottom=Side(style='thin', color='BDD7EE')
        )
    }


def cor_score(score):
    """
    Retorna a cor de fundo para células de score baseada no valor.
    Verde para scores altos, amarelo para médios, laranja para baixos.
    """
    if score >= 7.0:
        return '70AD47'   # Verde — excelente
    elif score >= 5.0:
        return 'FFC000'   # Amarelo — bom
    else:
        return 'FF6B35'   # Laranja — aceitável (passou nos filtros mas score menor)


def gerar_planilha_excel(df, nome_arquivo):
    """
    Gera o arquivo Excel formatado com os resultados da análise.

    A planilha tem duas abas:
    1. "Ranking Geral" — todos os FIIs aprovados, ordenados por score
    2. "Metadados" — informações sobre a análise (data, filtros usados)

    Parâmetros:
        df:           DataFrame com os FIIs filtrados e rankeados
        nome_arquivo: nome do arquivo .xlsx a ser gerado

    Retorna:
        None (salva o arquivo em disco)
    """

    print(f"\n📝 Gerando planilha Excel: {nome_arquivo}")

    # Cria um novo workbook (arquivo Excel em branco)
    wb = Workbook()

    # -------------------------------------------------------
    # ABA 1: RANKING GERAL
    # -------------------------------------------------------
    aba_ranking = wb.active
    aba_ranking.title = "Ranking Geral"

    # --- Definição das colunas da planilha ---
    # Cada tupla: (nome da coluna no df, título na planilha, largura, formato numérico)
    colunas = [
        ('ranking',         '#',              6,   None),
        ('ticker',          'Ticker',         12,  None),
        ('segmento',        'Segmento',       22,  None),
        ('cotacao',         'Cotação (R$)',   14,  'R$ #,##0.00'),
        ('dy_anual',        'DY Anual',       12,  '0.00%'),
        ('pvp',             'P/VP',           10,  '0.00'),
        ('vacancia',        'Vacância',       12,  '0.00%'),
        ('liquidez_diaria', 'Liquidez/dia',   16,  'R$ #,##0'),
        ('qtd_imoveis',     'Qtd Imóveis',   13,  '0'),
        ('score_dy',        'Score DY',       11,  '0.0'),
        ('score_pvp',       'Score P/VP',     12,  '0.0'),
        ('score_vacancia',  'Score Vacância', 15,  '0.0'),
        ('score_liquidez',  'Score Liq.',     12,  '0.0'),
        ('score_final',     'Score Final',    12,  '0.00'),
    ]

    # --- Linha de título da planilha ---
    aba_ranking.merge_cells('A1:N1')  # Mescla da coluna A até N
    celula_titulo = aba_ranking['A1']
    celula_titulo.value = f"📊 Ranking de FIIs — Análise Previdenciária — {datetime.now().strftime('%d/%m/%Y')}"
    celula_titulo.font = Font(name='Arial', bold=True, size=14, color='1F4E79')
    celula_titulo.alignment = Alignment(horizontal='center', vertical='center')
    aba_ranking.row_dimensions[1].height = 30

    # --- Linha de cabeçalho (linha 2) ---
    estilo_cab = criar_estilo_cabecalho()
    aba_ranking.row_dimensions[2].height = 35

    for col_idx, (_, titulo, largura, _) in enumerate(colunas, start=1):
        letra = get_column_letter(col_idx)
        celula = aba_ranking[f'{letra}2']
        celula.value = titulo

        # Aplica todos os estilos do cabeçalho
        celula.font      = estilo_cab['font']
        celula.fill      = estilo_cab['fill']
        celula.alignment = estilo_cab['alignment']
        celula.border    = estilo_cab['border']

        # Define a largura da coluna
        aba_ranking.column_dimensions[letra].width = largura

    # --- Linhas de dados (a partir da linha 3) ---
    for idx_linha, (_, linha) in enumerate(df.iterrows()):

        # A linha do Excel começa em 3 (1 = título, 2 = cabeçalho)
        linha_excel = idx_linha + 3
        estilo = criar_estilo_linha(idx_linha)

        for col_idx, (nome_col, _, _, formato) in enumerate(colunas, start=1):
            letra = get_column_letter(col_idx)
            celula = aba_ranking[f'{letra}{linha_excel}']

            # Pega o valor do DataFrame (ou vazio se a coluna não existir)
            valor = linha.get(nome_col, '')

            # Trata valores nulos para não mostrar "None" na planilha
            if pd.isna(valor) if not isinstance(valor, str) else False:
                celula.value = '-'
            else:
                celula.value = valor

            # Aplica formato numérico se definido
            if formato and not pd.isna(valor) if not isinstance(valor, str) else False:
                celula.number_format = formato

            # Aplica estilo de linha (fundo alternado)
            celula.fill      = estilo['fill']
            celula.alignment = estilo['alignment']
            celula.font      = estilo['font']
            celula.border    = estilo['border']

            # Estilo especial para a coluna de Score Final
            if nome_col == 'score_final' and isinstance(valor, (int, float)):
                celula.fill = PatternFill('solid', start_color=cor_score(valor))
                celula.font = Font(name='Arial', bold=True, size=10, color='FFFFFF')
                celula.number_format = '0.00'

            # Ticker em negrito para facilitar identificação
            if nome_col == 'ticker':
                celula.font = Font(name='Arial', bold=True, size=10, color='1F4E79')
                celula.alignment = Alignment(horizontal='left', vertical='center')

    # --- Congela as linhas de título e cabeçalho ao rolar ---
    # A partir da linha 3, os dados rolam mas o cabeçalho fica fixo
    aba_ranking.freeze_panes = 'A3'

    # -------------------------------------------------------
    # ABA 2: METADADOS (informações sobre a análise)
    # -------------------------------------------------------
    aba_meta = wb.create_sheet("Metadados")

    metadados = [
        ("INFORMAÇÕES DA ANÁLISE", ""),
        ("", ""),
        ("Data e hora da coleta",    datetime.now().strftime('%d/%m/%Y às %H:%M')),
        ("Fonte dos dados",          "Fundamentus.com.br"),
        ("Total de FIIs analisados", "Ver Ranking Geral"),
        ("", ""),
        ("FILTROS APLICADOS", ""),
        ("", ""),
        ("Liquidez mínima diária",   f"R$ {FILTRO_LIQUIDEZ_MINIMA:,.0f}"),
        ("P/VP mínimo",              f"{FILTRO_PVP_MINIMO}"),
        ("P/VP máximo",              f"{FILTRO_PVP_MAXIMO}"),
        ("DY anual mínimo",          f"{FILTRO_DY_MINIMO*100:.0f}%"),
        ("Vacância máxima",          f"{FILTRO_VACANCIA_MAXIMA*100:.0f}%"),
        ("", ""),
        ("PESOS DO RANKING", ""),
        ("", ""),
        ("Peso DY",                  f"{PESO_DY*100:.0f}%"),
        ("Peso P/VP",                f"{PESO_PVP*100:.0f}%"),
        ("Peso Vacância",            f"{PESO_VACANCIA*100:.0f}%"),
        ("Peso Liquidez",            f"{PESO_LIQUIDEZ*100:.0f}%"),
        ("", ""),
        ("AVISO LEGAL", ""),
        ("", ""),
        ("Disclaimer",
         "Esta análise é para fins informativos e educacionais. "
         "Não constitui recomendação de investimento. "
         "Consulte um profissional certificado (CFP/CNPI) antes de investir."),
    ]

    # Estilo para seções da aba de metadados
    estilo_secao = Font(name='Arial', bold=True, size=11, color='1F4E79')
    estilo_normal = Font(name='Arial', size=10)

    for idx, (chave, valor) in enumerate(metadados, start=1):
        aba_meta[f'A{idx}'] = chave
        aba_meta[f'B{idx}'] = valor

        # Seções em azul e negrito
        if valor == "":
            aba_meta[f'A{idx}'].font = estilo_secao
        else:
            aba_meta[f'A{idx}'].font = estilo_normal
            aba_meta[f'B{idx}'].font = estilo_normal

    aba_meta.column_dimensions['A'].width = 30
    aba_meta.column_dimensions['B'].width = 70

    # -------------------------------------------------------
    # SALVAR O ARQUIVO
    # -------------------------------------------------------
    try:
        wb.save(nome_arquivo)
        print(f"✅ Planilha salva: {nome_arquivo}")
        return True
    except Exception as e:
        print(f"❌ Erro ao salvar a planilha: {e}")
        return False


# ===========================================================
# CÉLULA 6 — Execução principal
# ===========================================================

def executar_analise():
    """
    Orquestra todas as etapas da análise:
    1. Coleta de dados
    2. Limpeza e padronização
    3. Aplicação dos filtros de qualidade
    4. Cálculo do score e ranking
    5. Geração da planilha Excel
    6. Download automático (se no Colab)
    """

    print("=" * 55)
    print("  📊 ANÁLISE DE FIIs — CARTEIRA PREVIDENCIÁRIA")
    print("=" * 55)

    # --- ETAPA 1: Coletar dados brutos do Fundamentus ---
    df_bruto = coletar_fiis_fundamentus()

    if df_bruto is None:
        print("\n❌ Análise interrompida: falha na coleta de dados.")
        return None

    # Pausa respeitosa entre operações
    time.sleep(1)

    # --- ETAPA 2: Limpar e padronizar ---
    df_limpo = limpar_e_padronizar(df_bruto)

    if df_limpo.empty:
        print("\n❌ Análise interrompida: nenhum dado após limpeza.")
        return None

    # --- ETAPA 3: Aplicar filtros de qualidade ---
    df_filtrado = aplicar_filtros_qualidade(df_limpo)

    if df_filtrado.empty:
        print("\n⚠️ Nenhum FII passou em todos os filtros.")
        print("   Considere afrouxar os critérios nas configurações.")
        return None

    # --- ETAPA 4: Calcular scores e ranking ---
    df_rankeado = calcular_score(df_filtrado)

    # --- ETAPA 5: Gerar planilha Excel ---
    sucesso = gerar_planilha_excel(df_rankeado, NOME_ARQUIVO)

    if not sucesso:
        return None

    # --- ETAPA 6: Download automático no Colab ---
    if RODANDO_NO_COLAB:
        print(f"\n⬇️  Iniciando download da planilha...")
        files.download(NOME_ARQUIVO)

    # --- Resumo final no terminal ---
    print("\n" + "=" * 55)
    print("  ✅ ANÁLISE CONCLUÍDA COM SUCESSO")
    print("=" * 55)
    print(f"\n🏆 TOP 10 FIIs por Score:")
    print(f"{'#':<5} {'Ticker':<10} {'DY':>8} {'P/VP':>7} {'Vacância':>10} {'Score':>7}")
    print("-" * 50)

    for _, linha in df_rankeado.head(10).iterrows():
        ticker    = str(linha.get('ticker',    '-'))
        dy        = linha.get('dy_anual',   0) or 0
        pvp       = linha.get('pvp',        0) or 0
        vacancia  = linha.get('vacancia',   None)
        score     = linha.get('score_final', 0) or 0
        ranking   = int(linha.get('ranking', 0))

        vacancia_str = f"{vacancia*100:.1f}%" if pd.notna(vacancia) else "N/A"

        print(f"{ranking:<5} {ticker:<10} {dy*100:>7.2f}% {pvp:>7.2f} {vacancia_str:>10} {score:>7.2f}")

    print(f"\n📁 Arquivo gerado: {NOME_ARQUIVO}")

    return df_rankeado


# --- Ponto de entrada ---
# Esta linha chama a função principal quando o script é executado
df_resultado = executar_analise()
