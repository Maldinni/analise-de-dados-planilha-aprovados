# bibliotecas utilizadas
import pandas as pd
import streamlit as st
from collections import defaultdict
import seaborn as sns
import matplotlib.pyplot as plt
import numpy as np

st.set_page_config(page_title="Análise de Inscrições", layout="wide")

st.title("📊 Análise de Inscrições")

# carrega a planilha e a trata para manipulá-la no código com pandas
def carregar_planilha(uploaded_file):
    try:
        planilha_carregada = pd.read_excel(uploaded_file, sheet_name="UEMA1")
        st.success("✅ Planilha carregada com sucesso!")
        return planilha_carregada
    except:
        st.error("❌ Erro ao carregar a planilha.")
        return None

# função desenvolvida pra buscar "palavras-chave", mais especificamente, termos dos clusters
def encontrar_chave(planilha, clusters_termos):
    resultados = []
    for nome_cluster, termos in clusters_termos.items():
        if isinstance(termos, dict):
            termos = [t for sub in termos.values() for t in sub]
        for termo in termos:
            termo = str(termo).lower()
            for index, linha in planilha.iterrows():
                for coluna, valor in linha.items():
                    if termo in str(valor).lower():
                        resultados.append({
                            'Cluster': nome_cluster,
                            'Termo Encontrado': termo,
                            'Linha': index + 2,
                            'Coluna': coluna,
                            'Valor': valor
                        })
    return resultados

# abaixo duas validações simples para as cidades que apresentaram confusão na busca por seus termos figurarem tanto no nome, e-mail e cidade,
#  assim os blocos removem essa redundancia e retornam os resultados tratados
def verificacao_pinheiro(linha):
    cidade = str(linha['Cidade']).lower()  
    nome = str(linha['Nome completo do aluno']).lower()
    nome_social = str(linha['Nome social']).lower()
    email = str(linha['E-mail']).lower()
    return ('pinheiro' in cidade and 'pinheiro' not in nome and 'pinheiro' not in nome_social and 'pinheiro' not in email)

def verificacao_ribamar(linha):
    cidade = str(linha['Cidade']).lower()  
    nome = str(linha['Nome completo do aluno']).lower()
    nome_social = str(linha['Nome social']).lower()
    email = str(linha['E-mail']).lower()
    return ('ribamar' in cidade and 'ribamar' not in nome and 'ribamar' not in nome_social and 'ribamar' not in email)

# gera o heatmap com as correlações entre as categorias nomeadas nos blocos que também são os clusters que contem os termos pesquisados
# o heatmap foi gerado inspirado e guiado pelo artigo citado no README do repositório
def gerar_heatmap_correlacoes(resultados_uni, linhas_classificado):
    resultados_para_heatmap = [
        {**item, 'Cluster': 'DEMAIS CIDADES'} if item['Cluster'] in {'PINHEIRO', 'RIBAMAR'} else item 
        for item in resultados_uni
    ]

    categorias = [
        'PÚBLICAS', 'PRIVADAS', 'AMPLA', 'MULHERES', 'PCD', 
        '45+', 'NEGROS, PARDOS, INDÍGENAS E QUILOMBOLAS', 'SÃO LUÍS', 'DEMAIS CIDADES',
        'ENGENHARIAS (EXCETO TI)', 'CURSOS DE TECNOLOGIA DA INFORMAÇÃO (TI)',
        'CIÊNCIAS EXATAS/TECNOLÓGICOS (EXCETO TI)', 'FORMAÇÃO TÉCNICA'
    ]

    co_ocorrencias = pd.DataFrame(0, index=categorias, columns=categorias)

    for linha in linhas_classificado:
        clusters_linha = {
            item['Cluster'] for item in resultados_para_heatmap
            if item['Linha'] == linha and item['Cluster'] in categorias
        }
        for cat1 in clusters_linha:
            for cat2 in clusters_linha:
                co_ocorrencias.loc[cat1, cat2] += 1

    st.subheader("📍 Heatmap de Correlações entre candidatos classificados.")
    plt.figure(figsize=(12, 10))
    sns.heatmap(
        co_ocorrencias,
        annot=True,
        fmt='d',
        cmap='YlOrRd',
        linewidths=.5,
        cbar_kws={'label': 'Número de Alunos'}
    )
    st.pyplot(plt.gcf())
    plt.clf()

# ====== INTERFACE STREAMLIT ======

uploaded_file = st.file_uploader("📁 Envie a planilha .xlsx", type="xlsx")

if uploaded_file:
    planilha = carregar_planilha(uploaded_file)
    
    # validação simples para garantir que haja dados na planilha
    if planilha is not None:
        # definição manual dos cluster de acordo com o que é requisitado para as correlações e 
        # cada cluster possui os termos a serem pesquisados e suas variações
        clusters_uni = {

            # ===== clusters que separam as universidades públicas de privadas ===== 
            'PÚBLICAS': {
                'CEFET-MA': ['cefet-ma'], 'IFMA': ['ifma'], 'IFPA': ['ifpa'], 'IFRB': ['ifrb'],
                'IFTMA': ['iftma'], 'IEMA': ['iema'], 'UEMA': ['uema'], 'UFCG': ['ufcg'],
                'UFMA': ['ufma'], 'UFPA': ['ufpa'], 'UFPI': ['ufpi'], 'UFPR': ['ufpr'],
                'UFTMA': ['uftma'], 'Universidad de Pinar de Río': ['universidad de pinar del rio'],
                'UNV PARAGUAI': ['unv paraguai']
            },
            'PRIVADAS': {
                'PITAGORAS': ['pitágoras', 'pitagoras'], 'ANHANGUERA': ['anhanguera'],
                'CEUMA': ['ceuma', 'uniceuma'], 'CEUPI': ['ceupi'], 'CRUZEIO DO SUL': ['cruzeiro'],
                'EDUFOR': ['edufor'], 'ESTACIO': ['estacio'], 'FAVYT': ['favyt'],
                'FEDERAÇÃO': ['federação das escolas superiores'], 'FLORENCE': ['florence'],
                'IPOG': ['ipog'], 'ISFS': ['isfs'], 'MACKENZIE': ['mackenzie'],
                'SENAI': ['senai'], 'UNDB': ['undb'], 'UNIFAEL': ['unifael'],
                'UNIFSA': ['unifsa'], 'UNINASSAU': ['uninassau'], 'UNISA': ['unisa'],
                'UNITINS': ['unitins']
            },

            # ===== clusters que definem cotas e ampla concorrência =====
            'AMPLA': ['ampla', 'concorrencia'],
            '45+': ['45+'],
            'MULHERES': ['mulheres'],
            'NEGROS, PARDOS, INDÍGENAS E QUILOMBOLAS': ['negros'],
            'PCD': ['pcds'],

            # ===== clusters que definem São Luís, a cidade matriz, e as demais cidades (pinheiro e são josé de ribamar estão separados devido a suas particularidades explicadas nos blocos de validação) =====
            'SÃO LUÍS': ['são luís'],
            'PINHEIRO': ['pinheiro'],
            'RIBAMAR': ['ribamar'],
            'DEMAIS CIDADES': ['bacabal', 'balsas', 'barra do corda', 'bom jesus das selvas', 'buriticupu', 
                        'caxias', 'chapadinha', 'cruz das almas', 'cururupu', 'governador', 'grajaú',
                        'imperatriz', 'itapecuru mirim', 'paço do lumiar', 'pindaré-mirim', 'pio xii',
                        'recife', 'rosário', 'santa inês', 'são bento', 'nepomuceno', 'teresina', 
                        'tianguá', 'timon'],

            # ===== clusters de área de conhecimento dos classificados =====
            'ENGENHARIAS (EXCETO TI)': [
                'Cursos de Engenharia (exceto os de Engenharia de Software, Engenharia de Computação ou similares)'
            ],
            
            'CURSOS DE TECNOLOGIA DA INFORMAÇÃO (TI)': [
                'Cursos de Engenharia de Software, Engenharia de Computação, Ciência da Computação, Sistemas de Informação, Análise de Sistemas ou similares'
            ],
            
            'CIÊNCIAS EXATAS/TECNOLÓGICOS (EXCETO TI)': [
                'Outros cursos de ciências exatas ou tecnológicos (exceto os de Engenharia de Software, Engenharia de Computação, Ciência da Computação, Sistemas de Informação, Análise de Sistemas ou similares)'
            ],
            
            'FORMAÇÃO TÉCNICA': [
                'Formados em cursos técnicos de nível médio na área de Ciências Exatas'
            ],

            # ===== cluster para pesquisar os alunos classificados =====
                'CLASSIFICADO': ['classificado', 'classificada', 'classificar']
            }

        resultados = encontrar_chave(planilha, clusters_uni)

        # tratamento dos dados retornados após a validação, chamam a remoção de redundância e 
        # retornam para a variável resultados apenas os termos importantes 
        cidadaos_pinheiro = [
            item for item in resultados 
            if item['Cluster'] == 'PINHEIRO' 
            and 'pinheiro' in str(item['Valor']).lower()
            and verificacao_pinheiro(planilha.iloc[item['Linha'] - 2])
        ]

        resultados = [item for item in resultados if item['Cluster'] != 'PINHEIRO'] + cidadaos_pinheiro

        cidadaos_ribamar = [
            item for item in resultados 
            if item['Cluster'] == 'RIBAMAR' 
            and 'ribamar' in str(item['Valor']).lower()
            and verificacao_ribamar(planilha.iloc[item['Linha'] - 2])
        ]

        resultados = [item for item in resultados if item['Cluster'] != 'RIBAMAR'] + cidadaos_ribamar

        linhas_classificado = {item['Linha'] for item in resultados if item['Cluster'] == 'CLASSIFICADO'}

        # Gera o heatmap após inserir a planilha
        st.subheader("📌 Total de alunos classificados:")
        st.metric("Classificados", len(linhas_classificado))

        gerar_heatmap_correlacoes(resultados, linhas_classificado)