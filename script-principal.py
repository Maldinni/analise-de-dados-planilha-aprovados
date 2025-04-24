import pandas as pd
from collections import defaultdict


def carregar_planilha(caminho_arquivo):
    try:
        planilha_carregada = pd.read_excel(caminho_arquivo, sheet_name="UEMA1")
        print("Planilha carregada.")
        return planilha_carregada
    except:
        print("Erro. Planilha não carregada")
        return None

def encontrar_chave(planilha, clusters_termos):
    """
    Busca por clusters de termos na planilha.
    
    Args:
        planilha: DataFrame do pandas
        clusters_termos: Dicionário com {nome_do_cluster: [lista_de_termos]}
    
    Returns:
        Lista de dicionários com resultados, incluindo o nome do cluster
    """
    resultados = []
    
    for nome_cluster, termos in clusters_termos.items():
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

def verificacao_pinheiro(linha_planilha):
    """Verifica se 'Pinheiro' aparece APENAS na coluna de cidade (e não no nome/e-mail)."""
    cidade = str(linha_planilha['Cidade']).lower()  
    nome = str(linha_planilha['Nome completo do aluno']).lower()
    nome_social = str(linha_planilha['Nome social']).lower()
    email = str(linha_planilha['E-mail']).lower()
    
    return (
        'pinheiro' in cidade  # "Pinheiro" está na coluna da cidade
        and 'pinheiro' not in nome  # Não está no nome
        and 'pinheiro' not in nome_social  # Não está no nome social
        and 'pinheiro' not in email  # Não está no e-mail
    )

def verificacao_ribamar(linha_planilha):
    """Verifica se 'Ribamar' aparece APENAS na coluna de cidade (e não no nome/e-mail)."""
    cidade = str(linha_planilha['Cidade']).lower() 
    nome = str(linha_planilha['Nome completo do aluno']).lower()
    nome_social = str(linha_planilha['Nome social']).lower()
    email = str(linha_planilha['E-mail']).lower()
    
    return (
        'ribamar' in cidade  # "Ribamar" está na coluna da cidade
        and 'ribamar' not in nome  # Não está no nome
        and 'ribamar' not in nome_social  # Não está no nome social
        and 'ribamar' not in email  # Não está no e-mail
    )

def main():
    caminho = pd.ExcelFile("dados-leitura/Analise Inscrições Pâmela.xlsx")
    planilha = carregar_planilha(caminho)
    
    # Definindo seus clusters manualmente
    clusters_uni = {
        'PITAGORAS': ['pitágoras', 'pitagoras'],
        'ANHANGUERA': ['anhanguera'],
        'CEFET-MA': ['cefet-ma'],
        'CEUMA': ['ceuma', 'uniceuma'],
        'CEUPI': ['ceupi'],
        'CRUZEIRO DO SUL': ['cruzeiro'],
        'EDUFOR': ['edufor'],
        'ESTACIO': ['estacio'],
        'FAVYT': ['favyt'],
        'FEDERAÇÃO': ['federação das escolas superiores'],
        'FLORENCE': ['florence'],
        'IEMA': ['iema'],
        'IFMA': ['ifma'],
        'IFPA': ['ifpa'],
        'IFRB': ['ifrb'],
        'IFTMA': ['iftma'],
        'IPOG': ['ipog'],
        'ISFS': ['isfs'],
        'MACKENZIE': ['mackenzie'],
        'Senai': ['senai'],
        'UEMA': ['uema'],
        'UFCG': ['ufcg'],
        'UFMA': ['ufma'],
        'UFPA': ['ufpa'],
        'UFPI': ['ufpi'],
        'UFPR': ['ufpr'],
        'UFTMA': ['uftma'],
        'UNDB': ['undb'],
        'UNIFAEL': ['unifael'],
        'UNIFSA': ['unifsa'],
        'UNINASSAU': ['uninassau'],
        'UNISA': ['unisa'],
        'Unitins': ['unitins'],
        'Universidad de Pinar de Río': ['universidad de pinar del rio'],
        'UNV PARAGUAI': ['unv paraguai'],
        'Vazias': ['#N/D'],
        'AMPLA': ['ampla', 'concorrencia'],
        '45+': ['45+'],
        'MULHERES': ['mulheres'],
        'NPIQ': ['negros'],
        'PCD': ['pcds'],
        'SLZ': ['são luís'],
        'PINHEIRO': ['pinheiro'],
        'RIBAMAR': ['ribamar'],
        'CIDADES': ['bacabal', 'balsas', 'barra do corda', 'bom jesus das selvas',
                   'buriticupu', 'caxias', 'chapadinha', 'cruz das almas', 'cururupu',
                   'governador', 'grajaú', 'imperatriz', 'itapecuru mirim', 'paço do lumiar',
                   'pindaré-mirim', 'pio xii', 'recife', 'rosário', 'santa inês', 'são bento',
                   'nepomuceno', 'teresina', 'tianguá', 'timon'],
        'CURSOS': [''],
        'CLASSIFICADO': ['classificado', 'classificada', 'classificar']
    }

    # Busca por todos os clusters
    resultados_uni = encontrar_chave(planilha, clusters_uni)

    cidadaos_pinheiro = [
        item for item in resultados_uni 
        if item['Cluster'] == 'PINHEIRO' 
        and 'pinheiro' in str(item['Valor']).lower()
        and verificacao_pinheiro(planilha.iloc[item['Linha'] - 2])
    ]

    # Atualiza os resultados (remove CIDADES brutos e adiciona só os filtrados)
    resultados_uni = [
        item for item in resultados_uni 
        if item['Cluster'] != 'PINHEIRO'
    ] + cidadaos_pinheiro

    # Exibe apenas os resultados de Pinheiro (opcional)
    #print("\n=== OCORRÊNCIAS DA CIDADE PINHEIRO (VALIDADAS) ===")
    #for item in cidadaos_pinheiro:
    #    print(f"Linha {item['Linha']} | Coluna '{item['Coluna']}': {item['Valor']}")
    
    cidadaos_ribamar = [
        item for item in resultados_uni 
        if item['Cluster'] == 'RIBAMAR' 
        and 'ribamar' in str(item['Valor']).lower()
        and verificacao_ribamar(planilha.iloc[item['Linha'] - 2])
    ]

    resultados_uni = [
        item for item in resultados_uni 
        if item['Cluster'] != 'RIBAMAR'
    ] + cidadaos_ribamar

    # Agrupando resultados por cluster para análise
    resultados_por_cluster = defaultdict(list)
    
    for item in resultados_uni:
        resultados_por_cluster[item['Cluster']].append(item)
    
    total_classificados = 0

    # Exibindo resultados consolidados por cluster
    for cluster_uni, itens in resultados_por_cluster.items():
        print(f"\n=== CLUSTER: {cluster_uni} ===")
        print(f"Total de ocorrências: {len(itens)}")
        print("Termos incluídos:", ', '.join(clusters_uni[cluster_uni]))
        if cluster_uni == 'CLASSIFICADO':
            total_classificados = len(itens)

    print(total_classificados)
        
        #Laço genérico para printar todas as ocorrencias de cada termo no cluster por extenso
        #for item in itens:
            #print(f"Linha {item['Linha']} | Coluna '{item['Coluna']}': {item['Valor']}")
    
    # Análise de relacionamento entre clusters (exemplo: PITAGORAS e CLASSIFICADO)
    print("\n=== RELAÇÃO ENTRE CLUSTERS ===")
    
    # Obtém linhas com CLASSIFICADO
    linhas_classificado = {item['Linha'] for item in resultados_uni if item['Cluster'] == 'CLASSIFICADO'}
    #linhas_mulheres = {item['Linha'] for item in resultados_uni if item['Cluster'] == 'MULHERES'}
    #linhas_ampla = {item['Linha'] for item in resultados_uni if item['Cluster'] == 'AMPLA'}

    # Filtra resultados de PITAGORAS que estão nessas linhas
    pitagoras_classificado = [
        item for item in resultados_uni 
        if item['Cluster'] == 'PITAGORAS' and item['Linha'] in linhas_classificado
    ]
    
    mulheres_classificado = [
        item for item in resultados_uni 
        if item['Cluster'] == 'MULHERES' and item['Linha'] in linhas_classificado
    ]

    pcd_classificado = [
        item for item in resultados_uni 
        if item['Cluster'] == 'PCD' and item['Linha'] in linhas_classificado
    ]

    quarentaecinco_classificado = [
        item for item in resultados_uni 
        if item['Cluster'] == '45+' and item['Linha'] in linhas_classificado
    ]

    npiq_classificado = [
        item for item in resultados_uni 
        if item['Cluster'] == 'NPIQ' and item['Linha'] in linhas_classificado
    ]

    slz_classificado = [
        item for item in resultados_uni 
        if item['Cluster'] == 'SLZ' and item['Linha'] in linhas_classificado
    ]

    pinheiro_classificado = [
        item for item in resultados_uni 
        if item['Cluster'] == 'PINHEIRO' and item['Linha'] in linhas_classificado
    ]
    
    ribamar_classificado = [
        item for item in resultados_uni 
        if item['Cluster'] == 'RIBAMAR' and item['Linha'] in linhas_classificado
    ]

    cidades_classificado = [
        item for item in resultados_uni 
        if item['Cluster'] == 'CIDADES' and item['Linha'] in linhas_classificado
    ]

    ampla_classificado = [
        item for item in resultados_uni 
        if item['Cluster'] == 'AMPLA' and item['Linha'] in linhas_classificado
    ]

    #print(f"Ocorrências de PITAGORAS em linhas CLASSIFICADO: {len(pitagoras_classificado)}")
    #for item in pitagoras_classificado:
        #print(f"Linha {item['Linha']}: {item['Valor']}")

    print(f"Ocorrências de MULHERES em linhas CLASSIFICADO: {len(mulheres_classificado)}")
    for item in mulheres_classificado:
        print(f"Linha {item['Linha']}: {item['Valor']}")
    print(f"===================================================")

    print(f"Ocorrências de PCDS em linhas CLASSIFICADO: {len(pcd_classificado)}")
    for item in pcd_classificado:
        print(f"Linha {item['Linha']}: {item['Valor']}")
    print(f"===================================================")

    print(f"Ocorrências de 45+ em linhas CLASSIFICADO: {len(quarentaecinco_classificado)}")
    for item in quarentaecinco_classificado:
        print(f"Linha {item['Linha']}: {item['Valor']}")
    print(f"===================================================")

    print(f"Ocorrências de NEGROS, PARDOS, INDÍGENAS E QUILOMBOLAS em linhas CLASSIFICADO: {len(npiq_classificado)}")
    for item in npiq_classificado:
        print(f"Linha {item['Linha']}: {item['Valor']}")
    print(f"===================================================")

    print(f"Ocorrências de São Luís em linhas CLASSIFICADO: {len(slz_classificado)}")
    for item in slz_classificado:
        print(f"Linha {item['Linha']}: {item['Valor']}")
    print(f"===================================================")

    print(f"Ocorrências das demais cidades elegíveis em linhas CLASSIFICADO: {len(cidades_classificado)+len(pinheiro_classificado)+len(ribamar_classificado)}")
    for item in cidades_classificado:
        print(f"Linha {item['Linha']}: {item['Valor']}")
    for item in pinheiro_classificado:
        print(f"Linha {item['Linha']}: {item['Valor']}")
    for item in ribamar_classificado:
        print(f"Linha {item['Linha']}: {item['Valor']}")
    print(f"===================================================")

    #print(f"Ocorrências de AMPLA em linhas CLASSIFICADO: {len(ampla_classificado)}")
    #for item in ampla_classificado:
        #print(f"Linha {item['Linha']}: {item['Valor']}")

    #print(resultados_por_cluster)

if __name__ == "__main__":
    main()