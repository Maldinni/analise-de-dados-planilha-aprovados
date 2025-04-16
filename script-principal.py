import pandas as pd

def carregar_planilha(caminho_arquivo):
    try:
        planilha_carregada = pd.read_excel(caminho_arquivo, sheet_name="UEMA1")
        print("Planilha carregada.")
        return planilha_carregada
    except:
        print("Erro. Planilha não carregada")
        return None

def encontrar_chave(planilha, lista_termos):
    resultados = []

    for termo in lista_termos:
        termo = str(termo).lower()
        for index, linha in planilha.iterrows():
            for coluna, valor in linha.items():
                if termo in str(valor).lower():
                    resultados.append({'Termo': termo,
                                    'Linha': index + 2,
                                    'Coluna': coluna,
                                    'Valor': valor})
    return resultados

def main():
    caminho = pd.ExcelFile("/home/lincprog/Documentos/Enzo/Dados-planilha-BRISA/dados-leitura/Analise Inscrições Pâmela.xlsx")

    # if teste is not None:
    #     print("sucesso")
    # else:
    #     print("erro")

    termo = ['classificado']
    resultados = encontrar_chave(carregar_planilha(caminho), termo)

    if resultados:
        print(f"\nForam encontradas {len(resultados)} ocorrências de '{termo}':")
        for item in resultados:
            print(f"Linha {item['Linha']}, Coluna '{item['Coluna']}': {item['Valor']}")
    else:
        print(f"\nNenhum resultado encontrado para '{termo}'.")

    contagem = 0

    for item in resultados:
        contagem += 1
    
    print(contagem)

if __name__ == "__main__":
    main()