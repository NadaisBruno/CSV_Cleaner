import pandas as pd
from tkinter import Tk, filedialog
from openpyxl.styles import Font, PatternFill  # sub modulo onde vivem as classes para formata o texto (negrito, etc)
                                  # Styles e o modulo que contem Font,Alignment,PatternFill

def escolher_csv():
    janela = Tk()  # cria a janela do Tkinter
    janela.withdraw()  # Esconde a janela principal do Tkinter
    # Pedir ao utilizador que ficheiro quer abrir
    path = filedialog.askopenfilename(
        title="Escolher um ficheiro CSV",
        defaultextension=".csv",  # se o utilizador não escrever a extensão o programa acrescenta automaticamente
        filetypes=[("csv", "*.csv")]  # mostra apenas os ficheiros .csv
    )
    if not path:
        return None

    return path


def ler_csv(path):  # Lemos o csv
    # se o utilizador nao escolheu o ficheiro cancela a funcao e nao tenta ler nada
    if path is None:
        return None
    # verificamos se o utilizador escolheu o ficheiro com a extensao .csv
    if not path.lower().endswith(".csv"):
        print("Erro: ficheiro selecionado não é um CSV.")
        return None

    df = pd.read_csv(path, encoding="latin1")

    if df.empty:  # se o DataFrame estiver vazio(csv vazio) devolve None para o programa nao crashar
        return None

    if df.columns.size == 0:  # se o CSV nao conter colunas devolve None para o programa nao chashar
        print("Erro: o CSV não contém colunas.")
        return None

    if df.shape[1] == 1:  # df shape[1] representa o número de colunas do CSV // Podiamos usar df.columns == 1, mas nao compensa
        print("Erro: o CSV só contem uma coluna.")
        return None

    # Verificamos se o csv tem as acolunas certas e obrigatorias
    colunas_obrigatorias = ["data", "produto", "quantidade", "preco"]
    for i in colunas_obrigatorias:
        if i not in df.columns:
            print(f"Erro: o csv nao contem a coluna obrigatoria: {i}")
            return None

    # Converte a coluna "data" para o tipo datetime.
    # Qualquer valor que não seja uma data válida é transformado em NaT (Not a Time)(data inválida),
    # o que nos permite detetar e tratar datas erradas no CSV e assim o programa nao crashar
    df["data"] = pd.to_datetime(df["data"], errors="coerce")  # coerce significa que o programa nao vai abaixo e, em vez, disso transforma em NaT
    # Verifica se existe alguma data inválida na coluna "data".
    # 'isna()' marca cada linha como True se for inválida (NaT, vazia, erro na conversão).
    # 'any()' devolve True se existir pelo menos uma linha inválida.
    if df["data"].isna().any():
        print("Erro: existem datas invalidas no CSV")
        return None

    # validamos se os valores sao ...
    if df["produto"].isna().any():
        print("Erro: existem valores invalidos na coluna produto do CSV")
        return None
    # str.strip() remove espacos antes e depois // =="" deteta ‘strings’ vazia // any() verifica se pelo menos um e vazio
    if (df["produto"].str.strip() == "").any():
        print("Erro: existem colunas vazias na coluna produto do CSV")
        return None
    # Verifica se algum produto é composto APENAS por dígitos.
    # isdigit() → True se a string for só números
    # any() → True se pelo menos um valor for só números
    # Produtos só com números normalmente indicam erro no CSV.
    if df["produto"].str.isdigit().any():
        print("Erro: existem produtos compostos apenas por numeros na coluna produto do CSV")
        return None
    # Verifica se algum produto contém caracteres proibidos (@, #, $, %, ?, !, /, ~)
    # str.contains() procura no texto usando regex; o padrão [@#$%?!/~] significa "qualquer caractere destes"
    # r"" → diz ao Python que é um raw ‘string’, necessário para regex(regular expressions). Regex é uma mini-linguagem para procurar padrões dentro de texto.
    if df["produto"].str.contains(r"[@#$%?!/~]").any():
        print("Erro: existem caracteres invalidos(@, #, $, %, ?, !, /, ~) na coluna produto do CSV")
        return None

    # validamos se os valores sao numericos
    df["quantidade"] = pd.to_numeric(df["quantidade"], errors="coerce")  # to_numeric converte tudo para numero
    if df["quantidade"].isna().any():
        print("Erro: existem valores invalidos na coluna quantidade do CSV")
        return None

    df["preco"] = pd.to_numeric(df["preco"], errors="coerce")
    if df["preco"].isna().any():
        print("Erro: existem valores invalidos na coluna preço do CSV")
        return None

    # validamos se os valores nao sao negativos
    if (df["quantidade"] <= 0).any():
        print("Erro: não podem existir valores negativos na coluna quantidade do CSV")
        return None

    if (df["preco"] <= 0).any():
        print("Erro: não podem existir valores negativos na coluna preço do CSV")
        return None

    # validamos se os valores sao inteiros e nao floats
    if (df["quantidade"] % 1 != 0).any():  # % 1 != 0 significa se tem numeros decimais/any verifica se existe ao menos um valor decimal
        print("Erro: não podem existir valores decimais na coluna quantidade do CSV")
        return None

    print(df.head())
    print(df.shape)  # mostra o tamanho do Dataframe, por exemplo (100(linhas), 5(colunas))
    return df


def limpar_dados(df):
    df = df.dropna()  # Remove linhas com valores ausentes(NaN). Sem este passo, CSVs com linhas "mortas" podem fazer o programa crashar ou gerar relatórios errados.
    df = df.drop_duplicates()  # Remove linhas repetidas/evita contagens duplicadas

    # remover espacos na coluna produtos
    df["produto"] = df["produto"].str.strip().str.lower()  # str.strip() so funciona em strings

    # removemos caracteres especiais
    carateres_especiais_remover = r"[@#$%?!/~]"
    df["produto"] = df["produto"].str.replace(carateres_especiais_remover, "", regex=True)

    # Remover espacos nas colunas quantidade e preco
    # O astype(str) garante que todos os valores destas colunas são tratados como texto,
    # permitindo aplicar .str.strip() sem erros.
    # Isto é necessário porque, na validação la em cima na def ler_csv(), já converti estas colunas para número,
    # e números não têm métodos de ‘string’.
    df["quantidade"] = df["quantidade"].astype(str).str.strip()
    df["preco"] = df["preco"].astype(str).str.strip()

    # aqui convertemos novamente para número com pd.to_numeric, para garantir que ficam valores numéricos
    # válidos prontos para cálculos mais tarde.
    df["quantidade"] = pd.to_numeric(df["quantidade"])
    df["preco"] = pd.to_numeric(df["preco"])

    # converter data para datetime
    df["data"] = pd.to_datetime(df["data"])

    # ordenar a coluna data
    df = df.sort_values(by="data")

    # reset ao index par ficar ordenado
    df = df.reset_index(drop=True)

    print("\nDados Limpos: ")
    print(df.head())
    print(df.shape)
    return df  # Devolve o dataframe limpo


def gerar_relatorio(df):
    linhas = df.shape[0]
    colunas = df.shape[1]
    total_quantidade = df["quantidade"].sum()  # somamos com sum() todos os valores da coluna quantidade
    total_preco = df["preco"].sum()  # somamos com sum() todos os valores da coluna preco

    # Criamos uma coluna chamada "total_linha"
    # Esta coluna representa o total por cada linha = quantidade * preço
    # Exemplo: 3 camisolas a 10 € = 30 €
    df["total_linha"] = df["quantidade"] * df["preco"]
    # Soma total da coluna total_linha
    total_faturado = df["total_linha"].sum()

    # Calcular a media do preco // mean() soma todos os valores e divide pelo número de linhas
    media_preco = df["preco"].mean()
    # Calcular a média da quantidade // # Calculamos a média da quantidade — indica o valor médio por linha.
    media_quantidade = df["quantidade"].mean()

    # calcular o produto mais vendido//idxmax() devolve o nome do produto(leite,etc) com o valor maximo
    produto_mais_vendido = df.groupby("produto")["quantidade"].sum().idxmax()
    quantidade_maxima = df.groupby("produto")["quantidade"].sum().max()  # Aqui devolve o valor(75 unidades). Maior quantidade vendida entre todos os produtos

    # calcular o produto mais caro
    produto_mais_caro = df.groupby("produto")["preco"].sum().idxmax()  # idxmax() devolve o nome do produto com maior soma de preços
    preco_maximo = df.groupby("produto")["preco"].sum().max()  # Aqui devolve o valor

    # data mais antiga/recente
    primeira_data = df["data"].min()
    ultima_data = df["data"].max()

    # produtos unicos = quantos produtos diferentes existem no CSV
    produtos_unicos = df["produto"].nunique()

    # produto vendido uma unica vez
    quantidades_por_produto = df.groupby("produto")["quantidade"].sum()
    # Selecionamos apenas os produtos cuja soma total é igual a 1.
    # Ou seja, produtos vendidos exatamente uma vez.
    # .index devolve o nome dos produtos e .tolist() transforma em lista normal.
    # (não podemos usar idxmax porque aqui não existe "valor máximo", mas sim vários possíveis valores iguais a 1)
    produtos_vendidos_uma_vez = quantidades_por_produto[quantidades_por_produto == 1].index.tolist()

    # Calcula o top 3 produtos mais vendidos:
    # 1) soma as quantidades por produto,
    # 2) ordena do maior para o menor,
    # 3) seleciona os 3 mais vendidos.
    top_3_produtos = df.groupby("produto")["quantidade"].sum().sort_values(ascending=False).head(3)

    linhas_top3 = []  # Criamos uma lista vazia
    posicao = 1

    for produto, quantidade in top_3_produtos.items():
        linhas_top3.append(f"{posicao}º {produto} - {quantidade} unidades")
        posicao += 1

    top_3_produtos_formatado = " | ".join(linhas_top3)  # " | " separa o top 3 com pipe(|)

    # percentagens que cada produto representa no total
    percentagens_por_produto = df.groupby("produto")["quantidade"].sum().sort_values(ascending=False)

    linhas_percentagens = []

    for produto, quantidade in percentagens_por_produto.items():
        calculo_percentagem = (quantidade / total_quantidade) * 100
        linhas_percentagens.append(f"{produto}: {calculo_percentagem:.2f}%")

    percentagens_por_produto_formatado = " | ".join(linhas_percentagens)

    nomes_colunas = list(df.columns)

    nomes_colunas_formatado = " | ".join(df.columns
                                         )
    # Criamos um dicionário com todos os resultados para enviar para o Excel depois
    relatorio = {
        "linhas": linhas,
        "colunas": colunas,
        "total_quantidade": total_quantidade,
        "soma_precos_unitarios": total_preco,
        "total_faturado": total_faturado,
        "media_preco": media_preco,
        "media_quantidade": media_quantidade,
        "produto_mais_vendido": produto_mais_vendido,
        "quantidade_maxima": quantidade_maxima,
        "produto_mais_caro": produto_mais_caro,
        "preco_maximo": preco_maximo,
        "primeira_data": primeira_data,
        "ultima_data": ultima_data,
        "produtos_unicos": produtos_unicos,
        "produtos_vendidos_uma_vez": produtos_vendidos_uma_vez,
        "top_3_produtos": top_3_produtos_formatado,
        "percentagens": percentagens_por_produto_formatado,
        "nomes_colunas": nomes_colunas_formatado
    }
    return relatorio


def exportar_excel(df_limpo, relatorio):
    # Transformamos o relatorio em um DataFrame
    df_relatorio = pd.DataFrame.from_dict(relatorio, orient="index")  #

    with pd.ExcelWriter("relatorio.xlsx", engine="openpyxl") as writer:
        df_limpo.to_excel(writer, sheet_name="Dados_limpos", index=False)  # index=False elimina a coluna com o nome index que iria aparecer no excel
        df_relatorio.to_excel(writer, sheet_name="Relatorio")
        ws_dados = writer.book["Dados_limpos"]
        ws_relatorio = writer.book["Relatorio"]

        # Cabecalhos a negrito na folha de dados limpos
        for i in ws_dados[1]:  # devolve a primeira linha que e o cabecalho
            i.font = Font(bold=True)

        # Cabecalhos a negrito na folha de relatorio
        for i in ws_relatorio[1]:
            i.font = Font(bold=True)

        # ---------------ajustar a largura das colunas na folha "Dados_limpos"-------------------------------------
        for coluna in ws_dados.columns:  # itera sobre cada coluna
            nome_coluna = coluna[0].value  # por exemplo quantidade, preco, produto

            # largura baseada no tamanho do nome da coluna + margem
            largura = len(str(nome_coluna)) + 5

            # Define a largura da coluna usando a letra do Excel (A, B, C, ...)
            ws_dados.column_dimensions[coluna[0].column_letter].width = largura

        # ---------------ajustar a largura das colunas na folha "Relatorio"-------------------------------------
        for coluna in ws_relatorio.columns:
            nome_coluna = coluna[0].value

            largura = len(str(nome_coluna)) + 5

            ws_relatorio.column_dimensions[coluna[0].column_letter].width = largura

        # --------- FORMATAR CABEÇALHO DAS DUAS FOLHAS ----------------
        # Criamos um preenchimento azul-claro para o fundo
        fundo_azul = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")

        # Criamos um estilo de letra branca e a negrito
        letra_branca = Font(color="FFFFFF", bold=True)

        # Aplicar estilos à primeira linha (linha 1) da folha 'Dados_limpos'
        for cell in ws_dados[1]:  # devolve todas as celulas na linha 1
            cell.fill = fundo_azul  # pinta o fundo azul-claro
            cell.font = letra_branca

        # Aplicar estilos à primeira linha (linha 1) da folha 'Relatorio'
        for cell in ws_relatorio[1]:
            cell.fill = fundo_azul
            cell.font = letra_branca





def main():
    caminho = escolher_csv()
    # se o utilizador cancelar a escolha na janela,a def escolher_csv() devolve None e termina o programa sem crashar
    if caminho is None:
        return "Nenhum ficheiro foi selecionado"
    # le o ficheiro csv escolhido e transforma-o num DataFrame do pandas
    df = ler_csv(caminho)
    if df is None:
        return "Erro: ficheiro selecionado não é um CSV.Certifique-se de que escolheu um ficheiro válido."
    # limpa os dados: remove linhas vazia, repetidas, etc
    df_limpo = limpar_dados(df)
    # gera um relatorio com informacoes do DataFrame(numero de linhas, colunas e nomes das colunas)
    relatorio = gerar_relatorio(df_limpo)
    # exportamos os dados limpos + relatorio para um unico ficheiro excel com duas folhas
    exportar_excel(df_limpo, relatorio)
    print("Relatório criado com sucesso!")


# Este bloco so e executado se este ficheiro for executado diretamente. Serve como ponto de entrada principal do programa
if __name__ == "__main__":
    # Chama a funcao main() para iniciar o processo(gatilho para arrancar o programa)
    main()

