import pandas as pd
from datetime import datetime
from openpyxl.styles import Font, PatternFill, numbers, Alignment, Side, Border  # submodulo onde vivem as classes para formata o texto (negrito, etc), Styles e o modulo que contem Font,Alignment,PatternFill,numbers
from openpyxl.worksheet.table import Table, TableStyleInfo
import streamlit as st


def registar_erro(mensagem):
    # Esta função regista erros num ficheiro chamado "log.txt".
    # Permite guardar um histórico de erros que acontecem durante o programa.
    # Isto permite diagnosticar problemas, ajudar utilizadores

    formato = "%d-%m-%Y %H:%M:%S"
    data_hora_actual = datetime.now().strftime(formato)  # a que data e hora o erro ocorreu

    # Abrimos o ficheiro "log.txt" em modo "append" (a).
    # Isto significa que cada novo erro será acrescentado no fim do ficheiro, sem apagar o que já está escrito
    with open("log.txt", "a", encoding="utf8") as file:
        file.write(f"{data_hora_actual} - {mensagem}\n")


# Esta função permite mostrar mensagens de erro ao utilizador de forma clara
# e, simultaneamente, registar o erro no ficheiro log.txt.
# Assim evitamos repetir vários prints e chamadas ao registar_erro() em cada validação.
# Sempre que houver um erro no CSV, chamamos erro_usuario(mensagem, sugestao)
# para manter o código limpo, organizado e profissional.
def erro_usuario(mensagem, sugestao):
    st.error(f"Erro: {mensagem}")
    st.error(f"→ {sugestao}")
    registar_erro(mensagem)


def ler_csv(uploaded_file):  # Lemos o csv
    # se o utilizador nao escolheu o ficheiro cancela a funcao e nao tenta ler nada
    if uploaded_file is None:
        return None
    # verificamos se o utilizador escolheu o ficheiro com a extensao .csv
    if not uploaded_file.name.lower().endswith(".csv"):
        erro_usuario(
            "Ficheiro selecionado não é um CSV.",
            "O ficheiro tem de conter a extensão CSV(exemplo: xxxx.csv)."
        )
        return None
    try:
        df = pd.read_csv(uploaded_file, encoding="latin1", sep=None, engine="python")
    except Exception as e:
        erro_usuario(
            "Erro ao ler o ficheiro CSV.",
            f"O ficheiro pode estar corrompido ou com formato inválido. Detalhes: {e}"
        )
        return None

    if df.empty:  # se o DataFrame estiver vazio(csv vazio) devolve None para o programa nao crashar
        erro_usuario(
            "O CSV está vazio.",
            "Certifique-se que o CSV contem pelo menos uma linha de dados."
        )
        return None

    if df.columns.size == 0:  # se o CSV nao conter colunas devolve None para o programa nao chashar
        erro_usuario(
            "O CSV não contém colunas.",
            "Certifique-se que o ficheiro tem cabeçalhos(ex:data,produto,quantidade,preco)."
        )
        return None

    if df.shape[1] == 1:  # df shape[1] representa o número de colunas do CSV // Podiamos usar df.columns == 1, mas nao compensa
        erro_usuario(
            "O CSV só contém uma coluna.",
            "Verifique se o ficheiro está separado corretamente por vírgulas e não por outro separador."
        )
        return None

    # Verificamos se o csv tem as acolunas certas e obrigatorias
    colunas_obrigatorias = ["data", "produto", "quantidade", "preco"]
    for i in colunas_obrigatorias:
        if i not in df.columns:
            erro_usuario(
                f"O CSV não contém a coluna obrigatória: {i}.",
                "O ficheiro deve conter as colunas: data, produto, quantidade, preco."
            )
            return None

    # Converte a coluna "data" para o tipo datetime.
    # Qualquer valor que não seja uma data válida é transformado em NaT (Not a Time)(data inválida),
    # o que nos permite detetar e tratar datas erradas no CSV e assim o programa nao crashar
    df["data"] = pd.to_datetime(df["data"], errors="coerce")  # coerce significa que o programa nao vai abaixo e, em vez, disso transforma em NaT
    # Verifica se existe alguma data inválida na coluna "data".
    # 'isna()' marca cada linha como True se for inválida (NaT, vazia, erro na conversão).
    # 'any()' devolve True se existir pelo menos uma linha inválida.
    if df["data"].isna().any():
        erro_usuario(
            "Existem datas inválidas no CSV",
            "Certifique-se que as datas estão no formato válido(DD-MM-AAAA ou YYYY-MM-DD)"
        )
        return None

    # validamos se os valores sao ...
    if df["produto"].isna().any():
        erro_usuario(
            "Existem valores inválidos na coluna produto do CSV",
            "Certifique-se que o campo 'produto' não contem valores numéricos"
        )
        return None
    # str.strip() remove espacos antes e depois // == "" deteta ‘strings’ vazia // any() verifica se pelo menos um e vazio
    if (df["produto"].str.strip() == "").any():
        erro_usuario(
            "Existem produtos vazios na coluna 'produto'.",
            "Preencha todos os nomes de produtos e remova linhas onde o produto esteja vazio"
        )
        return None
    # Verifica se algum produto é composto APENAS por dígitos.
    # isdigit() → True se a string for só números
    # any() → True se pelo menos um valor for só números
    # Produtos só com números normalmente indicam erro no CSV.
    if df["produto"].str.isdigit().any():
        erro_usuario(
            "Existem produtos compostos apenas por numeros na 'coluna' produto do CSV",
            "Os nomes dos produtos devem conter texto. Substitua códigos numéricos por nomes reais."
        )
        return None
    # Verifica se algum produto contém caracteres proibidos (@, #, $, %, ?, !, /, ~)
    # str.contains() procura no texto usando regex; o padrão [@#$%?!/~] significa "qualquer caractere destes"
    # r"" → diz ao Python que é um raw ‘string’, necessário para regex(regular expressions). Regex é uma mini-linguagem para procurar padrões dentro de texto.
    if df["produto"].str.contains(r"[@#$%?!/~]").any():
        erro_usuario(
            "Existem caracteres inválidos(@, #, $, %, ?, !, /, ~) na coluna 'produto'.",
            "Remova estes caracteres especiais dos nomes dos produtos para garantir um formato válido."
        )
        return None

    # validamos se os valores sao numericos
    df["quantidade"] = pd.to_numeric(df["quantidade"], errors="coerce")  # to_numeric converte tudo para numero
    if df["quantidade"].isna().any():
        erro_usuario(
            "Existem valores não numéricos na coluna 'quantidade'.",
            "Certifique-se de que todos os valores desta coluna são números inteiros (ex.: 1, 2, 3)."
        )
        return None

    df["preco"] = pd.to_numeric(df["preco"], errors="coerce")
    if df["preco"].isna().any():
        erro_usuario(
            "Existem valores não numéricos na coluna 'preco'.",
            "Certifique-se de que todos os valores desta coluna são números válidos, usando ponto como separador decimal (ex.: 2.50)"
        )
        return None

    # validamos se os valores nao sao negativos
    if (df["quantidade"] <= 0).any():
        erro_usuario(
            "Existem valores negativos ou zero na coluna 'quantidade'.",
            "A quantidade deve ser sempre um número inteiro maior que zero (ex.: 1, 2, 3)."
        )
        return None

    if (df["preco"] <= 0).any():
        erro_usuario(
            "Existem valores negativos ou zero na coluna 'preco'.",
            "O preço deve ser um número maior que zero. Pode incluir casas decimais (ex.: 2.50)."
        )
        return None

    # validamos se os valores sao inteiros e nao floats
    if (df["quantidade"] % 1 != 0).any():  # % 1 != 0 significa se tem numeros decimais/any verifica se existe ao menos um valor decimal
        erro_usuario(
            "Não podem existir valores decimais na coluna 'quantidade'.",
            "A quantidade deve ser um número inteiro (ex.: 1, 2, 3). Remova valores como 1.5 ou 2,7."
        )
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

    nomes_colunas_formatado = " | ".join(df.columns)
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
    relatorio = {k: str(v) for k, v in relatorio.items()}  # k=key(por exemplo, top_3_produtos, v=value(por exemplo, arroz, massa etc)//items() percorre tudo no dicionario chave/valor
    # Transformamos o relatorio em um DataFrame
    # Criamos o DataFrame do relatório com duas colunas fixas:
    #   Metrica | Valor
    # Este formato é 100% compatível com Tabelas Oficiais do Excel,
    # porque garante:
    #   - cabeçalhos válidos (‘strings’, sem caracteres especiais)
    #   - estrutura tabular real de 2 colunas
    #   - ausência de listas ou objetos não suportados
    # Por isso NÃO é necessário aplicar limpeza aos nomes das colunas.
    df_relatorio = pd.DataFrame(list(relatorio.items()), columns=["Metrica", "Valor"])



    with pd.ExcelWriter("relatorio.xlsx", engine="openpyxl") as writer:
        df_limpo.to_excel(writer, sheet_name="Dados_limpos", index=False)  # index=False elimina a coluna com o nome index que iria aparecer no excel
        df_relatorio.to_excel(writer, sheet_name="Relatorio", index=False)
        ws_dados = writer.book["Dados_limpos"]
        ws_relatorio = writer.book["Relatorio"]

        # Congelar a primeira linha de ambas as folhas (mantem o cabeçalho) - se o cliente fazer ‘scroll’ a linha do cabecalho fica sempre à vista
        ws_dados.freeze_panes = "A2"  # A2 significa que tudo a cima e á esquerda desta celula fica congelado
        ws_relatorio.freeze_panes = "A2"

        #  AUTOAJUSTE DA LARGURA DAS COLUNAS
        #  Percorremos todas as colunas da folha e determinamos o tamanho
        #  máximo do conteúdo de cada coluna (incluindo o cabeçalho). Depois
        #  aplicamos uma largura proporcional no Excel. Isto evita textos
        #  cortados, valores escondidos e melhora muito a legibilidade.
        #  Este método torna o ficheiro final profissional e evita que o
        #  utilizador tenha de ajustar colunas manualmente.
        #  Autoajuste das colunas em dados_limpos
        for coluna in ws_dados.columns:  # ws_dados.columns devolve todas as colunas
            letra = coluna[0].column_letter  # coluna[0] primeira celula dessa coluna // column_letter devolve A, B, C, etc
            ws_dados.column_dimensions[letra].width = 20
        #  Autoajuste das colunas em relatorio
        for coluna in ws_relatorio.columns:
            letra = coluna[0].column_letter
            ws_relatorio.column_dimensions[letra].width = 25

        # -------------------- Bordas Finas -----------------------------------
        # Criamos um estilo de borda
        bordas_finas = Border(
            left=Side(style="thin"),  # borda esquerda fina
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )
        # ------------------- Aplicar bordas finas na folha 'Dados_limpos' ---------------------
        for linha in ws_dados.iter_rows(
            min_row=1,  # comecamos na linha 1 que inclui o cabecalho
            max_row=ws_dados.max_row,  # ultima linha que contem dados
            min_col=1,  # comecamos na primeira coluna (A)
            max_col=ws_dados.max_column  # ultima coluna que contem dados
        ):
            #  percorremos cada celula dentro da linha
            for cell in linha:
                cell.border = bordas_finas

        # ------------------- Aplicar bordas finas na folha 'Relatorio' --------------------
        for linha in ws_relatorio.iter_rows(
            min_row=1,
            max_row=ws_relatorio.max_row,
            min_col=1,
            max_col=ws_relatorio.max_column
        ):
            for cell in linha:
                cell.border = bordas_finas

        # ----------------Criar tabela oficial na folha 'Dados_limpos'--------------------------
        # CRIAÇÃO DE TABELAS OFICIAIS DO EXCEL (Structured Tables)
        #
        # Transformamos o intervalo de dados de cada folha numa "Tabela Oficial"
        # do Excel. Diferente de simples células, estas tabelas têm funções
        # avançadas e tornam o relatório muito mais profissional.
        #
        # VANTAGENS PRINCIPAIS:
        #   - Filtros automáticos no cabeçalho (ordenar, filtrar, pesquisar)
        #   - Formatação profissional com linhas alternadas e cabeçalho destacado
        #   - A tabela expande automaticamente se o cliente adicionar novas linhas
        #   - Fórmulas ficam estruturadas (ex.: [@quantidade] * [@preco])
        #   - Melhor compatibilidade com Power BI, Power Query e automações
        #   - Visual muito mais limpo e apresentável para clientes
        # NOTA:
        #   O nome da tabela (displayName) não pode ter espaços ou acentos.
        #   O intervalo (ref) precisa de ser válido e conter cabeçalhos em texto.
        # Em resumo:
        # Criar Tabelas Oficiais torna o Excel inteligente, dinâmico e pronto para
        # uso empresarial — um grande diferencial para um produto de freelancing.
        # range completo da tabela, por exemplo: "A1:D150"
        tabela_range_dados = ws_dados.dimensions  # calcula automaticamente o tamanho da tabela

        # criamos a tabela com um nome unico
        tabela_dados = Table(displayName="Tabela_Dados_limpos", ref=tabela_range_dados)

        # estilo da tabela(cores alternadas + cabecalho especial)
        estilo = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=True
        )
        tabela_dados.tableStyleInfo = estilo
        ws_dados.add_table(tabela_dados)

        # ----------------Criar tabela oficial na folha 'Relatorio'--------------------------
        tabela_range_relatorio = ws_relatorio.dimensions

        tabela_relatorio = Table(displayName="Tabela_Relatorio", ref=tabela_range_relatorio)
        tabela_relatorio.tableStyleInfo = estilo
        ws_relatorio.add_table(tabela_relatorio)


        # Cabecalhos a Negrito na folha de dados limpos
        for i in ws_dados[1]:  # devolve a primeira linha que e o cabecalho
            i.font = Font(bold=True)

        # Cabecalhos a Negrito na folha de relatorio
        for i in ws_relatorio[1]:
            i.font = Font(bold=True)

        # ------------------- Formatacao dinamica de colunas importantes----------------------------------------
        # Quando exportamos um DataFrame para Excel usando pandas,
        # a estrutura fica SEMPRE da seguinte forma:
        #
        #   Linha 1 → Cabeçalho (nomes das colunas)
        #   Linha 2 → Primeira linha de dados reais
        #   Linha 3 → Segunda linha de dados reais
        #   ...
        #
        # Por isso, ao aplicar formatações numéricas (moeda, datas,
        # percentagens, etc.) NUNCA devemos incluir a linha 1,
        # porque ela contém texto ("data", "produto", "preco", etc.).
        #
        # Se formatássemos o cabeçalho como moeda ou número,
        # o Excel iria substituir o texto por valores como €0,00.
        # Isso destruiria os nomes das colunas.
        #
        # Assim, usamos SEMPRE:
        #   min_row = 2 → começa a formatação na linha 2 (dados)
        #
        # Também usamos col_x = df.columns.get_loc("nome") + 1
        # para descobrir dinamicamente em que coluna está cada campo,
        # mesmo que a ordem das colunas mude no ficheiro do cliente.
        #
        # Isto torna o programa 100% DINÂMICO e compatível com
        # qualquer CSV que siga as regras de validação.

        # Descobrir dinamicamente em que coluna do Excel está cada campo importante
        coluna_data = df_limpo.columns.get_loc("data") + 1  # perguntamos ao pandas em que posicao esta a coluna 'data' com get_loc
        coluna_produto = df_limpo.columns.get_loc("produto") + 1
        coluna_quantidade = df_limpo.columns.get_loc("quantidade") + 1
        coluna_preco = df_limpo.columns.get_loc("preco") + 1

        for row in ws_dados.iter_rows(min_row=2, min_col=coluna_preco, max_col=coluna_preco):  # iter_rows itera as linhas e celulas dessa folha
            for cell in row:
                cell.number_format = "€#,##0.00"  # por exemplo: 10 → €10,00

        for row in ws_dados.iter_rows(min_row=2, min_col=coluna_data, max_col=coluna_data):
            for cell in row:
                cell.number_format = "DD-MM-YYYY"

        for row in ws_dados.iter_rows(min_row=2, min_col=coluna_quantidade, max_col=coluna_quantidade):
            for cell in row:
                cell.number_format = "0"  # indica ao Excel que os valores devem ser apresentados como números inteiros, sem casas decimais.

        for row in ws_dados.iter_rows(min_row=2, min_col=coluna_produto, max_col=coluna_produto):
            for cell in row:
                cell.alignment = Alignment(horizontal="left")  # A coluna "produto" contém texto, não números.
                                                               # Por isso NÃO aplicamos number_format aqui — caso contrário o Excel
                                                               # iria tentar converter texto em número, destruindo os nomes


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


#def main():
    #caminho = escolher_csv()
    # se o utilizador cancelar a escolha na janela,a def escolher_csv() devolve None e termina o programa sem crashar
    #if caminho is None:
        #return "Nenhum ficheiro foi selecionado"
    # le o ficheiro csv escolhido e transforma-o num DataFrame do pandas
    #df = ler_csv(caminho)
    #if df is None:
        #return "Erro: ficheiro selecionado não é um CSV.Certifique-se de que escolheu um ficheiro válido."
    # limpa os dados: remove linhas vazia, repetidas, etc
    #df_limpo = limpar_dados(df)
    # gera um relatorio com informacoes do DataFrame(numero de linhas, colunas e nomes das colunas)
    #relatorio = gerar_relatorio(df_limpo)
    # exportamos os dados limpos + relatorio para um unico ficheiro excel com duas folhas
    #exportar_excel(df_limpo, relatorio)
    #print("Relatório criado com sucesso!")


# Este bloco so e executado se este ficheiro for executado diretamente. Serve como ponto de entrada principal do programa
#if __name__ == "__main__":
    # Chama a funcao main() para iniciar o processo(gatilho para arrancar o programa)
    #main()

