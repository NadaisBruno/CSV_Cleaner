# CSV Cleaner and Report Generator

Aplicaçao desenvolvida em Python para validar, limpar e gerar relatorios a partir de ficheiros CSV com uma ‘interface’ web simples criada com Streamlit.
Este projeto tem como objetivo automatizar tarefas repetitivas de limpeza de dados.

## Funcionalidades

 - Upload de ficheiros CSV atraves do navegador(Streamlit)
 - Validação automática da estrutura do CSV(colunas obrigatórias)
 - Identificação de erros nos dados (datas no formato inválido, valores não numéricos, carateres inválidos, valores negativos, etc)
 - Limpeza de dados com a remoção de duplicados, remoção de espaços em todas as colunas, ordenação de datas
 - Visualização de dados brutos e dos dados limpos
 - Geração de relatório com metricas importantes(médias, produto mais vendido, produto mais caro, top 3 produtos mais vendidos, etc)
 - Exportação automatica de um ficheiro Excel formatado

## Como usar

1 - Executar a aplicação com o comando streamlit run app.py

2 - Seguir as instruções no navegador

## Requisitos do Ficheiro CSV para permitir a validação correta dos dados

O ficheiro deve conter as seguintes colunas obrigatórias:
 - data
 - produto
 - quantidade
 - preco

## Tecnologias Utilizadas
    - Python >= 3.10
    - OpenPyXL
    - Pandas
    - Streamlit

