import pandas as pd
import plotly.graph_objects as go

# Extraindo dados da planilha para DataFrame
entrada_realizado = pd.read_excel(
    open('arquivos/dados.xlsx','rb'),
    sheet_name='realizado',
    header=1,
    index_col=0
    )

entrada_orcado = pd.read_excel(
    open('arquivos/dados.xlsx','rb'),
    sheet_name='orcado'
    )

# Preparando DataFrame de saída
saida_merge = pd.DataFrame(entrada_orcado)

# Salvando valores da sheet 'realizado' em lista
realizado_coluna = []
for label, content in entrada_realizado.items():
    realizado_coluna.append(content[0])

# Adicionando a lista como coluna no DataFrame de saída
saida_merge['realizado'] = realizado_coluna

# Calculando diferença entre valores de 'orcado' e 'realizado'
diferenca_coluna = []
for i in range(saida_merge.shape[0]):
    diferenca_coluna.append(saida_merge['orcado'][i] - saida_merge['realizado'][i])

# Adicionando valores de diferença no DataFrame de saída
saida_merge['diferenca'] = diferenca_coluna

# Exportando DataFrame de saída como CSV
saida_merge.to_csv('saida.csv',index=False)

# Gerando gráfico
lista_meses = []
lista_realizado = []
lista_orcado = []

for i in range(saida_merge.shape[0]):
    lista_meses.append(saida_merge['mês'][i])
    lista_realizado.append(saida_merge['realizado'][i])
    lista_orcado.append(saida_merge['orcado'][i])
    
img = go.Figure(go.Bar(x=lista_meses, y=lista_orcado, name='Orçado'))
img.add_trace(go.Bar(x=lista_meses, y=lista_realizado, name='Realizado', base=0))

img.update_layout(barmode='stack')
img.show()

img.write_image('grafico.png')