# ----- CARREGANDO DADOS ----- #
dados_vendas = pd.read_excel('Vendas Copa.XLSX')


# ----- LIMPEZA (RETIRANDO COLUNAS) ----- #
dados_vendas.drop(['Nº documento refer.', 'Ordem do cliente', 'Item ord.cliente', 'Exercício', 'Tipo doc.vendas', 'Equipe de vendas',
                  'Código da cidade', 'Grupo de preço', 'BANDEIRA', 'Unid.medida básica', 'Condições pgto.', 'Taxa de imposto', 'Base do imposto', 
                  'ICMS Compra', 'Dest.', 'Valor condição', 'Preço interno', 'ICM3', 'ICS3', 'PIS', 'COFINS', 'Complemento', 
                  'Ressarcimento', 'Desc.PIS/COF'], axis=1, inplace=True)


# ----- RENOMEANDO COLUNAS ----- #
dados_vendas.rename(columns = {'Nota Fiscal': 'NF-E', 'Centro': 'FILIAL', 'Data de lançamento': 'DATA', 
                               'Artigo': 'COD. MATERIAL', 'Material básico': 'MATERIAL', 'Denominação': 'VENDEDOR', 
                               'Tp.doc.faturamento': 'TIPO DOC.', 'Nome 1': 'NOME FILIAL', 'Incoterms': 'TIPO FRETE', 
                               'Escritório de vendas': 'ESCRITÓRIO', 'Denominação.1': 'DESCR. ESCRITÓRIO', 'Local expedição/recebimento': 'LOCAL EXPEDIÇÃO', 
                               'Região': 'UF', 'Denominação.2': 'TIPO', 'Qtd.vendas': 'VOLUME 20º', 
                               'Custo Final de Venda': 'CUSTO FINAL', 'Custo Produto': 'CUSTO PRODUTO', 'Acresc/Redução Fisca': 'FISCAL', 
                               'Custo Financeiro': 'FINANCEIRO', 'Custo Frete': 'FRETE', 'Margem Comercial': 'MARGEM COMERCIAL', 
                               'Custo de Partida': 'CUSTO PARTIDA', 'Quantidade-UMR': 'VOLUME AMB.', 'Custo Mix Compra': 'CUSTO MIX', 
                               'Margem real': 'MARGEM REAL', 'ID parceiro': 'COD. CLIENTE', 'Nome': 'NOME CLIENTE', 
                               'Rua': 'RUA', 'Local': 'MUNICÍPIO', 'Bairro': 'BAIRRO', 
                               'Região.1': 'ESTADO', 'Código postal': 'CEP', 'Unnamed: 56': 'TOTAL NF-E', 
                               'Unnamed: 57': 'MARGEM TOTAL', 'Unnamed: 58': 'MARGEM POR LITRO', 'Cidade Geo': 'CIDADE GEO', 
                               'Mesorregião': 'MESORREGIÃO', 'Microrregião': 'MICRORREGIÃO', 'Latitude': 'LATITUDE', 
                               'Longitude': 'LONGITUDE', 'Cidade': 'CIDADE'}, inplace=True)


# ----- CRIANDO COLUNAS (DIA | MÊS | ANO) ----- #
dados_vendas['DIA'] = pd.DatetimeIndex(dados_vendas['DATA']).day
dados_vendas['MÊS'] = pd.DatetimeIndex(dados_vendas['DATA']).month
dados_vendas['ANO'] = pd.DatetimeIndex(dados_vendas['DATA']).year
dados_vendas['ANO/MÊS'] = dados_vendas['DATA'].dt.to_period('M').astype(str)
dados_vendas['SEMANA'] = dados_vendas['DATA'].dt.to_period('W').astype(str)


# ----- VERIFICANDO - (QTD. LINHAS | COLUNAS) ----- #
print('{} Linhas - {} Colunas'.format(dados_vendas.shape[0], dados_vendas.shape[1]))


# ----- VERIFICANDO - DADOS NULOS ----- #
print('{} Dados sem valor'.format(dados_vendas[dados_vendas['COD. CLIENTE'].isnull()].shape[0]))


# ----- VERIFICANDO - (DATA INICIAL | DATA FINAL) ----- #
print('Datas de {} até {}'.format(dados_vendas['DATA'].min(), dados_vendas['DATA'].max()))


# ----- QTD. DIAS - (DATA INICIAL | DATA FINAL) ----- #
print(dados_vendas['DATA'].max() - dados_vendas['DATA'].min())


# ----- AGRUPAMENTO DADOS - (VOLUME | TOTAL NF-E | MARGEM TOTAL) ----- #
dados_agrupados = dados_vendas.groupby(['DATA', 'COD. CLIENTE', 'ESTADO', 'CIDADE', 
                                       'MESORREGIÃO', 'MICRORREGIÃO', 'LATITUDE', 'LONGITUDE']).agg({'VOLUME 20º': lambda x: x.sum(), 
                                                                                                     'TOTAL NF-E': lambda x: x.sum(), 
                                                                                                     'MARGEM TOTAL': lambda x: x.sum()}).reset_index()



### ----- DEFININDO (DATA ATUAL). SERÁ UTILIZADA COMO REFERÊNCIA PARA CALCULAR A PONTUAÇÃO DA RECÊNCIA ----- ###

# ----- VERIFICANDO DATA ATUAL ----- #
data_atual = dados_agrupados['DATA'].max() + timedelta(days=1)


# ----- PERIODO (QTD. DIAS - (DATA INICIAL | DATA FINAL)) ----- #
qtd_dias = (dados_agrupados['DATA'].max() - dados_agrupados['DATA'].min()) / np.timedelta64(1, 'D')



# ----- CALCULANDO RECÊNCIA | FREQUÊNCIA | MONETÁRIO PARA CADA CLIENTE ----- #

# ----- CALCULANDO QTD. DIAS TOTAL (INICIAL - FINAL) ----- #
dados_agrupados['Nº DIAS'] = dados_agrupados['DATA'].apply(lambda x: (data_atual - x).days)


# ----- CALCULANDO QTD. DIAS TOTAL (DATA ÚLTIMA COMPRA) ----- #
dados_data = {'Nº DIAS': lambda x: x.min(), 'DATA': lambda x: len([d for d in x if d >= data_atual - timedelta(days=qtd_dias)])}  


# ----- GERANDO NOVA DATAFRAME (RFM) ----- #
rfm = dados_agrupados.groupby(['COD. CLIENTE', 'ESTADO', 'CIDADE', 'MESORREGIÃO', 'MICRORREGIÃO', 'LATITUDE', 'LONGITUDE']).agg(dados_data).reset_index()
rfm.rename(columns={'Nº DIAS': 'RECÊNCIA', 'DATA': 'FREQUÊNCIA'}, inplace=True)


# ----- CALCULANDO VOLUME MÉDIO POR CLIENTE ----- #
rfm['MONTANTE'] = rfm['COD. CLIENTE'].apply(lambda x: dados_agrupados[(dados_agrupados['COD. CLIENTE'] == x) & \
                                                           (dados_agrupados['DATA'] >= data_atual - timedelta(days=qtd_dias))]\
                                                           ['VOLUME 20º'].sum())


# ----- CALCULANDO QTD. PRODUTOS POR CLIENTE ----- #
dados_mix = pd.crosstab(dados_vendas['COD. CLIENTE'], dados_vendas['MATERIAL'], dados_vendas['COD. MATERIAL'], aggfunc='count', margins=False, margins_name='Média Geral')
dados_mix = pd.DataFrame(dados_mix)
dados_mix.fillna(0)


# ----- UNINDO OS DATAFRAMES (DADOS_VENDA + RFM) ----- #
dados_final = pd.merge(left= rfm, right= dados_mix, on='COD. CLIENTE', how='outer')


# ----- RETIRANDO TODOS OS VALORES NULOS ----- #
dados_final['ANIDRO'] = dados_final['ANIDRO'].fillna(0)
dados_final['B100'] = dados_final['B100'].fillna(0)
dados_final['DIESEL S10'] = dados_final['DIESEL S10'].fillna(0)
dados_final['DIESEL S500'] = dados_final['DIESEL S500'].fillna(0)
dados_final['ETANOL'] = dados_final['ETANOL'].fillna(0)
dados_final['GASOLINA'] = dados_final['GASOLINA'].fillna(0)
dados_final['GASOLINA A'] = dados_final['GASOLINA A'].fillna(0)
dados_final['S10 A'] = dados_final['S10 A'].fillna(0)
dados_final['S500 A'] = dados_final['S500 A'].fillna(0)


# ----- LIMPEZA (RETIRANDO COLUNAS) ----- #
dados_final.drop(['ANIDRO', 'B100', 'GASOLINA A', 'S10 A', 'S500 A'], axis=1, inplace=True)


# ----- DIVISÃO EM QUINTIL (CADA QUINTIL CONTÉM 20% DA POPULAÇÃO) ----- #
divisao = dados_final[['RECÊNCIA', 'FREQUÊNCIA', 'MONTANTE', 'DIESEL S10', 'DIESEL S500', 'ETANOL', 'GASOLINA']].quantile([.2, .4, .6, .8]).to_dict()



### ----- CLASSIFICAÇÃO (RECÊNCIA | FREQUÊNCIA | MONTANTE | DIESEL S500 | DIESEL S10 | GASOLINA | ETANOL) ----- ###

# ----- DISTRIBUIÇÃO RECÊNCIA ----- #
def r_score(x):
  if x <= divisao['RECÊNCIA'][.2]:
    return 5
  elif x <= divisao['RECÊNCIA'][.4]:
    return 4  
  elif x <= divisao['RECÊNCIA'][.6]:
    return 3
  elif x <= divisao['RECÊNCIA'][.8]:
    return 2
  else:
    return 1

# ----- DISTRIBUIÇÃO FREQUÊNCIA | MONTANTE ----- #
def fm_score(x, c):
  if x <= divisao[c][.2]:
    return 1
  elif x <= divisao[c][.4]:
    return 2
  elif x <= divisao[c][.6]:
    return 3
  elif x <= divisao[c][.8]:
    return 4
  else:
    return 5     

# ----- DISTRIBUIÇÃO DIESEL S10 ----- #
def S10(x):
  if x <= divisao['DIESEL S10'][.2]:
    return 1
  elif x <= divisao['DIESEL S10'][.4]:
    return 2  
  elif x <= divisao['DIESEL S10'][.6]:
    return 3
  elif x <= divisao['DIESEL S10'][.8]:
    return 4
  else:
    return 5

# ----- DISTRIBUIÇÃO DIESEL S500 ----- #
def S500(x):
  if x <= divisao['DIESEL S500'][.2]:
    return 1
  elif x <= divisao['DIESEL S500'][.4]:
    return 2  
  elif x <= divisao['DIESEL S500'][.6]:
    return 3
  elif x <= divisao['DIESEL S500'][.8]:
    return 4
  else:
    return 5

# ----- DISTRIBUIÇÃO GASOLINA ----- #
def GMIX(x):
  if x <= divisao['GASOLINA'][.2]:
    return 1
  elif x <= divisao['GASOLINA'][.4]:
    return 2  
  elif x <= divisao['GASOLINA'][.6]:
    return 3
  elif x <= divisao['GASOLINA'][.8]:
    return 4
  else:
    return 5

# ----- DISTRIBUIÇÃO ETANOL ----- #
def AEHC(x):
  if x <= divisao['ETANOL'][.2]:
    return 1
  elif x <= divisao['ETANOL'][.4]:
    return 2  
  elif x <= divisao['ETANOL'][.6]:
    return 3
  elif x <= divisao['ETANOL'][.8]:
    return 4
  else:
    return 5



### ----- APLICANDO CLASSIFICAÇÃO ----- ###

# ----- RECÊNCIA ----- #
dados_final['R'] = dados_final['RECÊNCIA'].apply(lambda x: r_score(x))


# ----- FREQUÊNCIA ----- #
dados_final['F'] = dados_final['FREQUÊNCIA'].apply(lambda x: fm_score(x, 'FREQUÊNCIA'))


# ----- MONTANTE ----- #
dados_final['M'] = dados_final['MONTANTE'].apply(lambda x: fm_score(x, 'MONTANTE'))



### ----- GERAÇÃO SCORE RFM ----- ###

# ----- SCORE FV ----- #
dados_final['FV'] = round((dados_final['F'] + dados_final['M']) / 2, 0)

# ----- SCORE RFV ----- #
dados_final['RFV'] = dados_final['R'].map(str) + dados_final['FV'].map(str)

# ----- SCORE RFM ----- #
dados_final['RFM NOTA'] = dados_final['R'].map(str) + dados_final['F'].map(str) + dados_final['M'].map(str)



### ----- CLASSIFICAÇÃO SEGMENTAÇÃO ----- ###

# ----- GERAÇÃO SCORE RFM ----- #
mapa_segmentacao = {r'15.0': 'Não Podemos Perder', r'55.0': 'Campeão', r'33.0': 'Precisa de Atenção', r'41.0': 'Promissor', r'51.0': 'Novo',
                    r'22.0': 'Hibernando', r'52.0': 'Potencialmente Fiel', r'53.0': 'Potencialmente Fiel', r'42.0': 'Potencialmente Fiel',
                    r'43.0': 'Potencialmente Fiel', r'13.0': 'Em Risco', r'14.0': 'Em Risco', r'23.0': 'Em Risco', r'24.0': 'Em Risco',
                    r'25.0': 'Em Risco', r'34.0': 'Cliente Fiel', r'35.0': 'Cliente Fiel', r'44.0': 'Cliente Fiel', r'45.0': 'Cliente Fiel',
                    r'54.0': 'Cliente Fiel', r'11.0': 'Perdido', r'12.0': 'Perdido', r'21.0': 'Perdido', r'31.0': 'Prestes a Hibernar',
                    r'32.0': 'Prestes a Hibernar'}

dados_final['SEGMENTAÇÃO'] = dados_final['R'].map(str) + dados_final['FV'].map(str) 
dados_final['SEGMENTAÇÃO'] = dados_final['SEGMENTAÇÃO'].replace(mapa_segmentacao, regex=True)


# ----- PRÉVIA ARQUIVO FINAL ----- #
dados_final.head(20)


# ----- GERANDO ARQUIVO EM EXCEL ----- #
dados_final.to_excel('Tabela RFM.xlsx')