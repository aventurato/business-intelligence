# ----- LIMPEZA (RETIRANDO COLUNAS) ----- #
df_copa.drop(['Nº documento refer.', 'Ordem do cliente', 'Item ord.cliente', 'Exercício', 'Equipe de vendas', 'Escritório de vendas', 
              'Denominação.1', 'Local expedição/recebimento', 'Código da cidade', 'Grupo de preço', 'BANDEIRA', 'Unid.medida básica', 
              'Taxa de imposto', 'Base do imposto', 'ICMS Compra', 'Dest.', 'Valor condição', 'ICM3', 'ICS3', 'PIS', 'COFINS', 'Complemento', 
              'Ressarcimento', 'Desc.PIS/COF', 'Rua', 'Local', 'Bairro'], axis=1, inplace=True)


# ----- RENOMEANDO COLUNAS ----- #
df_copa.rename(columns = {'Nota Fiscal':'NF-e', 'Centro':'Filial', 'Data de lançamento':'Data', 'Artigo':'Cód. Mat.', 'Material básico':'Descr. Mat.', 'Denominação':'Vendedor', 'Tp.doc.faturamento':'Fat.', 
                               'Tipo doc.vendas':'Fat. 1', 'Incoterms':'Transporte', 'Condições pgto.':'Condição', 'Região':'UF', 'Denominação.2':'Tipo', 'Qtd.vendas':'Qtd. M³', 'Custo Final de Venda':'Custo Final', 
                               'Custo Produto':'Custo Tabela', 'Acresc/Redução Fisca':'Fiscal', 'Custo Financeiro':'Financeiro', 'Custo Frete':'Frete', 'Margem Comercial':'Comercial', 'Custo de Partida':'Custo Partida', 
                               'Quantidade-UMR':'Qtd. 20°', 'Custo Mix Compra':'MIX', 'Preço interno':'Preço Interno', 'Margem real':'Margem (R$)', 'ID parceiro':'Cód. Cliente', 'Nome':'Razão Social', 'Região.1':'Estado', 
                               'Código postal':'CEP', 'Unnamed: 56':'Total NF-e', 'Unnamed: 57':'Margem Interna (R$)', 'Unnamed: 58':'Margem Interna Unit.'}, inplace=True)


# ----- CRIANDO COLUNAS ----- #
df_copa['Qtd. Litros'] = (df_copa['Qtd. M³'] * 1000) #conversão m³ para litros.
df_copa['Unit. Frete'] = (df_copa['Frete'] / df_copa['Qtd. Litros']) #calculando valor unitário custo frete.
df_copa['Unit. Financ.'] = (df_copa['Financeiro'] / df_copa['Qtd. Litros']) #calculando valor unitário custo financeiro.
df_copa['Valor MIX'] = (df_copa['MIX'] / df_copa['Qtd. Litros']) #conversão de unidade.
df_copa['Custo Total'] = df_copa['Custo Tabela'] + df_copa['Fiscal']
df_copa['Tabela C'] = (df_copa['Custo Total'] / df_copa['Qtd. Litros']) #conversão de unidade.
df_copa['Tarifa'] = (df_copa['Custo Final'] - df_copa['Custo Tabela']) / (df_copa['Qtd. Litros']) #conversão de unidade.
df_copa['Preço Partida'] = (df_copa['Tabela C'] + df_copa['Tarifa']) #soma da tabela C + tabela de venda.
df_copa['Margem Comercial'] = (df_copa['Comercial'] / df_copa['Qtd. Litros']) #conversão de unidade.
df_copa['Margem MIX'] = (df_copa['Tabela C'] - df_copa['Valor MIX']) #calculando margem mix.
df_copa['Margem Total'] = (df_copa['Margem MIX'] + df_copa['Margem Comercial'] + df_copa['Tarifa']) #calculando margem total.
df_copa['Margem Total (R$)'] = (df_copa['Margem Total'] * df_copa['Qtd. Litros']) #conversão de unidade.


# ----- GERANDO ARQUIVO TRATADO ----- #
df_copa.to_excel('Base Margem.xlsx')