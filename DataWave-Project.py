import pandas as pd 
import openpyxl
import numpy as np

############### FECHA ##########
mes = [5]
año=[2023]
############### NEGOCIOS ########
negocio=['SUPERMERCADOS']
negocio2=['MEDICAMENTOS']
#1.1. Cargue & Formato de Bases.
cruce=pd.read_csv('/content/12.TablaTemporalCruces.csv',encoding = "ISO-8859-1",sep=';',on_bad_lines='skip')
cruce=cruce.drop(['TablaParaCruzar_Importancia_Importancia','ClasifEmpresarial'],axis=1)
homo=pd.read_excel('/content/Tabla_Homologacion.xlsx')
homo2=pd.read_excel('/content/Tabla_Homologacion2.xlsx')
homoeducacion=pd.read_excel('/content/Tabla_Homologacion_Educacion.xlsx')
pyg=pd.read_csv('/content/BD_P&G_SAP_2023.csv',encoding = "ISO-8859-1",sep=';',on_bad_lines='skip',header=0,names=None)
pyg=pyg[pyg['MES'].isin(mes)]
pyg=pyg[pyg['NEGOCIO'].isin(negocio)]
pyg2=pyg[pyg['MES'].isin(mes)]
pyg2=pyg[pyg['NEGOCIO'].isin(negocio2)]
#1.2. Cague Bases CET
basecet=pd.read_excel('/content/ABRIL.xlsx',sheet_name='IngOpe -CET',header=0,usecols='P:R',names=['Cuenta Contable','NumeroDocumento','Valor'])
basecet=basecet['Valor']*-1
basecet2=pd.read_excel('/content/ABRIL.xlsx',sheet_name='TD - Fondos -CET',header=2,names=['INFRA','Valor'])
#1.3. Cargue Bases Educación
baseeducacion=pd.read_excel('/content/5. Seguimiento por segmentos Educación y Cultura_2023.xlsx',sheet_name='05.Seguimiento por segmentos E',header=0)
#1.4. Cargue Bases Seguros
baseseguros=pd.read_excel('/content/05.  Retornos seguros Mayo de 2023.xlsx',sheet_name='Mayo 2023',header=0,usecols='E,F,A,B,G',names=None)
baseseguros=baseseguros.assign(TipoAfiliado='INDIVIDUALES')
#1.5 Cargue Bases Supermercados
basesuper1=pd.read_excel('/content/OMARSUPER-052023.xlsx',sheet_name='MAYO_MERCADOS',header=0,names=None)
basesuper1=basesuper1[basesuper1['TIPO DE VENTA']== 'Venta Comercial']
basesuper2=openpyxl.load_workbook('/content/Información cierre para segmentos (2).xlsx')
basesuper2 = basesuper2.get_sheet_by_name('Hoja1')
basesuper3=pd.read_excel('/content/Valores Financieros Finales - Supermercados - May2023.xlsx', header=0,names=None)
#1.6 Cargue Bases Medicamentos
basemed1=pd.read_excel('/content/OMARSUPER-052023.xlsx',sheet_name='MAYO_MERCADOS',header=0,names=None)
basemed1=basemed1[basemed1['TIPO DE VENTA']== 'Venta Comercial']
basemed2=openpyxl.load_workbook('/content/Información cierre para segmentos (2).xlsx')
basemed2 = basemed2.get_sheet_by_name('Hoja1')
basemed3=pd.read_excel('/content/Valores Financieros Finales - Supermercados - May2023.xlsx', header=0,names=None)
#1.7 Cargue Baseaybs Alimentos & Bebidas
baseayb1=pd.read_excel('/content/06.1 Baseayb de datos (facturación) AYB - Junio 2023.xlsx',sheet_name='Opera')
baseayb2=pd.read_excel('/content/06.1 Baseayb de datos (facturación) AYB - Junio 2023.xlsx',sheet_name='Operacional (SAP)')
baseayb3=pd.read_excel('/content/06.1 Baseayb de datos (facturación) AYB - Junio 2023.xlsx',sheet_name='No Operacional (SAP)')
baseayb4=pd.read_excel('/content/Consolidado A&B Junio 2023.xlsm',sheet_name='Homologación Negocio')
baseayb4=baseayb4.drop(['negocio'],axis=1)
baseaybcruzada1 = pd.concat([baseayb1,baseayb2,baseayb3],ignore_index=True)
baseaybcruzada1 = baseaybcruzada1[baseaybcruzada1['SAP'].notna()]
baseaybcruzada1 = baseaybcruzada1.reindex(columns=['Período','SAP','Concepto','CeBe','Negocio','Valor','No. De factura','Nit del cliente','Nombre cliente','Mercado','TIPOAFILIADO','MES','AÑO','Segmento','ORDEN','TipoAfiliado2','ClaseSegmento','VALOR_CONTABLE','negocio homologodo'])
baseaybcruzada1['TIPOAFILIADO'] = baseaybcruzada1['Mercado']
baseaybcruzada1=baseaybcruzada1.drop('Mercado',axis=1)
#1.8 Cargue Bases Recreación & Hotelería 
baseryt1=pd.read_excel('/content/Opera base junio.xlsx',sheet_name='Cubo',usecols='A:L')
baseryt1.columns=['Fe.contab.','Período','fuente','Denominación','CeBe','Nombre negocio','En MLCeBe','Clase','No. De factura','Nit del cliente','Nombre cliente','mercado']
baseryt2=pd.read_excel('/content/Opera base junio.xlsx',sheet_name='Bellavista',usecols='A:L')
baseryt2.columns=['Fe.contab.','Período','fuente','Denominación','CeBe','Nombre negocio','En MLCeBe','Clase','No. De factura','Nit del cliente','Nombre cliente','mercado']
baseryt3=pd.read_excel('/content/Opera base junio.xlsx',sheet_name='Colina',usecols='A:L')
baseryt3.columns=['Fe.contab.','Período','fuente','Denominación','CeBe','Nombre negocio','En MLCeBe','Clase','No. De factura','Nit del cliente','Nombre cliente','mercado']
baseryt4=pd.read_excel('/content/Opera base junio.xlsx',sheet_name='Hotel Alcaravan',usecols='A:L')
baseryt4.columns=['Fe.contab.','Período','fuente','Denominación','CeBe','Nombre negocio','En MLCeBe','Clase','No. De factura','Nit del cliente','Nombre cliente','mercado']
baseryt5=pd.read_excel('/content/Opera base junio.xlsx',sheet_name='Hotel Bosques')
baseryt5.columns=['Fe.contab.','Período','fuente','Denominación','CeBe','Nombre negocio','En MLCeBe','Clase','No. De factura','Nit del cliente','Nombre cliente','mercado']
baseryt6=pd.read_excel('/content/Opera base junio.xlsx',sheet_name='Hotel Peñalisa')
baseryt6.columns=['Fe.contab.','Período','fuente','Denominación','CeBe','Nombre negocio','En MLCeBe','Clase','No. De factura','Nit del cliente','Nombre cliente','mercado']
baseryt7=pd.read_excel('/content/Opera base junio.xlsx',sheet_name='hotel Lanceros',usecols='A:L')
baseryt7.columns=['Fe.contab.','Período','fuente','Denominación','CeBe','Nombre negocio','En MLCeBe','Clase','No. De factura','Nit del cliente','Nombre cliente','mercado']
baseryt8=pd.read_excel('/content/Opera base junio.xlsx',sheet_name='Hotel Colonial',usecols='A:L')
baseryt8.columns=['Fe.contab.','Período','fuente','Denominación','CeBe','Nombre negocio','En MLCeBe','Clase','No. De factura','Nit del cliente','Nombre cliente','mercado']
baseryt9=pd.read_excel('/content/Opera base junio.xlsx',sheet_name='Cantú')
baseryt9.columns=['Fe.contab.','Período','fuente','Denominación','CeBe','Nombre negocio','En MLCeBe','Clase','No. De factura','Nit del cliente','Nombre cliente','mercado']
baseryt10=pd.read_excel('/content/Opera base junio.xlsx',sheet_name='Sección deportes')
baseryt10.columns=['Fe.contab.','Período','fuente','Denominación','CeBe','Nombre negocio','En MLCeBe','Clase','No. De factura','Nit del cliente','Nombre cliente','mercado']
baseryt11=pd.read_excel('/content/Opera base junio.xlsx',sheet_name='bloc')
baseryt11.columns=['Fe.contab.','Período','fuente','Denominación','CeBe','Nombre negocio','En MLCeBe','Clase','No. De factura','Nit del cliente','Nombre cliente','mercado']
baseryt12=pd.read_excel('/content/Opera base junio.xlsx',sheet_name='CDE 195')
baseryt12.columns=['Fe.contab.','Período','fuente','Denominación','CeBe','Nombre negocio','En MLCeBe','Clase','No. De factura','Nit del cliente','Nombre cliente','mercado']
baseryt13=pd.read_excel('/content/Opera base junio.xlsx',sheet_name='Recreación')
baseryt13.columns=['Fe.contab.','Período','fuente','Denominación','CeBe','Nombre negocio','En MLCeBe','Clase','No. De factura','Nit del cliente','Nombre cliente','mercado']
baseryt14=pd.read_excel('/content/Opera base junio.xlsx',sheet_name='Programas turisticos')
baseryt14.columns=['Fe.contab.','Período','fuente','Denominación','CeBe','Nombre negocio','En MLCeBe','Clase','No. De factura','Nit del cliente','Nombre cliente','mercado']
baseryt15=pd.read_excel('/content/Opera base junio.xlsx',sheet_name='eventos y convenciones')
baseryt15.columns=['Fe.contab.','Período','fuente','Denominación','CeBe','Nombre negocio','En MLCeBe','Clase','No. De factura','Nit del cliente','Nombre cliente','mercado']
baseryt16=pd.read_excel('/content/Opera base junio.xlsx',sheet_name='Café Letras')
baseryt16.columns=['Fe.contab.','Período','fuente','Denominación','CeBe','Nombre negocio','En MLCeBe','Clase','No. De factura','Nit del cliente','Nombre cliente','mercado']
baseryt17=pd.read_excel('/content/Opera base junio.xlsx',sheet_name='No operacionales')
baseryt17.columns=['Fe.contab.','Período','fuente','Denominación','CeBe','Nombre negocio','En MLCeBe','Clase','No. De factura','Nit del cliente','Nombre cliente','mercado']
baseryt18=pd.read_excel('/content/Hercules base junio.xlsx',usecols='A:L')
baseryt18.columns=['Fe.contab.','Período','fuente','Denominación','CeBe','Nombre negocio','En MLCeBe','Clase','No. De factura','Nit del cliente','Nombre cliente','mercado']
baseryt19=pd.read_excel('/content/Simphony base junio.xlsx')
baseryt19.columns=['Fe.contab.','Período','fuente','Denominación','CeBe','Nombre negocio','En MLCeBe','Clase','No. De factura','Nit del cliente','Nombre cliente','mercado']
baseryt20=pd.read_excel('/content/Taquilla base junio.xlsx')
baseryt20.columns=['Fe.contab.','Período','fuente','Denominación','CeBe','Nombre negocio','En MLCeBe','Clase','No. De factura','Nit del cliente','Nombre cliente','mercado']
baseryt21=pd.read_excel('/content/Sap base base junio.xlsx',sheet_name='base sap',usecols='A:L')
baseryt21.columns=['Fe.contab.','Período','fuente','Denominación','CeBe','Nombre negocio','En MLCeBe','Clase','No. De factura','Nit del cliente','Nombre cliente','mercado']
baseryt22=pd.read_excel('/content/Sap base base junio.xlsx',sheet_name='No operacionales')
baseryt22.columns=['Fe.contab.','Período','fuente','Denominación','CeBe','Nombre negocio','En MLCeBe','Clase','No. De factura','Nit del cliente','Nombre cliente','mercado']
baseryt23=pd.read_excel('/content/Opera base junio.xlsx',sheet_name='Piscilago',usecols='A:L')
baseryt23.columns=['Fe.contab.','Período','fuente','Denominación','CeBe','Nombre negocio','En MLCeBe','Clase','No. De factura','Nit del cliente','Nombre cliente','mercado']

############### PROCESAMIENTO ##########

#2.1. Procesamiento - CET
#2.2. Cruce: Consolidado + Tabla Cruces = Tabla CET.
basecetcruzada1=basecet.merge(cruce,how='left')
#2.3. Tipificacion Corridas Access Base 1. 
basecetcruzada1.loc[basecetcruzada1['Cuenta Contable'] == '41600509 EMPRESARIAL', "Segmento"] = 'NULL'
basecetcruzada1.fillna('No Afiliado',inplace=True)
basecetcruzada1.drop('NumeroDocumento',axis=1)
basecetcruzada2=basecetcruzada1.assign(TipoAfiliado='INDIVIDUALES')
basecetcruzada2.loc[basecetcruzada1['Segmento'] == 'Empresas Afiliadas', "TipoAfiliado"] = 'EMPRESAS'
basecetcruzada2.loc[basecetcruzada1['Segmento'] == 'NULL', "TipoAfiliado"] = 'EMPRESAS'
#2.4. Tipificacion Access Tabla Homologacion 2 & Base 2
basecet2.loc[basecet2['INFRA'] == '41600511 FOSFEC', "INFRA"] = 'FOSFEC'
basecet2.loc[basecet2['INFRA'] == 'No ope', "INFRA"] = 'NO OPERACIONALES CET'
#2.5. Cruce: Tabla CET + Tabla Homologacion= Tabla Resultado 1 & Consolidado 2 + Tabla Homologacion2 = Tabla Resultado 2.
basecetcruzada3=basecetcruzada2.merge(homo,left_on=['Segmento','TipoAfiliado'], right_on=['Segmento','Tipo_Afiliado'],how='left')
basecetcruzada4=homo2.merge(basecet2,how='left')
#2.6. Agregacion: ACTUALIZAR TABLA FECHA.
basecetcruzada3=basecetcruzada3[basecetcruzada3['MES'].isin(mes)]
basecetcruzada4=basecetcruzada4[basecetcruzada4['MES'].isin(mes)]

#3.1. Procesamiento - Educación
#3.2. Cruce: Consolidado + Tabla Cruces = Tabla EDUCACIÓN.
baseeducacioncruzada1=baseeducacion.merge(cruce,left_on='NUM_DOC_CLI', right_on='NumeroDocumento',how='left')
baseeducacioncruzada1=baseeducacioncruzada1.drop('NumeroDocumento',axis=1) 
baseeducacioncruzada1['Segmento']= baseeducacioncruzada1['Segmento'].fillna('NULL')
#3.3. Cruce: Tabla Base 1+ Tabla Homologacion= Tabla Resultado / Consolidado 2 + Tabla Homologacion2 = Tabla Resultado 2.
baseeducacioncruzada2=baseeducacioncruzada1.merge(homo,left_on=['Segmento','LAE_CLI'], right_on=['Segmento','TipoAfiliado2'],how='left')

#4.1 Procesamiento - Seguros
#4.2. Cruce: Consolidado + Tabla Cruces = Tabla SEGUROS.
baseseguroscruzada1=baseseguros.merge(cruce,left_on='NÚMERO DE DOCUMENTO DEL CLIENTE', right_on='NumeroDocumento',how='left')
baseseguroscruzada1=baseseguroscruzada1.drop('NumeroDocumento',axis=1)
baseseguroscruzada1['Segmento'] = baseseguroscruzada1['Segmento'].fillna('NULL')
#4.3. Cruce: Tabla Base 1+ Tabla Homologacion= Tabla Resultado / Consolidado 2 + Tabla Homologacion2 = Tabla Resultado 2.
baseseguroscruzada2=baseseguroscruzada1.merge(homo,left_on=['Segmento','TipoAfiliado'], right_on=['Segmento','Tipo_Afiliado'],how='left')
#4.4 Valor Contable
valorcontablegeneral=basemedcruzada2['VENTA BRUTA'].sum()
valorcontable=(ventas_brutas_comercial_-descuentopyg)
basemedcruzada2['Valor Contable']=basemedcruzada2['VENTA BRUTA']*valorcontable/valorcontablegeneral

#5.1 Procesamiento - Supermercados
#5.2 Consolidación Valores: PYG
ventasbrutaspyg=pyg[pyg.CUENTA_NIVEL6.isin(['411010', '411005','411030'])]
ventasbrutaspyg=ventasbrutaspyg['ANIO_ACTUAL'].astype(float).sum()
descuentopyg=pyg[pyg['CUENTA_NIVEL6']== 411097]
descuentopyg=descuentopyg['ANIO_ACTUAL'].astype(float).sum()
otrosingoppyg=pyg[pyg['CUENTA_NIVEL6']== 411095]
otrosingoppyg=otrosingoppyg['ANIO_ACTUAL'].astype(float).sum()
ingnooppyg=pyg[pyg['GRUPO']== 42]
ingnooppyg=ingnooppyg['ANIO_ACTUAL'].astype(float).sum()
#5.3 Consolidación Valores: Financieros
ventas_brutas_comercial_= 0
descuento_de_ventas_comercial_=0
ventas_brutas_institucional_=0
descuento_de_ventas_institucional_=0
if (mes[0]== 1):
  ventas_brutas_comercial_= int(basesuper2['b64'].value)*1000000
  descuento_de_ventas_comercial_= int(basesuper2['b65'].value)*1000000
  ventas_brutas_institucional_=int(basesuper2['c64'].value)*1000000
if (mes[0]== 2):
  ventas_brutas_comercial_= int(basesuper2['d64'].value)*1000000
  descuento_de_ventas_comercial_= int(basesuper2['d65'].value)*1000000
  ventas_brutas_institucional_=int(basesuper2['e64'].value)*1000000
if (mes[0]== 3):
  ventas_brutas_comercial_= int(basesuper2['f64'].value)*1000000
  descuento_de_ventas_comercial_= int(basesuper2['f65'].value)*1000000
  ventas_brutas_institucional_=int(basesuper2['g64'].value)*1000000
if (mes[0]== 4):
  ventas_brutas_comercial_= int(basesuper2['h64'].value)*1000000
  descuento_de_ventas_comercial_= int(basesuper2['h65'].value)*1000000
  ventas_brutas_institucional_=int(basesuper2['i64'].value)*1000000
if (mes[0]== 5):
  ventas_brutas_comercial_= int(basesuper2['j64'].value)*1000000
  descuento_de_ventas_comercial_= int(basesuper2['j65'].value)*1000000
  ventas_brutas_institucional_=int(basesuper2['k64'].value)*1000000
if (mes[0]== 6):
  ventas_brutas_comercial_= int(basesuper2['l64'].value)*1000000
  descuento_de_ventas_comercial_= int(basesuper2['l65'].value)*1000000
  ventas_brutas_institucional_=int(basesuper2['m64'].value)*1000000
if (mes[0]== 7):
  ventas_brutas_comercial_= int(basesuper2['n64'].value)*1000000
  descuento_de_ventas_comercial_= int(basesuper2['n65'].value)*1000000
  ventas_brutas_institucional_=int(basesuper2['o64'].value)*1000000
if (mes[0]== 8):
  ventas_brutas_comercial_= int(basesuper2['p64'].value)*1000000
  descuento_de_ventas_comercial_= int(basesuper2['p65'].value)*1000000
  ventas_brutas_institucional_=int(basesuper2['q64'].value)*1000000
if (mes[0]== 9):
  ventas_brutas_comercial_= int(basesuper2['r64'].value)*1000000
  descuento_de_ventas_comercial_= int(basesuper2['r65'].value)*1000000
  ventas_brutas_institucional_=int(basesuper2['s64'].value)*1000000
if (mes[0]== 10):
  ventas_brutas_comercial_= int(basesuper2['t64'].value)*1000000
  descuento_de_ventas_comercial_= int(basesuper2['t65'].value)*1000000
  ventas_brutas_institucional_=int(basesuper2['u64'].value)*1000000
if (mes[0]== 11):
  ventas_brutas_comercial_= int(basesuper2['v64'].value)*1000000
  descuento_de_ventas_comercial_= int(basesuper2['v65'].value)*1000000
  ventas_brutas_institucional_=int(basesuper2['u64'].value)*1000000
if (mes[0]== 12):
  ventas_brutas_comercial_= int(basesuper2['x64'].value)*1000000
  descuento_de_ventas_comercial_= int(basesuper2['x65'].value)*1000000
  ventas_brutas_institucional_=int(basesuper2['y64'].value)*1000000
ventas_netas_comercial_= (ventas_brutas_comercial_-descuento_de_ventas_comercial_)
ventas_netas_institucional_=(ventas_brutas_institucional_-descuento_de_ventas_institucional_)

indices=pd.DataFrame({'Ingresos':['VENTAS_BRUTAS_PYG','DESCUENTO_DE_VENTAS_PYG','VENTAS_NETAS_PYG','OTROS_INGRESOS_NO_OPERACIONALES_PYG','INGRESOS_NO_OPERACIONALES_PYG','TOTAL_INGRESOS_PYG','VENTAS_BRUTAS_COMERCIAL_','DESCUENTO_DE_VENTAS_COMERCIAL_','VENTAS_NETAS_COMERCIAL_','VENTAS_BRUTAS_INSTITUCIONAL_','DESCUENTO_DE_VENTAS_INSTITUCIONAL_','VENTAS_NETAS_INSTITUCIONAL_','CONVENIOS','OTROS_INGRESOS_TOTAL','VALOR_CONTABLE']})
valores=pd.DataFrame({'AÃ‘O_ACTUAL':[ventasbrutaspyg,descuentopyg,(ventasbrutaspyg-descuentopyg),otrosingoppyg,ingnooppyg,(ventasbrutaspyg-descuentopyg+otrosingoppyg+ingnooppyg),ventas_brutas_comercial_, descuento_de_ventas_comercial_, ventas_netas_comercial_, ventas_brutas_institucional_, descuento_de_ventas_institucional_, ventas_netas_institucional_,(ventas_brutas_institucional_+descuento_de_ventas_institucional_),(otrosingoppyg+ingnooppyg),(ventas_brutas_comercial_-descuentopyg)]})
Año=año*15
Mes=mes*15
Fecha=pd.DataFrame({'Año':Año,'Mes':Mes})
resultado=pd.concat([Fecha,indices,valores], axis=1)
basesuper3=basesuper3.append(resultado)
#5.4 Agregación: Fecha & Tipo Afiliado
basesuper1['MES']=basesuper1['MES'].astype(str)
basesuper1['MES']=basesuper1['MES'].str[5:]
basesuper1=basesuper1.assign(TIPOAFILIADO='INDIVIDUALES')
#5.5 Cruce: Consolidado + Tabla Cruces = Tabla SUPERMERCADOS.
basesupercruzada1=basesuper1.merge(cruce,left_on='# DOCUMENTO', right_on='NumeroDocumento',how='left')
basesupercruzada1=basesupercruzada1.drop('NumeroDocumento',axis=1)
basesupercruzada1['Segmento'] = basesupercruzada1['Segmento'].fillna('NULL')
#5.6 Cruce: Tabla Base 1+ Tabla Homologacion= Tabla Resultado / Consolidado 2 + Tabla Homologacion2 = Tabla Resultado 2.
basesupercruzada2=basesupercruzada1.merge(homo,left_on=['Segmento','TIPOAFILIADO'], right_on=['Segmento','Tipo_Afiliado'],how='left')
#5.7 Actaulizar Valor Contable
valorcontablegeneral=basesupercruzada2['VENTA BRUTA'].sum()
valorcontable=(ventas_brutas_comercial_-descuentopyg)
basesupercruzada2['Valor Contable']=basesupercruzada2['VENTA BRUTA']*valorcontable/valorcontablegeneral

#6.1 Procesamiento - Medicamentos
#6.2 Consolidación Valores: pyg2
ventasbrutaspyg2=pyg2[pyg2.CUENTA_NIVEL6.isin(['411010', '411005','411030'])]
ventasbrutaspyg2=ventasbrutaspyg2['ANIO_ACTUAL'].astype(float).sum()
descuentopyg2=pyg2[pyg2['CUENTA_NIVEL6']== 411097]
descuentopyg2=descuentopyg2['ANIO_ACTUAL'].astype(float).sum()
otrosingoppyg2=pyg2[pyg2['CUENTA_NIVEL6']== 411095]
otrosingoppyg2=otrosingoppyg2['ANIO_ACTUAL'].astype(float).sum()
ingnooppyg2=pyg2[pyg2['GRUPO']== 42]
ingnooppyg2=ingnooppyg2['ANIO_ACTUAL'].astype(float).sum()
#6.3 Consolidación Valores: Financieros
ventas_brutas_comercial_= 0
descuento_de_ventas_comercial_=0
ventas_brutas_institucional_=0
descuento_de_ventas_institucional_=0
if (mes[0]== 1):
  ventas_brutas_comercial_= int(basemed2['b64'].value)*1000000
  descuento_de_ventas_comercial_= int(basemed2['b65'].value)*1000000
  ventas_brutas_institucional_=int(basemed2['c64'].value)*1000000
if (mes[0]== 2):
  ventas_brutas_comercial_= int(basemed2['d64'].value)*1000000
  descuento_de_ventas_comercial_= int(basemed2['d65'].value)*1000000
  ventas_brutas_institucional_=int(basemed2['e64'].value)*1000000
if (mes[0]== 3):
  ventas_brutas_comercial_= int(basemed2['f64'].value)*1000000
  descuento_de_ventas_comercial_= int(basemed2['f65'].value)*1000000
  ventas_brutas_institucional_=int(basemed2['g64'].value)*1000000
if (mes[0]== 4):
  ventas_brutas_comercial_= int(basemed2['h64'].value)*1000000
  descuento_de_ventas_comercial_= int(basemed2['h65'].value)*1000000
  ventas_brutas_institucional_=int(basemed2['i64'].value)*1000000
if (mes[0]== 5):
  ventas_brutas_comercial_= int(basemed2['j64'].value)*1000000
  descuento_de_ventas_comercial_= int(basemed2['j65'].value)*1000000
  ventas_brutas_institucional_=int(basemed2['k64'].value)*1000000
if (mes[0]== 6):
  ventas_brutas_comercial_= int(basemed2['l64'].value)*1000000
  descuento_de_ventas_comercial_= int(basemed2['l65'].value)*1000000
  ventas_brutas_institucional_=int(basemed2['m64'].value)*1000000
if (mes[0]== 7):
  ventas_brutas_comercial_= int(basemed2['n64'].value)*1000000
  descuento_de_ventas_comercial_= int(basemed2['n65'].value)*1000000
  ventas_brutas_institucional_=int(basemed2['o64'].value)*1000000
if (mes[0]== 8):
  ventas_brutas_comercial_= int(basemed2['p64'].value)*1000000
  descuento_de_ventas_comercial_= int(basemed2['p65'].value)*1000000
  ventas_brutas_institucional_=int(basemed2['q64'].value)*1000000
if (mes[0]== 9):
  ventas_brutas_comercial_= int(basemed2['r64'].value)*1000000
  descuento_de_ventas_comercial_= int(basemed2['r65'].value)*1000000
  ventas_brutas_institucional_=int(basemed2['s64'].value)*1000000
if (mes[0]== 10):
  ventas_brutas_comercial_= int(basemed2['t64'].value)*1000000
  descuento_de_ventas_comercial_= int(basemed2['t65'].value)*1000000
  ventas_brutas_institucional_=int(basemed2['u64'].value)*1000000
if (mes[0]== 11):
  ventas_brutas_comercial_= int(basemed2['v64'].value)*1000000
  descuento_de_ventas_comercial_= int(basemed2['v65'].value)*1000000
  ventas_brutas_institucional_=int(basemed2['u64'].value)*1000000
if (mes[0]== 12):
  ventas_brutas_comercial_= int(basemed2['x64'].value)*1000000
  descuento_de_ventas_comercial_= int(basemed2['x65'].value)*1000000
  ventas_brutas_institucional_=int(basemed2['y64'].value)*1000000
ventas_netas_comercial_= (ventas_brutas_comercial_-descuento_de_ventas_comercial_)
ventas_netas_institucional_=(ventas_brutas_institucional_-descuento_de_ventas_institucional_)
indices=pd.DataFrame({'Ingresos':['VENTAS_BRUTAS_pyg2','DESCUENTO_DE_VENTAS_pyg2','VENTAS_NETAS_pyg2','OTROS_INGRESOS_NO_OPERACIONALES_pyg2','INGRESOS_NO_OPERACIONALES_pyg2','TOTAL_INGRESOS_pyg2','VENTAS_BRUTAS_COMERCIAL_','DESCUENTO_DE_VENTAS_COMERCIAL_','VENTAS_NETAS_COMERCIAL_','VENTAS_BRUTAS_INSTITUCIONAL_','DESCUENTO_DE_VENTAS_INSTITUCIONAL_','VENTAS_NETAS_INSTITUCIONAL_','CONVENIOS','OTROS_INGRESOS_TOTAL','VALOR_CONTABLE']})
valores=pd.DataFrame({'AÃ‘O_ACTUAL':[ventasbrutaspyg2,descuentopyg2,(ventasbrutaspyg2-descuentopyg2),otrosingoppyg2,ingnooppyg2,(ventasbrutaspyg2-descuentopyg2+otrosingoppyg2+ingnooppyg2),ventas_brutas_comercial_, descuento_de_ventas_comercial_, ventas_netas_comercial_, ventas_brutas_institucional_, descuento_de_ventas_institucional_, ventas_netas_institucional_,(ventas_brutas_institucional_+descuento_de_ventas_institucional_),(otrosingoppyg2+ingnooppyg2),(ventas_brutas_comercial_-descuentopyg2)]})
Año=año*15
Mes=mes*15
Fecha=pd.DataFrame({'Año':Año,'Mes':Mes})
resultado=pd.concat([Fecha,indices,valores], axis=1)
basemed3=basemed3.append(resultado)
#6.4 Agregación: Fecha & Tipo Afiliado
basemed1['MES']=basemed1['MES'].astype(str)
basemed1['MES']=basemed1['MES'].str[5:]
basemed1=basemed1.assign(TIPOAFILIADO='INDIVIDUALES')
#6.5 Cruce: Consolidado + Tabla Cruces = Tabla MEDICAMENTOS.
basemedcruzada1=basemed1.merge(cruce,left_on='# DOCUMENTO', right_on='NumeroDocumento',how='left')
basemedcruzada1=basemedcruzada1.drop('NumeroDocumento',axis=1)
basemedcruzada1['Segmento'] = basemedcruzada1['Segmento'].fillna('NULL')
#6.6 Cruce: Tabla Base 1+ Tabla Homologacion= Tabla Resultado / Consolidado 2 + Tabla Homologacion2 = Tabla Resultado 2.
basemedcruzada2=basemedcruzada1.merge(homo,left_on=['Segmento','TIPOAFILIADO'], right_on=['Segmento','Tipo_Afiliado'],how='left')
#6.7 Actualización Valor Contable
valorcontablegeneral=basemedcruzada2['VENTA BRUTA'].sum()
valorcontable=(ventas_brutas_comercial_-descuentopyg)
basemedcruzada2['Valor Contable']=basemedcruzada2['VENTA BRUTA']*valorcontable/valorcontablegeneral

#7.1 Procesamiento - Alimentos & Bebibdas
#7.2 Tipificación: Tipo Afiliado
baseaybcruzada1.loc[baseaybcruzada1['TIPOAFILIADO'] == 'Público', "TIPOAFILIADO"] = 'INDIVIDUALES'
baseaybcruzada1.loc[baseaybcruzada1['TIPOAFILIADO'] == 'Empresarial', "TIPOAFILIADO"] = 'EMPRESAS'
baseaybcruzada1.loc[baseaybcruzada1['TIPOAFILIADO'] == 'Proyectos', "TIPOAFILIADO"] = 'ESFUERZO INSTITUCIONAL'
baseaybcruzada1.loc[baseaybcruzada1['TIPOAFILIADO'] == 'Otros', "TIPOAFILIADO"] = 'ESFUERZO INSTITUCIONAL'
baseaybcruzada1.loc[baseaybcruzada1['TIPOAFILIADO'] == 'Institucional', "TIPOAFILIADO"] = 'ESFUERZO INSTITUCIONAL'
#7.3 Aregación: FECHA
for row in baseaybcruzada1.loc[baseaybcruzada1.MES.isnull(), 'MES'].index:
    baseaybcruzada1.at[row, 'MES'] = mes
for row in baseaybcruzada1.loc[baseaybcruzada1.AÑO.isnull(), 'AÑO'].index:
    baseaybcruzada1.at[row, 'AÑO'] = año
#7.4 Homologación & Reemplazo de Negocio
baseaybcruzada2=baseaybcruzada1.merge(baseayb4,left_on='CeBe', right_on='CeBE',how='left')
baseaybcruzada2=baseaybcruzada2.drop(['CeBE'],axis=1)
baseaybcruzada2=baseaybcruzada2.drop(['negocio homologodo'],axis=1)
baseaybcruzada2['Negocio']=baseaybcruzada2['Homologación']
#7.5 Consolidación Tabla Financiera
tablafinanciera=pd.DataFrame(baseaybcruzada2.pivot_table('Valor','Negocio',aggfunc=np.sum))
tablafinanciera=tablafinanciera.reindex(columns=['Valor','MES','AÑO'])
for row in tablafinanciera.loc[tablafinanciera.MES.isnull(), 'MES'].index:
    tablafinanciera.at[row, 'MES'] = mes
for row in tablafinanciera.loc[tablafinanciera.AÑO.isnull(), 'AÑO'].index:
    tablafinanciera.at[row, 'AÑO'] = año

#8.1 Procesamiento - Recreación & Hotelería
#8.2 Consolidación de las bases
baseryt24=pd.read_excel('/content/Consolidado R&D Junio 2023.xlsx',sheet_name='Homologacion Negocio')
baserytcruzada1 = pd.concat([baseryt1,baseryt2,baseryt3,baseryt4,baseryt5,baseryt6,baseryt7,baseryt8,baseryt9,baseryt10,baseryt11,baseryt12,baseryt13,baseryt14,baseryt15,baseryt16,baseryt17,baseryt18,baseryt19,baseryt20,baseryt21,baseryt22,baseryt23])
baserytcruzada1 = baserytcruzada1.reindex(columns=['Fe.contab.','Período','fuente','Denominación','CeBe','Nombre negocio','En MLCeBe','Clase','No. De factura','Nit del cliente','Nombre cliente','mercado','TIPOAFILIADO','MES','AÑO','Segmento','ORDEN','TipoAfiliado2','ClaseSegmento','VALOR_CONTABLE','negocio homologacion','gerencia'])
baserytcruzada1 = baserytcruzada1[baserytcruzada1['Fe.contab.'].notna()]
#8.3 Tipificación Tipo Afiliado
baserytcruzada1.loc[baserytcruzada1['TIPOAFILIADO'] == 'Individual', "TIPOAFILIADO"] = 'INDIVIDUALES'
baserytcruzada1.loc[baserytcruzada1['TIPOAFILIADO'] == 'Empresarial', "TIPOAFILIADO"] = 'EMPRESAS'
baserytcruzada1.loc[baserytcruzada1['TIPOAFILIADO'] == 'Esfuerzo Institucional', "TIPOAFILIADO"] = 'ESFUERZO INSTITUCIONAL'
for row in baserytcruzada1.loc[baserytcruzada1.MES.isnull(), 'MES'].index:
    baserytcruzada1.at[row, 'MES'] = mes
for row in baserytcruzada1.loc[baserytcruzada1.AÑO.isnull(), 'AÑO'].index:
    baserytcruzada1.at[row, 'AÑO'] = año
#8.4 Homologación & Reemplazo de Negocio
baserytcruzada2=baserytcruzada1.merge(baseryt24,left_on='negocio', right_on='negocio',how='left')
baserytcruzada2=baserytcruzada2.drop(['gerencia', 'negocio homologacion'],axis=1)
baserytcruzada2['negocio']=baserytcruzada2['Homologación']
baserytcruzada2.head(20)
#8.5 Consolidación Tabla Financiera
tablafinanciera2=pd.DataFrame(baserytcruzada2.pivot_table('valor','negocio',aggfunc=np.sum))
baserytcruzada1.head(20)


#10. Exportacion Archivos

