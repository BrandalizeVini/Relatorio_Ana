from hdbcli import dbapi
import time
from pyhdbcli import Error
from openpyxl import Workbook
from decimal import Decimal, ROUND_DOWN
import os
from datetime import datetime, timedelta

today = datetime.now()
first_day_of_month = datetime(today.year, today.month, 1)
yesterday = today - timedelta(days=1)
first_day_of_year = datetime(today.year, 1, 1)
yesterday_str = yesterday.strftime('%Y%m%d')
year = today.year
first_day_of_year = first_day_of_year.strftime('%Y%m%d')
first_day_of_month_str = first_day_of_month.strftime('%Y%m%d')
first_day_of_last_year = datetime(today.year - 1, 1, 1)
first_day_of_last_year_str = first_day_of_last_year.strftime('%Y%m%d')
if today.month == 12:
    first_day_of_next_month = datetime(today.year + 1, 1, 1)
else:
    first_day_of_next_month = datetime(today.year, today.month + 1, 1)
last_day_of_current_month = first_day_of_next_month - timedelta(days=1)
last_day_of_current_month_str_ymd = last_day_of_current_month.strftime('%Y%m%d')


#first_day_of_month_str = '20231101'
#yesterday_str = '20231130'
#sublinha = 'MERCADORIAS DE REVENDA'

try:
   con = dbapi.connect(
      address="192.168.1.177",
      port=30415,
      user="SYSTEM",
      password="Tirol#1974",
      databasename='HAP'
   )

   '''declaracao = f"""
      SELECT "DT_BUDAT", "DT_ANO", "DT_PERIO","DscRepresentante","VTEXT_PAPH2", "DSC_COMPOSTO", "ARTNR", "MAKTX","PRCTR","LTEXT_PRCTR",sum("KF_QTDADE_VENDAS") AS "KF_QTDADE_VENDAS",sum("KF_PESO_LIQUIDO") AS KF_PESO_LIQUIDO, sum("KF_SOMA_RECEITA_BRUTA_VENDAS") AS KF_SOMA_RECEITA_BRUTA_VENDAS,sum("KF_DEDUCOES_VENDAS") AS KF_DEDUCOES_VENDAS,sum("KF_DEVOLUCAO_VENDAS") AS KF_DEVOLUCAO_VENDAS, sum("KF_DESCONTOS_INSTITUC") AS KF_DESCONTOS_INSTITUC, sum("KF_OVER_PRICING") AS KF_OVER_PRICING, sum("KF_IMPOSTOS_S_VENDAS") AS KF_IMPOSTOS_S_VENDAS,sum("KF_ICMS") AS KF_ICMS,sum("KF_PIS") AS KF_PIS,sum("KF_COFINS") AS KF_COFINS,sum("KF_REC_LIQ_VENDAS") AS KF_REC_LIQ_VENDAS,sum("KF_CUSTO_PROD_VENDIDO") AS KF_CUSTO_PROD_VENDIDO,sum("KF_MATERIAIS") AS KF_MATERIAIS,sum("KF_MATERIA_PRIMA") AS KF_MATERIA_PRIMA,sum("KF_INGREDIENTES") AS KF_INGREDIENTES,sum("KF_EMBALAGENS") AS KF_EMBALAGENS,sum("KF_SEMIACABADOS") AS KF_SEMIACABADOS,sum("KF_PRODUTO_ACABADO") AS KF_PRODUTO_ACABADO,sum("KF_DIFER_INVENTARIO") AS KF_DIFER_INVENTARIO,sum("KF_SERVICOS") AS KF_SERVICOS,sum("KF_SUBCONTRATACAO") AS KF_SUBCONTRATACAO,sum("KF_CMV_REVENDA") AS KF_CMV_REVENDA,sum("KF_RESULT_BRUTO_OPER") AS KF_RESULT_BRUTO_OPER,sum("KF_DESP_COM_VARIAVEIS") AS KF_DESP_COM_VARIAVEIS,sum("KF_COMISSOES_VENDAS") AS KF_COMISSOES_VENDAS,sum("KF_DESP_LOGISTICA_SOMA") AS KF_DESP_LOGISTICA_SOMA,sum("KF_MARGEM_DE_CONTRIB") AS KF_MARGEM_DE_CONTRIB,sum("KF_CUSTOS_FIXOS") AS KF_CUSTOS_FIXOS,sum("KF_MAO_DE_OBRA") AS KF_MAO_DE_OBRA,sum("KF_GASTOS_GERAIS_FAB") AS KF_GASTOS_GERAIS_FAB,sum("KF_DESPESAS_FIXAS") AS KF_DESPESAS_FIXAS,sum("KF_DESP_ADMINISTRATIV") AS KF_DESP_ADMINISTRATIV,sum("KF_DESP_COMERCIAIS_TRADE") AS KF_DESP_COMERCIAIS_TRADE,sum("KF_DESP_COMERCIAIS") AS KF_DESP_COMERCIAIS,sum("KF_DESP_TRADE_MKT") AS KF_DESP_TRADE_MKT,sum("KF_DESP_MKT") AS KF_DESP_MKT,sum("KF_DESP_OPERACIONAL") AS KF_DESP_OPERACIONAL,sum("KF_DESP_TRIBUTARIA") AS KF_DESP_TRIBUTARIA,sum("KF_DESPESAS_PERDA_E_PROV") AS KF_DESPESAS_PERDA_E_PROV,sum("KF_OUTROS_RESULT_OPER") AS KF_OUTROS_RESULT_OPER,sum("KF_OUTRAS_REC_OPER") AS KF_OUTRAS_REC_OPER,sum("KF_RES_ANTES_EF_FINAN") AS KF_RES_ANTES_EF_FINAN,sum("KF_RESULTADO_FINANC") AS KF_RESULTADO_FINANC,sum("KF_DESPESA_FINANCEIRA") AS KF_DESPESA_FINANCEIRA,sum("KF_RES_ANTES_IR_CSLL") AS KF_RES_ANTES_IR_CSLL,sum("KF_RES_APOS_IR_CSLL") AS KF_RES_APOS_IR_CSLl
      FROM "_SYS_BIC"."tirol.co/LUM_DRE_COPA_REAL_ROBO" 
      WHERE (("MANDT" IN ('300') )) 
      AND (("DT_BUDAT" BETWEEN ('{first_day_of_last_year_str}')
      AND ('{last_day_of_current_month_str_ymd}'))) 
      GROUP BY "MANDT", "DT_PERIO", "DT_ANO", "VRGAR", "VERSI", "DT_BUDAT", "ARTNR", "MAKTX", "VTEXT_PAPH2", "DSC_COMPOSTO", "PRCTR", "LTEXT_PRCTR","DscRepresentante"
      """'''
      
   declaracao = f"""
      SELECT "DT_BUDAT", "DT_ANO", "DT_PERIO","DscRepresentante","VTEXT_PAPH2", "DSC_COMPOSTO", "ARTNR", "MAKTX","PRCTR","LTEXT_PRCTR",sum("KF_QTDADE_VENDAS") AS "KF_QTDADE_VENDAS",sum("KF_PESO_LIQUIDO") AS KF_PESO_LIQUIDO, sum("KF_SOMA_RECEITA_BRUTA_VENDAS") AS KF_SOMA_RECEITA_BRUTA_VENDAS,sum("KF_DEDUCOES_VENDAS") AS KF_DEDUCOES_VENDAS,sum("KF_DEVOLUCAO_VENDAS") AS KF_DEVOLUCAO_VENDAS, sum("KF_DESCONTOS_INSTITUC") AS KF_DESCONTOS_INSTITUC, sum("KF_OVER_PRICING") AS KF_OVER_PRICING, sum("KF_IMPOSTOS_S_VENDAS") AS KF_IMPOSTOS_S_VENDAS,sum("KF_ICMS") AS KF_ICMS,sum("KF_PIS") AS KF_PIS,sum("KF_COFINS") AS KF_COFINS,sum("KF_REC_LIQ_VENDAS") AS KF_REC_LIQ_VENDAS,sum("KF_CUSTO_PROD_VENDIDO") AS KF_CUSTO_PROD_VENDIDO,sum("KF_MATERIAIS") AS KF_MATERIAIS,sum("KF_MATERIA_PRIMA") AS KF_MATERIA_PRIMA,sum("KF_INGREDIENTES") AS KF_INGREDIENTES,sum("KF_EMBALAGENS") AS KF_EMBALAGENS,sum("KF_SEMIACABADOS") AS KF_SEMIACABADOS,sum("KF_PRODUTO_ACABADO") AS KF_PRODUTO_ACABADO,sum("KF_DIFER_INVENTARIO") AS KF_DIFER_INVENTARIO,sum("KF_SERVICOS") AS KF_SERVICOS,sum("KF_SUBCONTRATACAO") AS KF_SUBCONTRATACAO,sum("KF_CMV_REVENDA") AS KF_CMV_REVENDA,sum("KF_RESULT_BRUTO_OPER") AS KF_RESULT_BRUTO_OPER,sum("KF_DESP_COM_VARIAVEIS") AS KF_DESP_COM_VARIAVEIS,sum("KF_COMISSOES_VENDAS") AS KF_COMISSOES_VENDAS,sum("KF_DESP_LOGISTICA_SOMA") AS KF_DESP_LOGISTICA_SOMA,sum("KF_MARGEM_DE_CONTRIB") AS KF_MARGEM_DE_CONTRIB,sum("KF_CUSTOS_FIXOS") AS KF_CUSTOS_FIXOS,sum("KF_MAO_DE_OBRA") AS KF_MAO_DE_OBRA,sum("KF_GASTOS_GERAIS_FAB") AS KF_GASTOS_GERAIS_FAB,sum("KF_DESPESAS_FIXAS") AS KF_DESPESAS_FIXAS,sum("KF_DESP_ADMINISTRATIV") AS KF_DESP_ADMINISTRATIV,sum("KF_DESP_COMERCIAIS_TRADE") AS KF_DESP_COMERCIAIS_TRADE,sum("KF_DESP_COMERCIAIS") AS KF_DESP_COMERCIAIS,sum("KF_DESP_TRADE_MKT") AS KF_DESP_TRADE_MKT,sum("KF_DESP_MKT") AS KF_DESP_MKT,sum("KF_DESP_OPERACIONAL") AS KF_DESP_OPERACIONAL,sum("KF_DESP_TRIBUTARIA") AS KF_DESP_TRIBUTARIA,sum("KF_DESPESAS_PERDA_E_PROV") AS KF_DESPESAS_PERDA_E_PROV,sum("KF_OUTROS_RESULT_OPER") AS KF_OUTROS_RESULT_OPER,sum("KF_OUTRAS_REC_OPER") AS KF_OUTRAS_REC_OPER,sum("KF_RES_ANTES_EF_FINAN") AS KF_RES_ANTES_EF_FINAN,sum("KF_RESULTADO_FINANC") AS KF_RESULTADO_FINANC,sum("KF_DESPESA_FINANCEIRA") AS KF_DESPESA_FINANCEIRA,sum("KF_RES_ANTES_IR_CSLL") AS KF_RES_ANTES_IR_CSLL,sum("KF_RES_APOS_IR_CSLL") AS KF_RES_APOS_IR_CSLl
      FROM "_SYS_BIC"."tirol.co/LUM_DRE_COPA_REAL_ROBO" 
      WHERE (("MANDT" IN ('300') )) 
      AND (("DT_BUDAT" BETWEEN ('{first_day_of_last_year_str}')
      AND ('{last_day_of_current_month_str_ymd}'))) 
      GROUP BY "MANDT", "DT_PERIO", "DT_ANO", "VRGAR", "VERSI", "DT_BUDAT", "ARTNR", "MAKTX", "VTEXT_PAPH2", "DSC_COMPOSTO", "PRCTR", "LTEXT_PRCTR","DscRepresentante"
      """
   
   cursor = con.cursor()
   cursor.execute(declaracao)
   linhas = cursor.fetchall()

   # Criar uma nova planilha
   workbook = Workbook()
   sheet = workbook.active

   #header = ['Data lançamento','Ano', 'Período','Representante','Sublinha desc.', 'Composto Descrição', 'Artigo', 'Artigo desc.','Centro de lucro','Centro de lucro desc', 'QUANTIDADE VENDAS LÍQUIDA', 'PESO VENDAS LÍQUIDO', 'RECEITA BRUTA DE VENDAS', '(=) Deduções Vendas', 'Devoluções Vendas', '-- Descontos Institucionais', 'Over Pricing', '- Impostos S/ Vendas', '-- ICMS', '-- PIS', '-- COFINS', '(=) RECEITA LÍQUIDA VENDAS', '- CUSTO PRODUTO VENDIDO', '-- Materiais', '--- Matéria Prima', '--- Ingredientes', '--- Embalagens', '--- Semiacabados', '--- Produtos Acabados', '--- Diferença de Inventário', '-- Serviços', '--- Subcontratação', '- CMV - REVENDA', '(=) RECEITA LÍQUIDA - CV.', '- Despesas Comerciais Variaveis', '-- Comissões de Vendas', '-- Despesas Logísticas', '(=) MARGEM de CONTRIBUIÇÃO', '- Custos Fixos', '-- Mão de Obra', '-- Gastos Gerais Fabricação', '- Despesas Fixas', '-- Despesas Administrativas', '-- Despesas Comerciais', '--- Desp. Comerciais (somente as desp comerciais)', '--- Desp. Trade Mkt (somente as desp de trade)', '-- Despesas de Marketing', '-- Despesas Operacional', '-- Despesas Tributárias', '-- Despesas com Perda e Provisões', '- Outros Resultados Operacionais', '-- Outras Receitas Operacionais', '(=) RES. ANTES EF. FINANCEIRO', '- Resultado Financeiro', '-- Despesa Financeira', 'RES. ANTES IR/CSLL', 'RES. APÓS IR/CSLL']

   # Adicionar cabeçalho
   header = ['Data lançamento','Ano', 'Período','Representante','Sublinha desc.', 'Composto Descrição', 'Artigo', 'Artigo desc.','Centro de lucro','Centro de lucro desc', 'QUANTIDADE VENDAS LÍQUIDA', 'PESO VENDAS LÍQUIDO', 'RECEITA BRUTA DE VENDAS', '(=) Deduções Vendas', 'Devoluções Vendas', '-- Descontos Institucionais', 'Over Pricing', '- Impostos S/ Vendas', '-- ICMS', '-- PIS', '-- COFINS', '(=) RECEITA LÍQUIDA VENDAS', '- CUSTO PRODUTO VENDIDO', '-- Materiais', '--- Matéria Prima', '--- Ingredientes', '--- Embalagens', '--- Semiacabados', '--- Produtos Acabados', '--- Diferença de Inventário', '-- Serviços', '--- Subcontratação', '- CMV - REVENDA', '(=) RECEITA LÍQUIDA - CV.', '- Despesas Comerciais Variaveis', '-- Comissões de Vendas', '-- Despesas Logísticas', '(=) MARGEM de CONTRIBUIÇÃO', '- Custos Fixos', '-- Mão de Obra', '-- Gastos Gerais Fabricação', '- Despesas Fixas', '-- Despesas Administrativas', '-- Despesas Comerciais', '--- Desp. Comerciais (somente as desp comerciais)', '--- Desp. Trade Mkt (somente as desp de trade)', '-- Despesas de Marketing', '-- Despesas Operacional', '-- Despesas Tributárias', '-- Despesas com Perda e Provisões', '- Outros Resultados Operacionais', '-- Outras Receitas Operacionais', '(=) RES. ANTES EF. FINANCEIRO', '- Resultado Financeiro', '-- Despesa Financeira', 'RES. ANTES IR/CSLL', 'RES. APÓS IR/CSLL']
   sheet.append(header)
   
   for linha in linhas:
      linha_lista = list(linha)
      valores = []
      for valor in linha_lista:
         if isinstance(valor, Decimal):
               valor_float = float(valor.quantize(Decimal('0.000'), rounding=ROUND_DOWN))
               valores.append(valor_float)
         else:
               valores.append(valor)
            
      sheet.append(valores)

   #file_path = 'BI.xlsx'
   file_path = 'T:/Controladoria/01-Controladoria/18 - Lucratividade/Lucratividade/Power BI/Teste Vini/BI.xlsx'
   if os.path.exists(file_path):
      os.remove(file_path)
   
   workbook.save(file_path)

except Error as e:
   print("Erorr - ", e)

finally:
   cursor.close()
   con.close()