import datetime, xlrd, glob
from os import devnull

def read_sheet(file_name):
  global fix_number
  fix_number = file_name == "InfoCEI.xls" or file_name == "InfoCEI(1).xls"

  loc = (file_name) 

  wb = xlrd.open_workbook(loc, logfile=open(devnull, 'w'))
  global datemode
  datemode = wb.datemode

  return wb.sheet_by_index(0)

#OPERATIONS
def find_interesting_lines(sheet):
  interest_line = 0
  end_line = 0

  for line in range(0,sheet.nrows):
    if sheet.cell_value(line, 3).startswith("INFORMAÇÕES DE NEGOCIAÇÃO DE ATIVOS"):
      interest_line = line+3
    
    if sheet.cell_value(line, 4).startswith("(C) Compra"):
      end_line = line-4
      break
  
  return {"begin": interest_line, "end": end_line}
def find_columns(sheet, lines):
  xls_to_cols = {
    "Cód":"TICKER",
    "Data Negócio": "DATA",
    "Qtde.Compra": "QTD_COMPRA",
    "Qtd.Venda": "QTD_VENDA",
    "Preço Médio Compra": "PRECO_COMPRA",
    "Preço Médio Venda": "PRECO_VENDA",
    "Posição": "OP"
  }
  cols = dict((el,None) for el in xls_to_cols.values())

  line = lines["begin"]-1
  for col in range(0,sheet.ncols):
    cell = sheet.cell_value(line, col)
    if cell in xls_to_cols.keys():
      cols[xls_to_cols[cell]] = col

      if cols["OP"] is not None:
        break

  return cols
def format_price(preco, qtd):
  fpreco = float(preco.replace(",", "."))

  if fix_number == True:
    index = preco.index(",")
    fpreco = float(f"{preco[0:index-1]}.{preco[index-1:index]}{preco[index+1:-2]}")

  return str(round(fpreco * qtd, 2)).replace(".",",")
def read_operations(sheet):
  lines = find_interesting_lines(sheet)
  if lines["begin"] is 0:
    return []

  cols = find_columns(sheet, lines)
  
  data = []
  for line in range(lines["begin"], lines["end"]):
    op = sheet.cell_value(line, cols["OP"])

    qtd = int(sheet.cell_value(line, cols["QTD_COMPRA"]) if op == "COMPRADA" else sheet.cell_value(line, cols["QTD_VENDA"]))

    data.append({
      "ticker": sheet.cell_value(line, cols["TICKER"])[0:5],
      "data": datetime.datetime(*xlrd.xldate_as_tuple(sheet.cell_value(line, cols["DATA"]), datemode)).strftime("%d/%m/%Y"),
      "qtd": qtd,
      "preço": format_price(sheet.cell_value(line, cols["PRECO_COMPRA"]) if op == "COMPRADA" else sheet.cell_value(line, cols["PRECO_VENDA"]), qtd),
      "op": "Comprar" if op == "COMPRADA" else "Vender"
    })

  return data

#EARNINGS
def find_earnings_lines(sheet):
  interest_line = 0
  end_line = 0

  for line in range(0,sheet.nrows):
    if sheet.cell_value(line, 5).startswith("PROVENTOS EM DINHEIRO - CREDITADOS"):
      interest_line = line+2
    elif sheet.cell_value(line, 5).startswith("TOTAL CREDITADO"):
      end_line = line
      break
  
  return {"begin": interest_line, "end": end_line}
def find_earnings_columns(sheet, lines):
  xls_to_cols = {
    "Cód. Neg.":"TICKER",
    "Creditado No Mês": "VALOR",
  }
  cols = dict((el,None) for el in xls_to_cols.values())

  line = lines["begin"]-1
  for col in range(0,sheet.ncols):
    cell = sheet.cell_value(line, col)
    if cell in xls_to_cols.keys():
      cols[xls_to_cols[cell]] = col

      if cols["VALOR"] is not None:
        break

  return cols
def find_report_date(sheet):
  report_raw_date = sheet.cell_value(6, 0).split(" ")

  months = {
    "Janeiro": 1,
    "Fevereiro": 2,
    "Março": 3,
    "Abril": 4,
    "Maio": 5,
    "Junho": 6,
    "Julho": 7,
    "Agosto": 8,
    "Setembro": 9,
    "Outubro": 10,
    "Novembro": 11,
    "Dezembro": 12,
  }

  return f"01/{months[report_raw_date[0]]:02}/{report_raw_date[-1]}"
def read_earnings(sheet):
  lines = find_earnings_lines(sheet)
  if lines["begin"] is 0:
    return []
  
  cols = find_earnings_columns(sheet, lines)
  base_earnings_date = find_report_date(sheet)
  data = []
  for line in range(lines["begin"], lines["end"]):
    ticker = sheet.cell_value(line, cols["TICKER"]+1)[0:6]
    if ticker[-1] is not '1' and ticker[-1] is not ' ':
      continue

    data.append({
      "ticker": ticker,
      "valor": str(sheet.cell_value(line, cols["VALOR"])).replace(".",","),
      "data": base_earnings_date 
    })

  return data

def process_all_files():
  all_ops = []
  all_earnings = []
  for file in glob.glob("*.xls"):
    print(f"> Processing file: {file}")

    sheet = read_sheet(file)
    operations = read_operations(sheet)
    all_earnings.append(read_earnings(sheet))
    all_ops.append(operations)

  return {
    "ops": [item for sublist in all_ops for item in sublist],
    "earnings": [item for sublist in all_earnings for item in sublist]
  }

#INFRA
def print_operations(ops):
  print(f"{'>'*10} OPERATIONS")
  print(f"Operação\tQtd.\tTicker\tQtd.\tPreço\tExecutada\tIgnorar")
  for op in ops:
    print(f"{op['op']}\t{op['qtd']}\t{op['ticker']}\t{op['qtd']}\tR$ {op['preço']}\tExecutada\t{op['data']}")
def print_earnings(earnings):
  print(f"{'>'*10} EARNINGS")
  print(f"Base Date\tTicker\tShares\tValue")
  for e in earnings:
    print(f"{e['data']}\t{e['ticker']}\t\t{e['valor']}")
def print_reports(reports):
  print_operations(reports["ops"])
  print_earnings(reports["earnings"])

def main():
  reports = process_all_files()
  print_reports(reports)

main()