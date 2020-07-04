import datetime, xlrd, glob, sys
from os import devnull

def read_sheet(file_name):
  global fix_number
  fix_number = True if file_name == "InfoCEI.xls" or file_name == "InfoCEI(1).xls" else False
  loc = (file_name) 

  wb = xlrd.open_workbook(loc, logfile=open(devnull, 'w'))
  global datemode
  datemode = wb.datemode
  return wb.sheet_by_index(0)

def find_interesting_lines(sheet):
  interest_line = 0
  end_line = 0

  for line in range(0,1000):
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
  for col in range(0,100):
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
  if lines["begin"] == 0:
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

def process_all_files():
  all_ops = []

  for file in glob.glob("*.xls"):
    print(f"File: {file}")

    sheet = read_sheet(file)
    operations = read_operations(sheet)
    all_ops.append(operations)

  return [item for sublist in all_ops for item in sublist]

ops = process_all_files()

print(f"Operação\tQtd.\tTicker\tQtd.\tPreço\tExecutada\tIgnorar")
for op in ops:
  print(f"{op['op']}\t{op['qtd']}\t{op['ticker']}\t{op['qtd']}\tR$ {op['preço']}\tExecutada\t{op['data']}")