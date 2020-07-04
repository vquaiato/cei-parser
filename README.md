This script is aimed to parse the CEI operations, in order to import them on my investments track spreadsheet.

In order to use this you need to download your statements from [cei.b3.com.br](http://cei.b3.com.br) in the [https://cei.b3.com.br/CEI_Responsivo/extrato-bmfbovespa.aspx](https://cei.b3.com.br/CEI_Responsivo/extrato-bmfbovespa.aspx) page. 

Put all downloaded `.xls` files in this script's directory.

Install `xlrd` dependency with pip:
```bash
pip3 install xlrd
```

After that you can run 
```bash
python3 parser.py
```