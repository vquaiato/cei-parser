# CEI parser
This script is aimed to parse the CEI operations and earnings, in order to import them on my investments track spreadsheet.

_Disclaimer:_ the output generted is tailored to my needs. If for any reason you would like to use this script and the output doesn't fit your needs, just change it. The main functions to change are: [`read_operations`](https://github.com/vquaiato/cei-parser/blob/master/parser.py#L70) and [`read_earnings`](https://github.com/vquaiato/cei-parser/blob/master/parser.py#L161).

## Running
In order to use this you need to download your statements from [cei.b3.com.br](http://cei.b3.com.br) in the [https://cei.b3.com.br/CEI_Responsivo/extrato-bmfbovespa.aspx](https://cei.b3.com.br/CEI_Responsivo/extrato-bmfbovespa.aspx) page. 

Put all downloaded `.xls` files in this script's directory. File's names are not important as long as they are `.xls` files the script will read all of them.

Install `xlrd` dependency with pip:
```bash
pip3 install xlrd
```

After that you can run 
```bash
python3 parser.py
```

## Caveats
For me 2 files had formatting issues, thus I had to make this _not so pretty fix_ [here](https://github.com/vquaiato/cei-parser/blob/master/parser.py#L6). It this isn't required for you, just comment it out.

## Warning
This project is not actively maintained. I've done as a pet 2h project.