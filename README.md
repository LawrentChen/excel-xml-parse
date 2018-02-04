## Description

A snippet to parse [XML Office Format](https://en.wikipedia.org/wiki/Microsoft_Office_XML_formats) Spreadsheet with packages `xml.sax` and `pandas`.

With it, you can:

- specify a directory which contains **Excel XML Spreadsheet** in a same structure and parse all of them
- get a single pandas dataframe as  result

## Story

It seems that only very few people are still using this somewhat deprecated format to transit data. But unfortunately I still encounter them.

By the latest commit of this repo, the two main packages of python to read excel files (`xlrd` and `pandas`) still can not import this format directly. It took me quite a long time to find the solution and I save it here.

## Reference

All credits should go to **Mr.Jonathan Whitemore** and **Mr.Javier Rovegno Campos**. 

- [Question in StackOverFlow](https://stackoverflow.com/questions/33470130/read-excel-xml-xls-file-with-pandas)
- [Issue in pandas #11503](https://github.com/pandas-dev/pandas/issues/11503)
- [Issue in xlrd #156](https://github.com/python-excel/xlrd/issues/156)