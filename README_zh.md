## 介绍

一小段使用`xml.sax`和`pandas`包来解析 [XML Office Format](https://en.wikipedia.org/wiki/Microsoft_Office_XML_formats) 数据表的代码。

使用它，你可以：

- 指定一个存有相同结构的**Excel XML**数据表的路径，全数解析它们
- 作为结果，得到一个pandas dataframe

## 故事

现在应该只有很少一部分人仍在使用这种大概已经过时的格式来传输数据。但很不幸的我还是碰上了它们。

截止到这个仓库最近的一次提交，python用于读取excel文件的两个主要的包 (`xlrd` 和 `pandas`) 依然不能直接读入这种格式。我花了相当长一段时间来寻找解决办法，代码就保存在这里。

## 参考

所有贡献都应该归功于 **Jonathan Whitemore** 先生和 **Javier Rovegno Campos** 先生. 

- [Question in StackOverFlow](https://stackoverflow.com/questions/33470130/read-excel-xml-xls-file-with-pandas)
- [Issue in pandas #11503](https://github.com/pandas-dev/pandas/issues/11503)
- [Issue in xlrd #156](https://github.com/python-excel/xlrd/issues/156)