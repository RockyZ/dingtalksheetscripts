# SearchOptions

## API描述

搜索选项。用于Sheet.findAll()、Range.find()、Range.findNext()或者Range.findPrevious()等API中指定搜索方式。

## 属性说明

| 属性 | 类型 | 是否必填 | 说明 |
| --- | --- | --- | --- |
| matchEntireCell | boolean | 否 | 匹配整个单元格。 |
| matchCase | boolean | 否 | 匹配大小写。 |
| useRegExp | boolean | 否 | 使用正则表达式。 |
| matchFormulaText | boolean | 否 | 在公式内搜索。 |