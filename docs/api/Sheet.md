# Sheet

## activate()

激活自身。

### 返回结果

Sheet - 当前Sheet实例。

### 示例代码

```
Workbook.getSheet('Sheet1').activate();
```

## getId()

获取ID。

### 返回结果

string - 工作表的ID。

### 示例代码

```
const sheetId = Workbook.getActiveSheet().getId();
```

## getIndex()

获取工作表的索引。

### 返回结果

number - 工作表的索引。

### 示例代码

```
const index = Workbook.getActiveSheet().getIndex();
```

## getName()

获取工作表的名字。

### 返回结果

string - 工作表的名字。

### 示例代码

```
const name = Workbook.getActiveSheet().getName();
```

## setTabColor(color)

设置Sheet标签颜色。

### 返回结果

Sheet - 当前Sheet实例。

### 参数说明

|  |  |  |  |
| --- | --- | --- | --- |
| 参数 | 类型 | 是否必传 | 说明 |
| color | string | null | 是 | 颜色。 |

### 示例代码

```
Workbook.getActiveSheet().setTabColor('yellow');
```

## getTabColor()

获取Sheet标签颜色。

### 返回结果

Color | null - 颜色。null代表没有设置颜色。

### 示例代码

```
const tabColor = Workbook.getActiveSheet().getTabColor();
```

## setName(name)

设置工作表的名字。

### 返回结果

Sheet - 当前Sheet实例。

### 参数说明

|  |  |  |  |
| --- | --- | --- | --- |
| 参数 | 类型 | 是否必传 | 说明 |
| name | string | 是 | 新名字。 |

### 示例代码

```
Workbook.getActiveSheet().setName('New name');
```

## getVisibility()

获取可见性。

### 返回结果

'visible' | 'hidden' - 可见性。

### 示例代码

```
const visibility = Workbook.getActiveSheet().getVisibility();
```

## getGridlinesVisibility()

获取网格线可见性。

### 返回结果

'visible' | 'hidden' - 可见性。

### 示例代码

```
const gridlinesVisibility = Workbook.getActiveSheet().getGridlinesVisibility();
```

## setGridlinesVisibility(visibility)

设置网格线可见性。

### 返回结果

Sheet - 当前Sheet实例。

### 参数类型

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| visibility | 'visible' | 'hidden' | 是 | 可见性。 |

### 示例代码

```
Workbook.getActiveSheet().setGridlinesVisibility('hidden');
```

## getFrozenColumnCount()

获取冻结的列。

### 返回结果

number - 列数，从1开始，0表示不冻结。

### 示例代码

```
const frozenColumnCount = Workbook.getActiveSheet().getFrozenColumnCount();
```

## getFrozenRowCount()

获取冻结的行。

### 返回结果

number - 行数，从1开始，0表示不冻结。

### 示例代码

```
const frozenRowCount = Workbook.getActiveSheet().getFrozenRowCount();
```

## setFrozenColumnCount(columnCount)

设置冻结列。

### 返回结果

Sheet - 当前Sheet实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| columnCount | number | 是 | 列数，从 1 开始，0表示不冻结。 |

### 示例代码

```
Workbook.getActiveSheet().setFrozenColumnCount(5);
```

## setFrozenRowCount(rowCount)

设置冻结行。

### 返回结果

Sheet - 当前Sheet实例。

### 参数类型

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| rowCount | number | 是 | 行数，从 1 开始。 |

### 示例代码

```
Workbook.getActiveSheet().setFrozenRowCount(5);
```

## getColumnCount()

获取当前Sheet的列数。

### 返回结果

number - 列数。

### 示例代码

```
const columnCount = Workbook.getActiveSheet().getColumnCount();
```

## getRowCount()

获取当前Sheet的行数。

### 返回结果

number - 行数。

### 示例代码

```
const rowCount = Workbook.getActiveSheet().getRowCount();
```

## setActiveRange(a1Notation)

激活选区。

### 返回结果

Range - 激活的选区。

### 示例代码

```
Workbook.getActiveSheet().setActiveRange("A1:C10");
```

## setActiveRange(range)

激活选区。

### 返回结果

Range - 激活的选区。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| range | Range | 是 | 选区。 |

### 示例代码

```
const range = Workbook.getRange(0, 0, 2, 2);
Workbook.getActiveSheet().setActiveRange(range);
```

## getActiveRange()

获取 activeCell 所在的选区。

### 返回结果

Range | null - activeCell 所在的选区

### 示例代码

```
const range = Workbook.getActiveSheet().getActiveRange();
```

## getEntireSheet()

获取整表区域。

### 返回结果

Range - 整表区域

### 示例代码

```
const range = Workbook.getActiveSheet().getEntireSheet();
```

## getEntireRows(row, rowCount)

获取整行。

### 返回结果

Range | null - 整行区域

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| row | number | 是 | 起始行。 |
| rowCount | number | 否 | 行数。 |

### 示例代码

```
const range = Workbook.getActiveSheet().getEntireRows(0, 3);
```

## getEntireColumns(column, columnCount)

获取整列。

### 返回结果

Range | null - 整列区域

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| column | number | 是 | 起始列。 |
| columnCount | number | 否 | 列数。 |

### 示例代码

```
const range = Workbook.getActiveSheet().getEntireColumns(0, 3);
```

## getRange(a1Notation)

获取选区范围。

### 返回结果

Range - 选区

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| a1Notation | string | 是 | 选区地址。 |

### 示例代码

```
const range = Workbook.getActiveSheet().getRange('A1:C10');
```

## getRange(row, column, rowCount, columnCount)

获取选区范围。

### 返回结果

Range - 选区

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| row | number | 是 | 行，从 0 开始。 |
| column | number | 是 | 列，从 0 开始。 |
| rowCount | number | 是 | 行数，从 1 开始。 |
| columnCount | number | 是 | 列数，从 1 开始。 |

### 示例代码

```
const range = Workbook.getActiveSheet().getRange(0, 0, 2, 2);
```

## setActiveRangeList(a1Notations)

激活多选区。

### 返回结果

RangeList - 多选区

### 参数说明

| 参数 | 类型 |
| --- | --- |
| a1Notations | string[] |
| rangeList | RangeList |

### 示例代码

```
const rangeList = Workbook.getActiveSheet().setActiveRangeList(['A1:C10', 'D2:E5']);
```

## setActiveRangeList(rangeList)

激活多选区。

### 返回结果

RangeList - 多选区

### 参数说明

| 参数 | 类型 |
| --- | --- |
| rangeList | RangeList |

### 示例代码

```
const rangeList = Workbook.getActiveSheet().getRangeList(['A1:C10', 'D2:E5']);
Workbook.getActiveSheet().setActiveRangeList(rangeList);
```

## getActiveRangeList()

获取激活的多选区。

### 返回结果

RangeList - 多选区

### 示例代码

```
const rangeList = Workbook.getActiveSheet().getActiveRangeList();
```

## getRangeList(a1Notations)

获取多选区。

### 返回结果

RangeList - 多选区

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| a1Notations | string[] | 是 | 选区地址数组。 |

### 示例代码

```
const rangeList = Workbook.getActiveSheet().getRangeList(['A1:C10', 'D2:E5']);
```

## getActiveCell()

获取active cell。

### 返回结果

Range - active cell

### 示例代码

```
const range = Workbook.getActiveSheet().getActiveCell();
```

## getCell(row, column)

获取单元格。

### 返回结果

Range - 单元格

### 参数说明

|  |  |  |  |
| --- | --- | --- | --- |
| 参数 | 类型 | 是否必传 | 说明 |
| row | number | 是 | 行。 |
| column | number | 是 | 列。 |

### 示例代码

```
const range = Workbook.getActiveSheet().getCell(0, 0);
```

## setRowsVisibility(row, rowCount, visibility)

显示/隐藏多行。

### 返回结果

Sheet - 当前Sheet实例

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| row | number | 是 | 行。 |
| rowCount | number | 是 | 行数。 |
| visibility | 'visible' | 'hidden' | 是 | 可见性。 |

### 示例代码

```
// 隐藏第一行
Workbook.getActiveSheet().setRowsVisibility(0, 1, 'hidden');
```

## setRowVisibility(row, visibility)

显示/隐藏单行。

### 返回结果

Sheet - 当前Sheet实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| row | number | 是 | 行。 |
| visibility | 'visible' | 'hidden' | 是 | 可见性。 |

### 示例代码

```
// 隐藏第二行
Workbook.getActiveSheet().setRowVisibility(1, 'hidden');
```

## setColumnsVisibility(column, columnCount, visibility)

显示/隐藏多列。

### 返回结果

Sheet - 当前Sheet实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| column | number | 是 | 列。 |
| columnCount | number | 是 | 列数。 |
| visibility | 'visible' | 'hidden' | 是 | 可见性。 |

### 示例代码

```
// 隐藏第二行
Workbook.getActiveSheet().setRowVisibility(1, 'hidden');
```

## setColumnVisibility(column, visibility)

显示/隐藏单列。

### 返回结果

Sheet - 当前Sheet实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| column | number | 是 | 列。 |
| visibility | 'visible' | 'hidden' | 是 | 可见性。 |

### 示例代码

```
// 隐藏第二列
Workbook.getActiveSheet().setColumnVisibility(1, 'hidden');
```

## setRowsHeight(row, rowCount, height)

设置多行的高度。

### 返回结果

Sheet - 当前Sheet实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| row | number | 是 | 行。 |
| rowCount | number | 是 | 行数。 |
| height | number | 是 | 高度。 |

### 示例代码

```
// 设置第一行到第三行，高度为50
Workbook.getActiveSheet().setRowsHeight(0, 3, 50);
```

## setRowHeight(row, height)

设置单行的高度。

### 返回结果

Sheet - 当前Sheet实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| row | number | 是 | 行。 |
| height | number | 是 | 高度。 |

### 示例代码

```
Workbook.getActiveSheet().setRowHeight(0, 50);
```

## getRowHeight(row)

获取行高。

### 返回结果

Sheet - 当前Sheet实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| row | number | 是 | 行。 |

### 示例代码

```
const height = Workbook.getActiveSheet().getRowHeight(0);
```

## autofitRows(row, rowCount)

自适应行高。

### 返回结果

Sheet - 当前Sheet实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| row | number | 是 | 行。 |
| rowCount | number | 是 | 行数。 |

### 示例代码

```
Workbook.getActiveSheet().autofitRows(0, 3);
```

## autofitRow(row)

自适应行高。

### 返回结果

Sheet - 当前Sheet实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| row | number | 是 | 行。 |

### 示例代码

```
Workbook.getActiveSheet().autofitRow(0);
```

## setColumnsWidth(column, columnCount, width)

设置列宽。

### 返回结果

Sheet - 当前Sheet实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| column | number | 是 | 列。 |
| columnCount | number | 是 | 列数。 |
| width | number | 是 | 宽度。 |

### 示例代码

```
// 设置A到C列，宽度为200
Workbook.getActiveSheet().setColumnsWidth(0, 3, 200);
```

## setColumnWidth(column, width)

设置列宽。

### 返回结果

Sheet - 当前Sheet实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| column | number | 是 | 列。 |
| width | number | 是 | 高度。 |

### 示例代码

```
Workbook.getActiveSheet().setColumnWidth(0, 200);
```

## getColumnWidth(column)

获取列宽。

### 返回结果

number - 列宽。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| column | number | 是 | 列。 |

### 示例代码

```
const width = Workbook.getActiveSheet().getColumnWidth(0);
```

## autofitColumns(column, columnCount)

自适应列宽。

### 返回结果

number - 列宽。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| column | number | 是 | 列。 |
| columnCount | number | 是 | 列。 |

### 示例代码

```
Workbook.getActiveSheet().autofitColumns(0, 3);
```

## autofitColumn(column)

自适应列宽。

### 返回结果

number - 列宽。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| column | number | 是 | 列。 |

### 示例代码

```
Workbook.getActiveSheet().autofitColumn(0);
```

## insertRowBefore(row)

在 row 行之前插入 1 行。

### 返回结果

Sheet - 当前Sheet实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| row | number | 是 | 行。 |

### 示例代码

```
Workbook.getActiveSheet().insertRowBefore(0);
```

## insertRowsBefore(row, rowCount)

在 row 行之前插入 n 行。

### 返回结果

Sheet - 当前Sheet实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| row | number | 是 | 行。 |
| rowCount | number | 是 | 行数。 |

### 示例代码

```
Workbook.getActiveSheet().insertRowsBefore(0, 10);
```

## insertRowAfter(row)

在 row 行之后插入 1 行。

### 返回结果

Sheet - 当前Sheet实例。

### 参数说明

|  |  |  |  |
| --- | --- | --- | --- |
| 参数 | 类型 | 是否必传 | 说明 |
| row | number | 是 | 行。 |

### 示例代码

```
Workbook.getActiveSheet().insertRowAfter(0);
```

## insertRowsAfter(row, rowCount)

在 row 行之后插入 n 行。

### 返回结果

Sheet - 当前Sheet实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| row | number | 是 | 行。 |
| rowCount | number | 是 | 行数。 |

### 示例代码

```
Workbook.getActiveSheet().insertRowsAfter(0, 10);
```

## insertColumnBefore(column)

在 column 行之前插入 1 列。

### 返回结果

Sheet - 当前Sheet实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| column | number | 是 | 列。 |

### 示例代码

```
Workbook.getActiveSheet().insertColumnBefore(0);
```

## insertColumnsBefore(column, columnCount)

在 column 行之前插入 n 列。

### 返回结果

Sheet - 当前Sheet实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| column | number | 是 | 列。 |
| columnCount | number | 是 | 列数。 |

### 示例代码

```
Workbook.getActiveSheet().insertColumnsBefore(0, 10);
```

## insertColumnAfter(column)

在 column 行之后插入 1 列。

### 返回结果

Sheet - 当前Sheet实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| column | number | 是 | 列。 |

### 示例代码

```
Workbook.getActiveSheet().insertColumnAfter(0);
```

## insertColumnsAfter(column, columnCount)

在 column 行之后插入 n 列。

### 返回结果

Sheet - 当前Sheet实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| column | number | 是 | 列。 |
| columnCount | number | 是 | 列数。 |

### 示例代码

```
Workbook.getActiveSheet().insertColumnsAfter(0, 10);
```

## deleteRows(row, rowCount)

删除行。

### 返回结果

Sheet - 当前Sheet实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| row | number | 是 | 行。 |
| rowCount | number | 是 | 行数。 |

### 示例代码

```
Workbook.getActiveSheet().deleteRows(0, 10);
```

## deleteRow(row)

删除行。

### 返回结果

Sheet - 当前Sheet实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| row | number | 是 | 行。 |

### 示例代码

```
Workbook.getActiveSheet().deleteRows(0);
```

## deleteColumns(column, columnCount)

删除列。

### 返回结果

Sheet - 当前Sheet实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| column | number | 是 | 列。 |
| columnCount | number | 是 | 列数。 |

### 示例代码

```
Workbook.getActiveSheet().deleteColumns(0， 10);
```

## deleteColumn(column)

删除列。

### 返回结果

Sheet - 当前Sheet实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| column | number | 是 | 列。 |

### 示例代码

```
Workbook.getActiveSheet().deleteColumn(0);
```

## findAll(text, findOptions)

根据指定的条件查找匹配给定字符串的所有Range。

### 返回结果

Range[] - 所有匹配到的区域的数组。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| text | string | 是 | 要查找的字符串。 |
| findOptions | SearchOptions | 否 | 其他搜索条件，包括搜索是否需要匹配整个单元格、区分大小写、是否使用正则、是否匹配公式。 |

### 示例代码

```
const ranges = sheet.findAll('str');
const ranges = sheet.findAll('str', {
 matchEntireCell: true,
  matchCase: true,
});
```

## replaceAll(text, replaceText, replaceOptions)

根据当前工作表中指定的条件查找并替换给定的字符串。

### 返回结果

number - 执行的替换数，可能与匹配单元格的数目不同。

### 参数说明

|  |  |  |  |
| --- | --- | --- | --- |
| 参数 | 类型 | 是否必传 | 说明 |
| text | string | 是 | 要查找的字符串。 |
| replaceText | string | 是 | 替换原始字符串的字符串。 |
| replaceOptions | SearchOptions |  | 其他替换条件，包括是否需要匹配整个单元格、区分大小写、是否使用正则、是否匹配公式。 |

### 示例代码

```
const replaceCount = sheet.replaceAll('str', 'otherStr');
const replaceCount = sheet.replaceAll('str', 'otherStr', {  matchEntireCell: true });
```

## filter(range)

创建Filter。

### 返回结果

Filter - Filter 实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| range | Range | 是 | 区域。 |

### 示例代码

```
const range = sheet.getRange('A1:C5');
sheet.filter(range);
```

## getFilter()

获取当前sheet上的filter实例。

### 返回结果

Filter ｜ null - Filter 实例。

### 示例代码

```
const filter = sheet.getFilter();
```

## deleteFilter()

删除当前sheet上的filter。

### 返回结果

void -

### 示例代码

```
sheet.deleteFilter();
```