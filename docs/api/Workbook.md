# Workbook

## insertSheet()

插入工作表。

### 返回结果

Sheet - 新增加的工作表

### 示例代码

```
Workbook.insertSheet();
```

## insertSheet(name)

插入工作表，并指定名字。

### 返回结果

Sheet - 新增加的工作表

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| name | string | 是 | 指定新工作表名字 |

### 示例代码

```
Workbook.insertSheet('A new sheet');
```

## insertSheet(index)

插入工作表，并指定位置。

### 返回结果

Sheet - 新增加的工作表

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| index | number | 是 | 指定新工作表位置，从 0 开始，在index 之后插入工作表。 |

### 示例代码

```
Workbook.insertSheet(0);
```

## insertSheet(name, index)

插入工作表，并指定名字和位置。

### 返回结果

Sheet - 新增加的工作表

### 示例代码

```
Workbook.insertSheet('hello world', 0);
```

## moveSheet(sheet, targetIndex)

移动 Worksheet 到指定位置。

### 返回结果

Workbook - 当前表格实例

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| sheet | Sheet | 是 | 要移动的目标工作表 |
| targetIndex | number | 是 | 目标的索引位置 |

### 示例代码

```
const sheet = Workbook.getSheetByIndex(0);
Workbook.moveSheet(sheet, 1);
```

## deleteSheet(name)

删除工作表。

### 返回结果

Workbook - 当前表格实例

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| name | Sheet | string | 是 | 工作表对象、ID或名字 |

### 示例代码

```
Workbook.deleteSheetByName('Sheet1');
```

## setActiveSheet(sheet)

激活工作表。

### 返回结果

Sheet - 激活的工作表

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| sheet | Sheet | 是 | 想激活的工作表对象、ID或名字 |

### 示例代码

```
const sheet = Workbook.getSheetByName('Sheet1');
Workbook.setActiveSheet(sheet);
```

## getSheets()

获取所有的工作表。

### 返回结果

Sheet[] - 所有的工作表

### 示例代码

```
const sheets = Workbook.getSheets();
```

## getSheetCount()

获取工作表的数量。

### 返回结果

number - 工作表的数量

### 示例代码

```
const sheetCount = Workbook.getSheetCount();

// 等价于
const sheets = Workbook.getSheets();
const sheetCount = sheets.length;
```

## getActiveSheet()

获取当前激活的工作表。

### 返回结果

Sheet - 当前激活的工作表

### 示例代码

```
const sheet = Workbook.getActiveSheet();
sheet.getRange(0, 0, 2, 2).setBackgroundColor('yellow');
```

## getSheet(key)

根据 ID 或者 name 获取工作表。

### 返回结果

Sheet | null - 查询到的工作表。null表示没有查到。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| key | string | 是 | 工作表的 ID 或者 name |

### 示例代码

```
const sheet1 = Workbook.getSheet('id1');
const sheet2 = Workbook.getSheet('Sheet2');
```

## getRange(a1Notation)

获取当前工作表的选区范围。

### 返回结果

Range - 选区

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| a1Notation | string | 是 | 选区地址 |

### 示例代码

```
const range = Workbook.getRange('A1:B4');

// 等价于
const range = Workbook.getActiveSheet().getRange('A1:B4');
```

## getRange(row, column, rowCount, columnCount)

获取当前工作表的选区范围。

### 返回结果

Range - 选区

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| row | number | 是 | 行，从 0 开始 |
| column | number | 是 | 列，从 0 开始 |
| rowCount | number | 是 | 行数，从 1 开始 |
| columnCount | number | 是 | 列数，从 1 开始 |

### 示例代码

```
const range = Workbook.getRange(0, 0, 2, 2);

// 等价于
const range = Workbook.getActiveSheet().getRange(0, 0, 2, 2);
```

## getActiveCell()

获取activeCell。

### 返回结果

Range - activeCell

### 示例代码

```
const range = Workbook.getActiveCell();

// 等价于
const range = Workbook.getActiveSheet().getActiveCell();
```

## getActiveRange()

获取激活的选区。如果激活了多个不连续区域，则返回activeCell所在的选区。

### 返回结果

Range | null - 当前激活的选区。如果当前选中了图表等，则返回null。

## getRangeList(a1Notations)

获取多选区。

### 返回结果

RangeList | null - 多选区

### 示例代码

```
const rangeList = Workbook.getRangeList(['A1:B10', 'D2:E5']);

// 等价于
const rangeList = Workbook.getActiveSheet().getRangeList(['A1:B10', 'D2:E5']);
```

## getActiveRangeList()

获取激活的多选区。

### 返回结果

RangeList - 激活的多选区

### 示例代码

```
const rangeList = Workbook.getActiveRangeList();

// 等价于
const rangeList = Workbook.getActiveSheet().getActiveRangeList();
```

## newFilterCriteriaBuilder()

新建一个筛选条件规则的生成器。

### 返回结果

FilterCriteriaBuilder - 筛选条件规则生成器

### 示例代码

```
const criteria = Workbook.newFilterCriteriaBuilder()
 .setVisibleValues(['3', '5'])
  .build();
const filter = Workbook.getActiveSheet().getFilter();
filter.setColumnFilterCriteria(1, criteria);
```