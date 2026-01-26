# Range

## isEntireSheet()

是否是整个工作表区域。

### 返回结果

boolean - 是否是整个工作表区域。

### 示例代码

```
const isEntireSheet = Workbook.getRange('A1').isEntireSheet();
```

## isEntireRow()

是否是整行。

### 返回结果

boolean - 是否是整行。

### 示例代码

```
const isEntireRow = Workbook.getRange('A1').isEntireRow();
```

## isEntireColumn()

是否是整列。

### 返回结果

boolean - 是否是整列。

### 示例代码

```
const isEntireColumn = Workbook.getRange('A1').isEntireColumn();
```

## isCell()

是否是一个单元格。

### 返回结果

boolean - 是否是一个单元格。

### 示例代码

```
const isCell = Workbook.getRange('A1').isCell();
```

## activate()

激活选区。

### 返回结果

Range - 区域。

### 示例代码

```
Workbook.getRange('A1').activate()
```

## clear()

清除样式和数据。

### 返回结果

Range - 区域。

### 示例代码

```
Workbook.getRange('A1').clear();
```

## clearFormat()

清除样式。

### 返回结果

Range - 区域。

### 示例代码

```
Workbook.getRange('A1').clearFormat();
```

## clearData()

清除数据。

### 返回结果

Range - 区域。

### 示例代码

```
Workbook.getRange('A1').clearData();
```

## setBackgroundColor(backgroundColor)

设置选区单元格背景颜色。

### 返回结果

Range - 区域。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| backgroundColor | string | null | 是 | 背景颜色, null 表示清除。使用十六进制字符串形式表示颜色，例如 '#ff0000' 表示红色。 |

### 示例代码

```
Workbook.getRange('A1').setBackgroundColor('#ff0000');
```

## getBackgroundColor()

获取左上角单元格背景色。

### 返回结果

Color - 背景颜色。

### 示例代码

```
const color = Workbook.getRange('A1').getBackgroundColor();
```

## getBackgroundColors()

获取区域单元格的背景色。

### 返回结果

Array<Array<Color>>> - 背景颜色，二维数组。

### 示例代码

```
const colors = Workbook.getRange('A1').getBackgroundColors();
```

## setFontColor(fontColor)

设置字体颜色。

### 返回结果

Range - 区域。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| fontColor | string | null | 是 | 字体颜色, null 表示清除。使用十六进制字符串形式表示颜色，例如 '#ff0000' 表示红色。 |

### 示例代码

```
Workbook.getRange('A1').setFontColor('#ff0000');
```

## getFontColor()

获取左上角单元格字体颜色。

### 返回结果

Color - 字体颜色。

### 示例代码

```
const color = Workbook.getRange('A1').getFontColor();
```

## getFontColors()

获取区域单元格字体颜色。

### 返回结果

Array<Array<Color>>> - 字体颜色。

### 示例代码

```
const colors = Workbook.getRange('A1').getFontColors();
```

## setFontFamily(fontFamily)

设置字体。

### 返回结果

Range - 区域。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| fontFamily | string | null | 是 | 字体, null 表示清除。 |

### 示例代码

```
Workbook.getRange('A1').setFontFamily('arial');
```

## getFontFamily()

获取左上角单元格字体。

### 返回结果

string - 字体。

### 示例代码

```
const family = Workbook.getRange('A1').getFontFamily();
```

## getFontFamilies()

获取区域单元格字体。

### 返回结果

Array<Array<string>> - 字体。

### 示例代码

```
const families = Workbook.getRange('A1').getFontFamilies();
```

## setFontSize(fontSize)

设置字体大小。

### 返回结果

Range - 区域。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| fontSize | number | null | 是 | 字体大小, null 表示清除。 |

### 示例代码

```
Workbook.getRange('A1').setFontSize(20);
```

## getFontSize()

获取左上角单元格字体大小。

### 返回结果

number - 字体大小。

### 示例代码

```
const size = Workbook.getRange('A1').getFontSize();
```

## getFontSizes()

获取区域单元格字体大小。

### 返回结果

Array<Array<number>> - 字体大小。

### 示例代码

```
const sizes = Workbook.getRange('A1').getFontSizes();
```

## setFontWeight(fontWeight)

设置字体粗体。

### 返回结果

Range - 区域。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| fontWeight | 'bold' | 'normal' | null | 是 | 字体粗细, null 表示清除。 |

### 示例代码

```
Workbook.getRange('A1').setFontWeight('bold');
```

## getFontWeight()

获取左上角单元格字体粗体。

### 返回结果

'bold' | 'normal' - 字体粗体。

### 示例代码

```
const weight = Workbook.getRange('A1').getFontWeight();
```

## getFontWeights()

获取区域单元格字体粗体。

### 返回结果

Array<Array<'bold' | 'normal'>> - 字体粗体。

### 示例代码

```
const weights = Workbook.getRange('A1').getFontWeights();
```

## setFontStyle(fontStyle)

设置字体样式。

### 返回结果

Range - 区域。

### 参数说明

|  |  |  |  |
| --- | --- | --- | --- |
| 参数 | 类型 | 是否必传 | 说明 |
| fontStyle | 'italic' | 'normal' | null | 是 | 字体样式, null 表示清除。 |

### 示例代码

```
Workbook.getRange('A1').setFontStyle('italic');
```

## getFontStyle()

获取左上角单元格字体样式。

### 返回结果

'italic' | 'normal' - 字体样式。

### 示例代码

```
const style = Workbook.getRange('A1').getFontStyle();
```

## getFontStyles()

获取区域单元格字体样式。

### 返回结果

Array<Array<'italic' | 'normal'>> - 字体样式。

### 示例代码

```
const styles = Workbook.getRange('A1').getFontStyles();
```

## setIndent(indent)

设置缩进。

### 返回结果

Range - 区域。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| indent | number | null | 是 | 缩进，必须为正整数，null 表示清除。 |

### 示例代码

```
Workbook.getRange('A1').setIndent(3);
```

## getIndent()

获取左上角单元格缩进。

### 返回结果

number - 缩进。

### 示例代码

```
const indent = Workbook.getRange('A1').getIndent();
```

## getIndents()

获取区域单元格缩进。

### 返回结果

Array<Array<number>> - 缩进。

### 示例代码

```
const indents = Workbook.getRange('A1').getIndents();
```

## setHorizontalAlignment(alignment)

设置水平对齐方式。

### 返回结果

Range - 区域。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| horizontalAlignment | 'left' | 'center' | 'right' | 'general' | null | 是 | 水平对齐方式, null 表示清除。 |

### 示例代码

```
Workbook.getRange('A1').setHorizontalAlignment('left');
```

## getHorizontalAlignment()

获取左上角单元格水平对齐方式。

### 返回结果

'left' | 'center' | 'right' | 'general' - 水平对齐方式。

### 示例代码

```
const alignment = Workbook.getRange('A1').getHorizontalAlignment();
```

## getHorizontalAlignments()

获取区域单元格水平对齐方式。

### 返回结果

Array<Array<'left' | 'center' | 'right' | 'general'>> - 水平对齐方式。

### 示例代码

```
const alignments = Workbook.getRange('A1').getHorizontalAlignments();
```

## setVerticalAlignment(alignment)

设置垂直对齐方式。

### 返回结果

Range - 区域。

### 参数说明

|  |  |  |  |
| --- | --- | --- | --- |
| 参数 | 类型 | 是否必传 | 说明 |
| verticalAlignment | 'top' | 'middle' | 'bottom' | null | 是 | 字体删除线, null 表示清除。 |

### 示例代码

```
Workbook.getRange('A1').setVerticalAlignments('top');
```

## getVerticalAlignment()

获取左上角单元格垂直对齐方式。

### 返回结果

'top' | 'middle' | 'bottom' - 垂直对齐方式。

### 示例代码

```
const alignment = Workbook.getRange('A1').getVerticalAlignment();
```

## getVerticalAlignments()

获取区域单元格垂直对齐方式。。

### 返回结果

Array<Array<'top' | 'middle' | 'bottom'>> - 垂直对齐方式。

### 示例代码

```
const alignments = Workbook.getRange('A1').getVerticalAlignments();
```

## setWordWrap(wordWrap)

设置换行方式。

### 返回结果

Range - 区域。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| wordWrap | 'overflow' | 'clip' | 'autoWrap' | null | 是 | 换行方式, null 表示清除。 |

### 示例代码

```
Workbook.getRange('A1').setWordWrap('overflow');
```

## getWordWrap()

获取左上角单元格换行方式。

### 返回结果

'overflow' | 'clip' | 'autoWrap' - 换行方式。

### 示例代码

```
const wordWrap = Workbook.getRange('A1').getWordWrap();
```

## getWordWraps()

获取区域单元格换行方式。

### 返回结果

Array<Array<'overflow' | 'clip' | 'autoWrap'>> - 换行方式。

### 示例代码

```
const wordWraps = Workbook.getRange('A1').getWordWraps();
```

## setBorder(null)

清除边框。

### 返回结果

Range - 区域。

### 示例代码

```
workbook.getActiveRange().setBorder(null);
```

## setBorder(type, color, style)

设置边框。

### 返回结果

Range - 区域。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| type | BorderType | 是 | 边框类型。 |
| color | string | 是 | 边框颜色。 |
| style | 'none' |  'dotted'|  'dashed'|  'solid'|  'medium'|  'thick'|  'double'|  'hair'|  'dashDotDot'|  'dashDot'|  'mediumDashDotDot'|  'slantDashDot'|  'mediumDashDot'|  'mediumDashed' | 是 | 边框样式。 |

### 示例代码

```
Workbook.getActiveRange().setBorder({ top: true, bottom: true },  '#ff00ff', 'solid');
```

## decreaseDecimal()

减少小数位，如果 activeCell 在 Range 内，以 activeCell 为参照，否则以左上角单元格为参照。

### 返回结果

Range - 区域。

### 示例代码

```
Workbook.getRange('A1').decreaseDecimal();
```

## increaseDecimal()

减少小数位，如果 activeCell 在 Range 内，以 activeCell 为参照，否则以左上角单元格为参照。

### 返回结果

Range - 区域。

### 示例代码

```
Workbook.getRange('A1').increaseDecimal();
```

## setNumberFormat(numberFormat)

设置数字格式。

### 返回结果

Range - 区域。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| numberFormat | string | 是 | 数字格式，详情参考[数字格式](/document/orgapp-client/number-format#topic-2193105)。 |

### 示例代码

```
Workbook.getActiveRange().setNumberFormat('0.00');
```

## getNumberFormat()

获取左上角单元格数字格式。

### 返回结果

string - 数字格式。

### 示例代码

```
const numberFormat = Workbook.getActiveRange().getNumberFormat();
```

## getNumberFormats()

获取区域单元格数字格式。

### 返回结果

Array<Array<string>> - 数字格式。

### 示例代码

```
const numberFormats = Workbook.getActiveRange().getNumberFormats();
```

## getFormula()

获取左上角单元格公式表达式。

### 返回结果

string - 公式表达式，无公式为空字符串。

### 示例代码

```
const formula = Workbook.getActiveRange().getFormula();
```

## getFormulas()

获取区域单元格公式表达式。

### 返回结果

Array<Array<string>> - 公式表达式，无公式为空字符串。

### 示例代码

```
const formulas = Workbook.getActiveRange().getFormulas();
```

## getDisplayValue()

获取左上角单元格显示值。显示值指基于数字格式等计算后的值。

### 返回结果

string - 单元格显示值。

### 示例代码

```
const value = Workbook.getActiveRange().getDisplayValue();
```

## getDisplayValues()

获取区域单元格显示值。显示值指基于数字格式等计算后的值。

### 返回结果

Array<Array<string>> - 单元格显示值。

### 示例代码

```
const formulas = Workbook.getActiveRange().getFormulas();
```

## getValue()

获取左上角单元格值。

### 返回结果

string | number | boolean | null - 单元格值。

### 示例代码

```
const value = Workbook.getActiveRange().getValue();
```

## getDisplayValues()

获取区域单元格值。

### 返回结果

Array<Array<string | number | boolean | null>> - 单元格值。

### 示例代码

```
const values = Workbook.getActiveRange().getValues();
```

## setValues(values, options)

设置值。

### 返回结果

Range - 区域。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| value | Array<Array<string | boolean | number | null>> | 是 | 值，传null 表示清除。 |
| options | SetValueOptions | 否 | 设置值的配置选项。 |

### 示例代码

```
Workbook.getRange('A1:B1').setValues([
  [1, 2],
]);

Workbook.getRange('A1:B1').setValues([
  [1, 2],
], { parseType: 'raw' });
```

## merge()

合并单元格。

### 返回结果

Range - 区域。

### 示例代码

```
Workbook.getActiveRange().merge();
```

## unmerge()

取消合并单元格。

### 返回结果

Range - 区域。

### 示例代码

```
Workbook.getActiveRange().unmerge();
```

## getMergedRanges()

获取合并单元格区域。

### 返回结果

Range[] - 合并单元格区域。

### 示例代码

```
const ranges = Workbook.getActiveRange().getMergedRanges();
```

## getA1Notation()

获取区域地址。

### 返回结果

string - 区域字符串地址。

### 示例代码

```
const notation = Workbook.getActiveRange().getA1Notation();
```

## getRow()

获取从0开始的行号。

### 返回结果

number - 行。

### 示例代码

```
const row = Workbook.getActiveRange().getRow();
```

## getColumn()

获取从0开始的列号。

### 返回结果

number - 列。

### 示例代码

```
const column = Workbook.getActiveRange().getColumn();
```

## getRowCount()

获取行数。

### 返回结果

number - 行数。

### 示例代码

```
const count = Workbook.getActiveRange().getRowCount();
```

## getColumnCount()

获取列数。

### 返回结果

number - 列数。

### 示例代码

```
const count = Workbook.getActiveRange().getColumnCount();
```

## getHyperlink()

获取单元格链接。

### 返回结果

Hyperlink | null - 单元格链接, null 表示无链接。

### 示例代码

```
const link = Workbook.getActiveRange().getHyperlink();
```

## sort(field)

排序当前区域。

### 返回结果

Range - 区域。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| field | SortField | number | 是 | 排序规则 ｜ 相对于当前 Range 首列的偏移量。 |

### 示例代码

```
// 基于当前选区中的第0列进行逆序排序
Workbook.getActiveRange().sort({  column: 0,  ascending: false });

// 基于当前选区中的第0列进行顺序排序
Workbook.getActiveRange().sort(0);
```

## find(text, findOptions)

根据指定的条件查找匹配给定字符串的第一个Range。

### 返回结果

Range - 区域。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| text | string | 是 | 要查找的字符串。 |
| findOptions | SearchOptions | 否 | 其他搜索条件，包括搜索是否需要匹配整个单元格、区分大小写、是否使用正则表达式、是否匹配公式。 |

### 示例代码

```
const range = Workbook.getRange('A1:C10');
const findRange1 = range.find('1');
const findRange2 = range.find('1', {  matchEntireCell: true });
```

## findNext(text, findOptions)

根据指定的条件查找匹配给定字符串下一个Range，搜索范围为以当前Range左上角的Cell之后开始的整个工作表。

### 返回结果

Range - 区域。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| text | string | 是 | 要查找的字符串。 |
| findOptions | SearchOptions | 否 | 其他搜索条件，包括搜索是否需要匹配整个单元格、区分大小写、是否使用正则、是否匹配公式。 |

### 示例代码

```
const range = Workbook.getRange('A2');
const findRange = range.findNext('1');
```

## findPrevious(text, findOptions)

根据指定的条件查找匹配给定字符串上一个Range，搜索范围为以当前Range左上角的Cell之前的整个工作表。

### 返回结果

Range - 区域。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| text | string | 是 | 要查找的字符串。 |
| findOptions | SearchOptions | 否 | 其他搜索条件，包括搜索是否需要匹配整个单元格、区分大小写、是否使用正则、是否匹配公式。 |

### 示例代码

```
const range = Workbook.getRange('A2');
const findRange = range.findPrevious('1');
```

## replaceAll(text, replaceText, replaceOptions)

根据指定的条件查找并替换给定的字符串。

### 返回结果

number - 执行的替换数，可能与匹配单元格的数目不同。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| text | string | 是 | 要查找的字符串。 |
| replaceText | string | 是 | 替换原始字符串的字符串。 |
| replaceOptions | SearchOptions |  | 其他替换条件，包括是否需要匹配整个单元格、区分大小写、是否使用正则、是否匹配公式。 |

### 示例代码

```
const range = Workbook.getRange('A1:C10');
const replaceCount = range.replaceAll('str', 'otherStr');
```

## insertCheckboxes()

插入复选框。

### 返回结果

Range - Range。

### 示例代码

```
const range = Workbook.getRange(2, 2, 2, 2);
range.insertCheckboxes();
```

## deleteCheckboxes()

清除复选框，value 不会清除。

### 返回结果

Range - Range。

### 示例代码

```
const range = Workbook.getRange(2, 2, 2, 2);
range.deleteCheckboxes();
```

## getCheckedState()

返回左上角单元格是否勾选，null表示单元格不是复选框。

### 返回结果

boolean | null - 是否勾选。

### 示例代码

```
const state = Workbook.getRange(2, 2, 2, 2).getCheckedState();
```

## getCheckedStates()

返回区域单元格是否勾选，null表示单元格不是复选框。

### 返回结果

Array<Array<boolean | null>> - 区域单元格勾选的状态。

### 示例代码

```
const states = Workbook.getRange(2, 2, 2, 2).getCheckedStates();
```

## setCheckedState(checked)

勾选/取消勾选复选框。

### 返回结果

Range - Range。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| checked | boolean | 是 | 是否勾选。 |

### 示例代码

```
Workbook.getRange(2, 2, 2, 2).setCheckedState(true);
```

## setCheckedStates(checkeds)

勾选/取消勾选复选框。

### 返回结果

Range - Range。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| checkeds | Array<Array<boolean | null>> | 是 | 是否勾选，null 表示跳过。 |

### 示例代码

```
Workbook.getRange(2, 2, 2, 2).setCheckedStates([[true, false], [false, null]]);
```

## insertDropdownLists(options)

设置下拉框。

### 返回结果

Range - Range。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| options | DropdownListOption[] | 是 | 下拉选项。 |

### 示例代码

```
const range = Workbook.getActiveRange();
range.insertDropdownLists([
    {
        value: '篮球',
        color: '#ff0000',
    },
    {
        value: '足球',
        color: '#00ff00',
    },
    {
        value: '排球',
        color: '#0000ff',
    },
])
```

## deleteDropdownLists()

清除下拉框，value 不会清除。

### 返回结果

Range - Range。

### 示例代码

```
Workbook.getActiveRange().deleteDropdownLists();
```