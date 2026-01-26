# RangeList

## activate()

激活多选区。

### 返回结果

RangeList - 当前RangeList实例。

### 示例代码

```
const rangeList = Workbook.getRangeList(['A1:C10', 'D2:E5']);
rangeList.activate();
```

## clear()

清除样式和数据。

### 返回结果

RangeList - 当前RangeList实例。

### 示例代码

```
const rangeList = Workbook.getRangeList(['A1:C10', 'D2:E5']);
rangeList.clear();
```

## clearFormat()

清除样式。

### 返回结果

RangeList - 当前RangeList实例。

### 示例代码

```
const rangeList = Workbook.getRangeList(['A1:C10', 'D2:E5']);
rangeList.clearFormat();
```

## clearData()

清除数据。

### 返回结果

RangeList - 当前RangeList实例。

### 示例代码

```
const rangeList = Workbook.getRangeList(['A1:C10', 'D2:E5']);
rangeList.clearData();
```

## setBackgroundColor(backgroundColor)

设置背景颜色。

### 返回结果

RangeList - 当前RangeList实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| backgroundColor | string | null | 是 | 背景颜色, null 表示清除。 |

### 示例代码

```
const rangeList = Workbook.getRangeList(['A1:C10', 'D2:E5']);
rangeList.setBackgroundColor('red');
```

## setFontColor(fontColor)

设置字体颜色。

### 返回结果

RangeList - 当前RangeList实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| fontColor | string | null | 是 | 字体颜色, null 表示清除。 |

### 示例代码

```
const rangeList = Workbook.getRangeList(['A1:C10', 'D2:E5']);
rangeList.setFontColor('green');
```

## setFontSize(fontSize)

设置字体大小。

### 返回结果

RangeList - 当前RangeList实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| fontSize | number | null | 是 | 字体大小, null 表示清除。 |

### 示例代码

```
const rangeList = Workbook.getRangeList(['A1:C10', 'D2:E5']);
rangeList.setFontSize(30);
```

## setFontFamily(fontFamily)

设置字体。

### 返回结果

RangeList - 当前RangeList实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| fontFamily | string | null | 是 | 字体, null 表示清除。 |

### 示例代码

```
const rangeList = Workbook.getRangeList(['A1:C10', 'D2:E5']);
rangeList.setFontFamily('impact');
```

## setFontWeight(fontWeight)

设置字体粗体。

### 返回结果

RangeList - 当前RangeList实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| fontWeight | 'bold' | 'normal' | null | 是 | 字体粗体, null 表示清除。 |

### 示例代码

```
const rangeList = Workbook.getRangeList(['A1:C10', 'D2:E5']);
rangeList.setFontWeight('bold');
```

## setFontStyle(fontStyle)

设置字体斜体。

### 返回结果

RangeList - 当前RangeList实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| fontStyle | 'italic' | 'normal' | null | 是 | 字体斜体, null 表示清除。 |

### 示例代码

```
const rangeList = Workbook.getRangeList(['A1:C10', 'D2:E5']);
rangeList.setFontStyle('italic');
```

## setIndent(indent)

设置缩进。

### 返回结果

RangeList - 当前RangeList实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| indent | number | null | 是 | 缩进，必须为正整数，null 表示清除。 |

### 示例代码

```
const rangeList = Workbook.getRangeList(['A1:C10', 'D2:E5']);
rangeList.setIndent(5);
```

## setTextLineThrough(lineThrough)

设置字体删除线。

### 返回结果

RangeList - 当前RangeList实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| lineThrough | 'lineThrough' | 'none' | null | 是 | 字体删除线, null 表示清除。 |

### 示例代码

```
const rangeList = Workbook.getRangeList(['A1:C10', 'D2:E5']);
rangeList.setTextLineThrough('lineThrough');
```

## setTextUnderline(textUnderline)

设置字体下划线。

### 返回结果

RangeList - 当前RangeList实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| textUnderline | 'single' | 'none' | null | 是 | 字体下划线, null 表示清除。 |

### 示例代码

```
const rangeList = Workbook.getRangeList(['A1:C10', 'D2:E5']);
rangeList.setTextUnderline('single');
```

## setHorizontalAlignment(alignment)

设置水平对齐方式。

### 返回结果

RangeList - 当前RangeList实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| horizontalAlignment | 'left' | 'center' | 'right' | 'general' | null | 是 | 水平对齐方式, null 表示清除。 |

### 示例代码

```
const rangeList = Workbook.getRangeList(['A1:C10', 'D2:E5']);
rangeList.setHorizontalAlignment('left');
```

## setVerticalAlignment(alignment)

设置垂直对齐方式。

### 返回结果

RangeList - 当前RangeList实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| verticalAlignment | 'top' | 'middle' | 'bottom' | null | 是 | 垂直对齐方式, null 表示清除。 |

### 示例代码

```
const rangeList = Workbook.getRangeList(['A1:C10', 'D2:E5']);
rangeList.setVerticalAlignment('top');
```

## setWordWrap(wordWrap)

设置换行方式。

### 返回结果

RangeList - 当前RangeList实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| wordWrap | 'overflow' | 'clip' | 'autoWrap' | null | 是 | 换行方式, null 表示清除。 |

### 示例代码

```
const rangeList = Workbook.getRangeList(['A1:C10', 'D2:E5']);
rangeList.setWordWrap('overflow');
```

## setNumberFormat(numberFormat)

设置数字格式。

### 返回结果

RangeList - 当前RangeList实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| numberFormat | string | 是 | 数字格式。 |

### 示例代码

```
const rangeList = Workbook.getRangeList(['A1:C10', 'D2:E5']);
rangeList.setNumberFormat('0.00');
```

## merge()

合并单元格。

### 返回结果

RangeList - 当前RangeList实例。

### 示例代码

```
const rangeList = Workbook.getRangeList(['A1:C10', 'D2:E5']);
rangeList.merge();
```

## getRanges()

获取当前多选区中包含的所有选区。

### 返回结果

Range[] - 所有选区。

### 示例代码

```
const rangeList = Workbook.getRangeList(['A1:C10', 'D2:E5']);
const ranges = rangeList.getRanges();
```