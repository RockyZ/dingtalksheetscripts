# FilterCriteriaBuilder

## copy()

根据当前 FilterCriteriaBuilder 生成新的 FilterCriteriaBuilder。

### 返回结果

FilterCriteriaBuilder - 新的FilterCriteriaBuilder。

### 示例代码

```
const builder = Workbook.newFilterCriteriaBuilder();
builder.setVisibleValues(['A', 'B', 'C']);
const copyBuilder = builder.copy();
filter.setColumnFilterCriteria(0, copyBuilder.build());
```

## setVisibleValues(values)

按值筛选-设置可见值。

### 返回结果

FilterCriteriaBuilder - 当前FilterCriteriaBuilder实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| values | string[] | 是 | 可见值。 |

### 示例代码

```
const builder = Workbook.newFilterCriteriaBuilder();
builder.setVisibleValues(['A', 'B', 'C']);
const criteria = builder.build();
filter.setColumnFilterCriteria(0, criteria);
```

## setVisibleBackgroundColor(visibleBackgroundColor)

按单元格颜色筛选。

### 返回结果

FilterCriteriaBuilder - 当前FilterCriteriaBuilder实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| visibleBackgroundColor | Color | 是 | 单元格颜色。 |

### 示例代码

```
const color = sheet.getRange('C3').getBackgroundColor();
const criteria = Workbook.newFilterCriteriaBuilder().
 setVisibleBackgroundColor(color).
  build();
filter.setColumnFilterCriteria(0, criteria);
```

## setVisibleFontColor(visibleFontColor)

按单元格颜色筛选。

### 返回结果

FilterCriteriaBuilder - 当前FilterCriteriaBuilder实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| visibleFontColor | Color | 是 | 字体颜色。 |

### 示例代码

```
const color = sheet.getRange('C3').getFontColor();
const criteria = workbook.newFilterCriteriaBuilder().
 setVisibleFontColor(color).
  build();
filter.setColumnFilterCriteria(0, criteria);
```

## setVisibleConditions(condition1, condition2, operator)

按条件筛选。

### 返回结果

FilterCriteriaBuilder - 当前FilterCriteriaBuilder实例。

### 参数说明

| 参数 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| condition1 | FilterCondition | 是 | 筛选条件 |
| condition2 | FilterCondition | 否 | 筛选条件 |
| operator | 'and' | 'or' | 否 | 两个筛选条件的关系 |

### 示例代码

```
const filterCondition1 = {
 operator: 'equal',
  value: 'C3',
};
const criteria = Workbook.newFilterCriteriaBuilder().
 setVisibleConditions(filterCondition1).
  build();
const filterCondition2 = {
 operator: 'contains',
  value: '4',
  };
const criteria2 = Workbook.newFilterCriteriaBuilder().
 setVisibleConditions(filterCondition1, filterCondition2, 'or').
  build();
```

## getVisibleValues()

获取设置的可见值。

### 返回结果

string[] | null - 可见值。

### 示例代码

```
builder.getVisibleValues();
```

## getVisibleFontColor()

获取设置的可见字体颜色。

### 返回结果

Color ｜ null - 可见字体颜色。

### 示例代码

```
builder.getVisibleFontColor();
```

## getVisibleBackgroundColor()

获取设置的可见单元格颜色。

### 返回结果

Color ｜ null - 可见单元格颜色。

### 示例代码

```
builder.getVisibleBackgroundColor();
```

## getVisibleConditions()

获取设置的筛选条件。

### 返回结果

FilterCondition[] | null - 筛选条件列表。

### 示例代码

```
builder.getVisibleConditions();
```

## getVisibleConditionsOperator()

获取设置的两个筛选条件关系。

### 返回结果

'and' | 'or' | null - 筛选条件关系。

### 示例代码

```
builder.getVisibleConditionsOperator();
```

## build()

生成筛选规则。

### 返回结果

FilterCriteria - 筛选规则。

### 示例代码

```
workbook.newFilterCriteriaBuilder().
 setVisibleValues(['A', 'B', 'C']).
  build();
```