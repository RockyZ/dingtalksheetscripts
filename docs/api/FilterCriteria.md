# FilterCriteria

## copy()

根据当前 FilterCriteria 生成新的 FilterCriteriaBuilder。

### 返回结果

FilterCriteriaBuilder - 新的FilterCriteriaBuilder。

### 示例代码

```
const builder = criteria.copy();
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