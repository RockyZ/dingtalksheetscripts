# FilterCriteriaBuilder

## copy()

忙聽鹿忙聧庐氓陆聯氓聣聧 FilterCriteriaBuilder 莽聰聼忙聢聬忙聳掳莽職聞 FilterCriteriaBuilder茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

FilterCriteriaBuilder - 忙聳掳莽職聞FilterCriteriaBuilder茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const builder = Workbook.newFilterCriteriaBuilder();
builder.setVisibleValues(['A', 'B', 'C']);
const copyBuilder = builder.copy();
filter.setColumnFilterCriteria(0, copyBuilder.build());
```

## setVisibleValues(values)

忙聦聣氓聙录莽颅聸茅聙聣-猫庐戮莽陆庐氓聫炉猫搂聛氓聙录茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

FilterCriteriaBuilder - 氓陆聯氓聣聧FilterCriteriaBuilder氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| values | string[] | 忙聵炉 | 氓聫炉猫搂聛氓聙录茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const builder = Workbook.newFilterCriteriaBuilder();
builder.setVisibleValues(['A', 'B', 'C']);
const criteria = builder.build();
filter.setColumnFilterCriteria(0, criteria);
```

## setVisibleBackgroundColor(visibleBackgroundColor)

忙聦聣氓聧聲氓聟聝忙聽录茅垄聹猫聣虏莽颅聸茅聙聣茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

FilterCriteriaBuilder - 氓陆聯氓聣聧FilterCriteriaBuilder氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| visibleBackgroundColor | Color | 忙聵炉 | 氓聧聲氓聟聝忙聽录茅垄聹猫聣虏茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const color = sheet.getRange('C3').getBackgroundColor();
const criteria = Workbook.newFilterCriteriaBuilder().
 setVisibleBackgroundColor(color).
  build();
filter.setColumnFilterCriteria(0, criteria);
```

## setVisibleFontColor(visibleFontColor)

忙聦聣氓聧聲氓聟聝忙聽录茅垄聹猫聣虏莽颅聸茅聙聣茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

FilterCriteriaBuilder - 氓陆聯氓聣聧FilterCriteriaBuilder氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| visibleFontColor | Color | 忙聵炉 | 氓颅聴盲陆聯茅垄聹猫聣虏茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const color = sheet.getRange('C3').getFontColor();
const criteria = workbook.newFilterCriteriaBuilder().
 setVisibleFontColor(color).
  build();
filter.setColumnFilterCriteria(0, criteria);
```

## setVisibleConditions(condition1, condition2, operator)

忙聦聣忙聺隆盲禄露莽颅聸茅聙聣茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

FilterCriteriaBuilder - 氓陆聯氓聣聧FilterCriteriaBuilder氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| condition1 | FilterCondition | 忙聵炉 | 莽颅聸茅聙聣忙聺隆盲禄露 |
| condition2 | FilterCondition | 氓聬娄 | 莽颅聸茅聙聣忙聺隆盲禄露 |
| operator | 'and' | 'or' | 氓聬娄 | 盲赂陇盲赂陋莽颅聸茅聙聣忙聺隆盲禄露莽職聞氓聟鲁莽鲁禄 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

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

猫聨路氓聫聳猫庐戮莽陆庐莽職聞氓聫炉猫搂聛氓聙录茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

string[] | null - 氓聫炉猫搂聛氓聙录茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
builder.getVisibleValues();
```

## getVisibleFontColor()

猫聨路氓聫聳猫庐戮莽陆庐莽職聞氓聫炉猫搂聛氓颅聴盲陆聯茅垄聹猫聣虏茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Color 茂陆聹 null - 氓聫炉猫搂聛氓颅聴盲陆聯茅垄聹猫聣虏茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
builder.getVisibleFontColor();
```

## getVisibleBackgroundColor()

猫聨路氓聫聳猫庐戮莽陆庐莽職聞氓聫炉猫搂聛氓聧聲氓聟聝忙聽录茅垄聹猫聣虏茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Color 茂陆聹 null - 氓聫炉猫搂聛氓聧聲氓聟聝忙聽录茅垄聹猫聣虏茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
builder.getVisibleBackgroundColor();
```

## getVisibleConditions()

猫聨路氓聫聳猫庐戮莽陆庐莽職聞莽颅聸茅聙聣忙聺隆盲禄露茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

FilterCondition[] | null - 莽颅聸茅聙聣忙聺隆盲禄露氓聢聴猫隆篓茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
builder.getVisibleConditions();
```

## getVisibleConditionsOperator()

猫聨路氓聫聳猫庐戮莽陆庐莽職聞盲赂陇盲赂陋莽颅聸茅聙聣忙聺隆盲禄露氓聟鲁莽鲁禄茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

'and' | 'or' | null - 莽颅聸茅聙聣忙聺隆盲禄露氓聟鲁莽鲁禄茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
builder.getVisibleConditionsOperator();
```

## build()

莽聰聼忙聢聬莽颅聸茅聙聣猫搂聞氓聢聶茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

FilterCriteria - 莽颅聸茅聙聣猫搂聞氓聢聶茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
workbook.newFilterCriteriaBuilder().
 setVisibleValues(['A', 'B', 'C']).
  build();
```