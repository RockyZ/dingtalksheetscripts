# Filter

## getId()

猫聨路氓聫聳 ID茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

string - ID茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const filterId = filter.getId();
```

## getRange()

猫聨路氓聫聳氓陆聯氓聣聧莽颅聸茅聙聣氓聦潞氓聼聼茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range | null - 莽颅聸茅聙聣忙聣聙氓聹篓莽職聞 Range茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const range = filter.getRange();
```

## delete()

莽搂禄茅聶陇氓陆聯氓聣聧 filter茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

void -

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
filter.delete();
```

## getColumnFilterCriteria(column)

猫聨路氓聫聳忙聦聡氓庐職氓聢聴莽職聞莽颅聸茅聙聣猫搂聞氓聢聶氓庐聻盲戮聥茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

FilterCriteria | null - 莽颅聸茅聙聣猫搂聞氓聢聶茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| column | number | 忙聵炉 | 莽聸赂氓炉鹿盲潞聨 filterRange 茅娄聳氓聢聴莽職聞氓聛聫莽搂禄茅聡聫茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const filterCriteria = filter.getColumnFilterCriteria(0);
```

## setColumnFilterCriteria(column, filterCriteria)

猫庐戮莽陆庐忙聦聡氓庐職氓聢聴莽職聞莽颅聸茅聙聣猫搂聞氓聢聶茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Filter - 氓陆聯氓聣聧Filter氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| column | number | 忙聵炉 | 莽聸赂氓炉鹿盲潞聨莽颅聸茅聙聣氓聦潞氓聼聼茅娄聳氓聢聴莽職聞氓聛聫莽搂禄茅聡聫茫聙聜 |
| filterCriteria | FilterCriteria | 忙聵炉 | 莽颅聸茅聙聣猫搂聞氓聢聶茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const criteria = Workbook.newFilterCriteriaBuilder().
 setVisibleValues(['A', 'B', 'C']).
  build();
filter.setColumnFilterCriteria(0, criteria);
```

## clearColumnFilterCriteria(column)

莽搂禄茅聶陇忙聦聡氓庐職氓聢聴莽職聞莽颅聸茅聙聣猫搂聞氓聢聶茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Filter - 氓陆聯氓聣聧Filter氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| column | number | 忙聵炉 | 莽聸赂氓炉鹿盲潞聨莽颅聸茅聙聣氓聦潞氓聼聼茅娄聳氓聢聴莽職聞氓聛聫莽搂禄茅聡聫茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
filter.clearColumnFilterCriteria(0);
```

## sort(field)

莽颅聸茅聙聣忙聨聮氓潞聫茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Filter - 氓陆聯氓聣聧Filter氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 猫炉麓忙聵聨 |
| --- | --- | --- |
| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 猫炉麓忙聵聨 |
| field | SortField | number    莽卤禄氓聻聥忙聵炉number忙聴露茂录聦忙聦聡氓庐職忙聨聮氓潞聫忙聦聣莽聟搂莽颅聸茅聙聣氓聦潞氓聼聼莽職聞茅娄聳氓聢聴莽職聞氓聛聫莽搂禄茅聡聫茂录聦茅聶聧氓潞聫忙聨聮氓潞聫 | 忙聨聮氓潞聫猫搂聞氓聢聶茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
filter.sort({ column: 0, ascending: false });
filter.sort(0);
```

## getColumnSortState(column)

猫聨路氓聫聳忙聨聮氓潞聫莽聤露忙聙聛茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

'none' | 'ascending' | 'descending' - 忙聨聮氓潞聫莽聤露忙聙聛茂录聦盲戮聺忙卢隆盲禄拢猫隆篓忙聴聽忙聨聮氓潞聫茫聙聛氓聧聡氓潞聫氓聮聦茅聶聧氓潞聫茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 猫炉麓忙聵聨 |
| --- | --- | --- |
| column | number | 莽聸赂氓炉鹿盲潞聨莽颅聸茅聙聣氓聦潞氓聼聼茅娄聳氓聢聴莽職聞氓聛聫莽搂禄茅聡聫茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const state = filter.getColumnSortState(0);
```