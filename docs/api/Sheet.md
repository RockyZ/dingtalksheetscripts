# Sheet

## activate()

忙驴聙忙麓禄猫聡陋猫潞芦茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧Sheet氓庐聻盲戮聥茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getSheet('Sheet1').activate();
```

## getId()

猫聨路氓聫聳ID茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

string - 氓路楼盲陆聹猫隆篓莽職聞ID茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const sheetId = Workbook.getActiveSheet().getId();
```

## getIndex()

猫聨路氓聫聳氓路楼盲陆聹猫隆篓莽職聞莽麓垄氓录聲茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

number - 氓路楼盲陆聹猫隆篓莽職聞莽麓垄氓录聲茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const index = Workbook.getActiveSheet().getIndex();
```

## getName()

猫聨路氓聫聳氓路楼盲陆聹猫隆篓莽職聞氓聬聧氓颅聴茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

string - 氓路楼盲陆聹猫隆篓莽職聞氓聬聧氓颅聴茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const name = Workbook.getActiveSheet().getName();
```

## setTabColor(color)

猫庐戮莽陆庐Sheet忙聽聡莽颅戮茅垄聹猫聣虏茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧Sheet氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

|  |  |  |  |
| --- | --- | --- | --- |
| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| color | string | null | 忙聵炉 | 茅垄聹猫聣虏茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getActiveSheet().setTabColor('yellow');
```

## getTabColor()

猫聨路氓聫聳Sheet忙聽聡莽颅戮茅垄聹猫聣虏茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Color | null - 茅垄聹猫聣虏茫聙聜null盲禄拢猫隆篓忙虏隆忙聹聣猫庐戮莽陆庐茅垄聹猫聣虏茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const tabColor = Workbook.getActiveSheet().getTabColor();
```

## setName(name)

猫庐戮莽陆庐氓路楼盲陆聹猫隆篓莽職聞氓聬聧氓颅聴茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧Sheet氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

|  |  |  |  |
| --- | --- | --- | --- |
| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| name | string | 忙聵炉 | 忙聳掳氓聬聧氓颅聴茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getActiveSheet().setName('New name');
```

## getVisibility()

猫聨路氓聫聳氓聫炉猫搂聛忙聙搂茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

'visible' | 'hidden' - 氓聫炉猫搂聛忙聙搂茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const visibility = Workbook.getActiveSheet().getVisibility();
```

## getGridlinesVisibility()

猫聨路氓聫聳莽陆聭忙聽录莽潞驴氓聫炉猫搂聛忙聙搂茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

'visible' | 'hidden' - 氓聫炉猫搂聛忙聙搂茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const gridlinesVisibility = Workbook.getActiveSheet().getGridlinesVisibility();
```

## setGridlinesVisibility(visibility)

猫庐戮莽陆庐莽陆聭忙聽录莽潞驴氓聫炉猫搂聛忙聙搂茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧Sheet氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳莽卤禄氓聻聥

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| visibility | 'visible' | 'hidden' | 忙聵炉 | 氓聫炉猫搂聛忙聙搂茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getActiveSheet().setGridlinesVisibility('hidden');
```

## getFrozenColumnCount()

猫聨路氓聫聳氓聠禄莽禄聯莽職聞氓聢聴茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

number - 氓聢聴忙聲掳茂录聦盲禄聨1氓录聙氓搂聥茂录聦0猫隆篓莽陇潞盲赂聧氓聠禄莽禄聯茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const frozenColumnCount = Workbook.getActiveSheet().getFrozenColumnCount();
```

## getFrozenRowCount()

猫聨路氓聫聳氓聠禄莽禄聯莽職聞猫隆聦茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

number - 猫隆聦忙聲掳茂录聦盲禄聨1氓录聙氓搂聥茂录聦0猫隆篓莽陇潞盲赂聧氓聠禄莽禄聯茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const frozenRowCount = Workbook.getActiveSheet().getFrozenRowCount();
```

## setFrozenColumnCount(columnCount)

猫庐戮莽陆庐氓聠禄莽禄聯氓聢聴茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧Sheet氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| columnCount | number | 忙聵炉 | 氓聢聴忙聲掳茂录聦盲禄聨 1 氓录聙氓搂聥茂录聦0猫隆篓莽陇潞盲赂聧氓聠禄莽禄聯茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getActiveSheet().setFrozenColumnCount(5);
```

## setFrozenRowCount(rowCount)

猫庐戮莽陆庐氓聠禄莽禄聯猫隆聦茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧Sheet氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳莽卤禄氓聻聥

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| rowCount | number | 忙聵炉 | 猫隆聦忙聲掳茂录聦盲禄聨 1 氓录聙氓搂聥茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getActiveSheet().setFrozenRowCount(5);
```

## getColumnCount()

猫聨路氓聫聳氓陆聯氓聣聧Sheet莽職聞氓聢聴忙聲掳茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

number - 氓聢聴忙聲掳茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const columnCount = Workbook.getActiveSheet().getColumnCount();
```

## getRowCount()

猫聨路氓聫聳氓陆聯氓聣聧Sheet莽職聞猫隆聦忙聲掳茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

number - 猫隆聦忙聲掳茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const rowCount = Workbook.getActiveSheet().getRowCount();
```

## setActiveRange(a1Notation)

忙驴聙忙麓禄茅聙聣氓聦潞茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 忙驴聙忙麓禄莽職聞茅聙聣氓聦潞茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getActiveSheet().setActiveRange("A1:C10");
```

## setActiveRange(range)

忙驴聙忙麓禄茅聙聣氓聦潞茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 忙驴聙忙麓禄莽職聞茅聙聣氓聦潞茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| range | Range | 忙聵炉 | 茅聙聣氓聦潞茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const range = Workbook.getRange(0, 0, 2, 2);
Workbook.getActiveSheet().setActiveRange(range);
```

## getActiveRange()

猫聨路氓聫聳 activeCell 忙聣聙氓聹篓莽職聞茅聙聣氓聦潞茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range | null - activeCell 忙聣聙氓聹篓莽職聞茅聙聣氓聦潞

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const range = Workbook.getActiveSheet().getActiveRange();
```

## getEntireSheet()

猫聨路氓聫聳忙聲麓猫隆篓氓聦潞氓聼聼茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 忙聲麓猫隆篓氓聦潞氓聼聼

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const range = Workbook.getActiveSheet().getEntireSheet();
```

## getEntireRows(row, rowCount)

猫聨路氓聫聳忙聲麓猫隆聦茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range | null - 忙聲麓猫隆聦氓聦潞氓聼聼

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| row | number | 忙聵炉 | 猫碌路氓搂聥猫隆聦茫聙聜 |
| rowCount | number | 氓聬娄 | 猫隆聦忙聲掳茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const range = Workbook.getActiveSheet().getEntireRows(0, 3);
```

## getEntireColumns(column, columnCount)

猫聨路氓聫聳忙聲麓氓聢聴茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range | null - 忙聲麓氓聢聴氓聦潞氓聼聼

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| column | number | 忙聵炉 | 猫碌路氓搂聥氓聢聴茫聙聜 |
| columnCount | number | 氓聬娄 | 氓聢聴忙聲掳茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const range = Workbook.getActiveSheet().getEntireColumns(0, 3);
```

## getRange(a1Notation)

猫聨路氓聫聳茅聙聣氓聦潞猫聦聝氓聸麓茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 茅聙聣氓聦潞

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| a1Notation | string | 忙聵炉 | 茅聙聣氓聦潞氓聹掳氓聺聙茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const range = Workbook.getActiveSheet().getRange('A1:C10');
```

## getRange(row, column, rowCount, columnCount)

猫聨路氓聫聳茅聙聣氓聦潞猫聦聝氓聸麓茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 茅聙聣氓聦潞

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| row | number | 忙聵炉 | 猫隆聦茂录聦盲禄聨 0 氓录聙氓搂聥茫聙聜 |
| column | number | 忙聵炉 | 氓聢聴茂录聦盲禄聨 0 氓录聙氓搂聥茫聙聜 |
| rowCount | number | 忙聵炉 | 猫隆聦忙聲掳茂录聦盲禄聨 1 氓录聙氓搂聥茫聙聜 |
| columnCount | number | 忙聵炉 | 氓聢聴忙聲掳茂录聦盲禄聨 1 氓录聙氓搂聥茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const range = Workbook.getActiveSheet().getRange(0, 0, 2, 2);
```

## setActiveRangeList(a1Notations)

忙驴聙忙麓禄氓陇職茅聙聣氓聦潞茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

RangeList - 氓陇職茅聙聣氓聦潞

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 |
| --- | --- |
| a1Notations | string[] |
| rangeList | RangeList |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const rangeList = Workbook.getActiveSheet().setActiveRangeList(['A1:C10', 'D2:E5']);
```

## setActiveRangeList(rangeList)

忙驴聙忙麓禄氓陇職茅聙聣氓聦潞茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

RangeList - 氓陇職茅聙聣氓聦潞

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 |
| --- | --- |
| rangeList | RangeList |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const rangeList = Workbook.getActiveSheet().getRangeList(['A1:C10', 'D2:E5']);
Workbook.getActiveSheet().setActiveRangeList(rangeList);
```

## getActiveRangeList()

猫聨路氓聫聳忙驴聙忙麓禄莽職聞氓陇職茅聙聣氓聦潞茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

RangeList - 氓陇職茅聙聣氓聦潞

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const rangeList = Workbook.getActiveSheet().getActiveRangeList();
```

## getRangeList(a1Notations)

猫聨路氓聫聳氓陇職茅聙聣氓聦潞茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

RangeList - 氓陇職茅聙聣氓聦潞

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| a1Notations | string[] | 忙聵炉 | 茅聙聣氓聦潞氓聹掳氓聺聙忙聲掳莽禄聞茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const rangeList = Workbook.getActiveSheet().getRangeList(['A1:C10', 'D2:E5']);
```

## getActiveCell()

猫聨路氓聫聳active cell茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - active cell

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const range = Workbook.getActiveSheet().getActiveCell();
```

## getCell(row, column)

猫聨路氓聫聳氓聧聲氓聟聝忙聽录茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 氓聧聲氓聟聝忙聽录

### 氓聫聜忙聲掳猫炉麓忙聵聨

|  |  |  |  |
| --- | --- | --- | --- |
| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| row | number | 忙聵炉 | 猫隆聦茫聙聜 |
| column | number | 忙聵炉 | 氓聢聴茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const range = Workbook.getActiveSheet().getCell(0, 0);
```

## setRowsVisibility(row, rowCount, visibility)

忙聵戮莽陇潞/茅職聬猫聴聫氓陇職猫隆聦茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧Sheet氓庐聻盲戮聥

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| row | number | 忙聵炉 | 猫隆聦茫聙聜 |
| rowCount | number | 忙聵炉 | 猫隆聦忙聲掳茫聙聜 |
| visibility | 'visible' | 'hidden' | 忙聵炉 | 氓聫炉猫搂聛忙聙搂茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
// 茅職聬猫聴聫莽卢卢盲赂聙猫隆聦
Workbook.getActiveSheet().setRowsVisibility(0, 1, 'hidden');
```

## setRowVisibility(row, visibility)

忙聵戮莽陇潞/茅職聬猫聴聫氓聧聲猫隆聦茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧Sheet氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| row | number | 忙聵炉 | 猫隆聦茫聙聜 |
| visibility | 'visible' | 'hidden' | 忙聵炉 | 氓聫炉猫搂聛忙聙搂茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
// 茅職聬猫聴聫莽卢卢盲潞聦猫隆聦
Workbook.getActiveSheet().setRowVisibility(1, 'hidden');
```

## setColumnsVisibility(column, columnCount, visibility)

忙聵戮莽陇潞/茅職聬猫聴聫氓陇職氓聢聴茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧Sheet氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| column | number | 忙聵炉 | 氓聢聴茫聙聜 |
| columnCount | number | 忙聵炉 | 氓聢聴忙聲掳茫聙聜 |
| visibility | 'visible' | 'hidden' | 忙聵炉 | 氓聫炉猫搂聛忙聙搂茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
// 茅職聬猫聴聫莽卢卢盲潞聦猫隆聦
Workbook.getActiveSheet().setRowVisibility(1, 'hidden');
```

## setColumnVisibility(column, visibility)

忙聵戮莽陇潞/茅職聬猫聴聫氓聧聲氓聢聴茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧Sheet氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| column | number | 忙聵炉 | 氓聢聴茫聙聜 |
| visibility | 'visible' | 'hidden' | 忙聵炉 | 氓聫炉猫搂聛忙聙搂茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
// 茅職聬猫聴聫莽卢卢盲潞聦氓聢聴
Workbook.getActiveSheet().setColumnVisibility(1, 'hidden');
```

## setRowsHeight(row, rowCount, height)

猫庐戮莽陆庐氓陇職猫隆聦莽職聞茅芦聵氓潞娄茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧Sheet氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| row | number | 忙聵炉 | 猫隆聦茫聙聜 |
| rowCount | number | 忙聵炉 | 猫隆聦忙聲掳茫聙聜 |
| height | number | 忙聵炉 | 茅芦聵氓潞娄茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
// 猫庐戮莽陆庐莽卢卢盲赂聙猫隆聦氓聢掳莽卢卢盲赂聣猫隆聦茂录聦茅芦聵氓潞娄盲赂潞50
Workbook.getActiveSheet().setRowsHeight(0, 3, 50);
```

## setRowHeight(row, height)

猫庐戮莽陆庐氓聧聲猫隆聦莽職聞茅芦聵氓潞娄茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧Sheet氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| row | number | 忙聵炉 | 猫隆聦茫聙聜 |
| height | number | 忙聵炉 | 茅芦聵氓潞娄茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getActiveSheet().setRowHeight(0, 50);
```

## getRowHeight(row)

猫聨路氓聫聳猫隆聦茅芦聵茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧Sheet氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| row | number | 忙聵炉 | 猫隆聦茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const height = Workbook.getActiveSheet().getRowHeight(0);
```

## autofitRows(row, rowCount)

猫聡陋茅聙聜氓潞聰猫隆聦茅芦聵茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧Sheet氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| row | number | 忙聵炉 | 猫隆聦茫聙聜 |
| rowCount | number | 忙聵炉 | 猫隆聦忙聲掳茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getActiveSheet().autofitRows(0, 3);
```

## autofitRow(row)

猫聡陋茅聙聜氓潞聰猫隆聦茅芦聵茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧Sheet氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| row | number | 忙聵炉 | 猫隆聦茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getActiveSheet().autofitRow(0);
```

## setColumnsWidth(column, columnCount, width)

猫庐戮莽陆庐氓聢聴氓庐陆茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧Sheet氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| column | number | 忙聵炉 | 氓聢聴茫聙聜 |
| columnCount | number | 忙聵炉 | 氓聢聴忙聲掳茫聙聜 |
| width | number | 忙聵炉 | 氓庐陆氓潞娄茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
// 猫庐戮莽陆庐A氓聢掳C氓聢聴茂录聦氓庐陆氓潞娄盲赂潞200
Workbook.getActiveSheet().setColumnsWidth(0, 3, 200);
```

## setColumnWidth(column, width)

猫庐戮莽陆庐氓聢聴氓庐陆茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧Sheet氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| column | number | 忙聵炉 | 氓聢聴茫聙聜 |
| width | number | 忙聵炉 | 茅芦聵氓潞娄茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getActiveSheet().setColumnWidth(0, 200);
```

## getColumnWidth(column)

猫聨路氓聫聳氓聢聴氓庐陆茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

number - 氓聢聴氓庐陆茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| column | number | 忙聵炉 | 氓聢聴茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const width = Workbook.getActiveSheet().getColumnWidth(0);
```

## autofitColumns(column, columnCount)

猫聡陋茅聙聜氓潞聰氓聢聴氓庐陆茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

number - 氓聢聴氓庐陆茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| column | number | 忙聵炉 | 氓聢聴茫聙聜 |
| columnCount | number | 忙聵炉 | 氓聢聴茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getActiveSheet().autofitColumns(0, 3);
```

## autofitColumn(column)

猫聡陋茅聙聜氓潞聰氓聢聴氓庐陆茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

number - 氓聢聴氓庐陆茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| column | number | 忙聵炉 | 氓聢聴茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getActiveSheet().autofitColumn(0);
```

## insertRowBefore(row)

氓聹篓 row 猫隆聦盲鹿聥氓聣聧忙聫聮氓聟楼 1 猫隆聦茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧Sheet氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| row | number | 忙聵炉 | 猫隆聦茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getActiveSheet().insertRowBefore(0);
```

## insertRowsBefore(row, rowCount)

氓聹篓 row 猫隆聦盲鹿聥氓聣聧忙聫聮氓聟楼 n 猫隆聦茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧Sheet氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| row | number | 忙聵炉 | 猫隆聦茫聙聜 |
| rowCount | number | 忙聵炉 | 猫隆聦忙聲掳茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getActiveSheet().insertRowsBefore(0, 10);
```

## insertRowAfter(row)

氓聹篓 row 猫隆聦盲鹿聥氓聬聨忙聫聮氓聟楼 1 猫隆聦茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧Sheet氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

|  |  |  |  |
| --- | --- | --- | --- |
| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| row | number | 忙聵炉 | 猫隆聦茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getActiveSheet().insertRowAfter(0);
```

## insertRowsAfter(row, rowCount)

氓聹篓 row 猫隆聦盲鹿聥氓聬聨忙聫聮氓聟楼 n 猫隆聦茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧Sheet氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| row | number | 忙聵炉 | 猫隆聦茫聙聜 |
| rowCount | number | 忙聵炉 | 猫隆聦忙聲掳茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getActiveSheet().insertRowsAfter(0, 10);
```

## insertColumnBefore(column)

氓聹篓 column 猫隆聦盲鹿聥氓聣聧忙聫聮氓聟楼 1 氓聢聴茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧Sheet氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| column | number | 忙聵炉 | 氓聢聴茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getActiveSheet().insertColumnBefore(0);
```

## insertColumnsBefore(column, columnCount)

氓聹篓 column 猫隆聦盲鹿聥氓聣聧忙聫聮氓聟楼 n 氓聢聴茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧Sheet氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| column | number | 忙聵炉 | 氓聢聴茫聙聜 |
| columnCount | number | 忙聵炉 | 氓聢聴忙聲掳茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getActiveSheet().insertColumnsBefore(0, 10);
```

## insertColumnAfter(column)

氓聹篓 column 猫隆聦盲鹿聥氓聬聨忙聫聮氓聟楼 1 氓聢聴茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧Sheet氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| column | number | 忙聵炉 | 氓聢聴茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getActiveSheet().insertColumnAfter(0);
```

## insertColumnsAfter(column, columnCount)

氓聹篓 column 猫隆聦盲鹿聥氓聬聨忙聫聮氓聟楼 n 氓聢聴茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧Sheet氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| column | number | 忙聵炉 | 氓聢聴茫聙聜 |
| columnCount | number | 忙聵炉 | 氓聢聴忙聲掳茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getActiveSheet().insertColumnsAfter(0, 10);
```

## deleteRows(row, rowCount)

氓聢聽茅聶陇猫隆聦茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧Sheet氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| row | number | 忙聵炉 | 猫隆聦茫聙聜 |
| rowCount | number | 忙聵炉 | 猫隆聦忙聲掳茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getActiveSheet().deleteRows(0, 10);
```

## deleteRow(row)

氓聢聽茅聶陇猫隆聦茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧Sheet氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| row | number | 忙聵炉 | 猫隆聦茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getActiveSheet().deleteRows(0);
```

## deleteColumns(column, columnCount)

氓聢聽茅聶陇氓聢聴茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧Sheet氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| column | number | 忙聵炉 | 氓聢聴茫聙聜 |
| columnCount | number | 忙聵炉 | 氓聢聴忙聲掳茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getActiveSheet().deleteColumns(0茂录聦 10);
```

## deleteColumn(column)

氓聢聽茅聶陇氓聢聴茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧Sheet氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| column | number | 忙聵炉 | 氓聢聴茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getActiveSheet().deleteColumn(0);
```

## findAll(text, findOptions)

忙聽鹿忙聧庐忙聦聡氓庐職莽職聞忙聺隆盲禄露忙聼楼忙聣戮氓聦鹿茅聟聧莽禄聶氓庐職氓颅聴莽卢娄盲赂虏莽職聞忙聣聙忙聹聣Range茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range[] - 忙聣聙忙聹聣氓聦鹿茅聟聧氓聢掳莽職聞氓聦潞氓聼聼莽職聞忙聲掳莽禄聞茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| text | string | 忙聵炉 | 猫娄聛忙聼楼忙聣戮莽職聞氓颅聴莽卢娄盲赂虏茫聙聜 |
| findOptions | SearchOptions | 氓聬娄 | 氓聟露盲禄聳忙聬聹莽麓垄忙聺隆盲禄露茂录聦氓聦聟忙聥卢忙聬聹莽麓垄忙聵炉氓聬娄茅聹聙猫娄聛氓聦鹿茅聟聧忙聲麓盲赂陋氓聧聲氓聟聝忙聽录茫聙聛氓聦潞氓聢聠氓陇搂氓掳聫氓聠聶茫聙聛忙聵炉氓聬娄盲陆驴莽聰篓忙颅拢氓聢聶茫聙聛忙聵炉氓聬娄氓聦鹿茅聟聧氓聟卢氓录聫茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const ranges = sheet.findAll('str');
const ranges = sheet.findAll('str', {
 matchEntireCell: true,
  matchCase: true,
});
```

## replaceAll(text, replaceText, replaceOptions)

忙聽鹿忙聧庐氓陆聯氓聣聧氓路楼盲陆聹猫隆篓盲赂颅忙聦聡氓庐職莽職聞忙聺隆盲禄露忙聼楼忙聣戮氓鹿露忙聸驴忙聧垄莽禄聶氓庐職莽職聞氓颅聴莽卢娄盲赂虏茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

number - 忙聣搂猫隆聦莽職聞忙聸驴忙聧垄忙聲掳茂录聦氓聫炉猫聝陆盲赂聨氓聦鹿茅聟聧氓聧聲氓聟聝忙聽录莽職聞忙聲掳莽聸庐盲赂聧氓聬聦茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

|  |  |  |  |
| --- | --- | --- | --- |
| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| text | string | 忙聵炉 | 猫娄聛忙聼楼忙聣戮莽職聞氓颅聴莽卢娄盲赂虏茫聙聜 |
| replaceText | string | 忙聵炉 | 忙聸驴忙聧垄氓聨聼氓搂聥氓颅聴莽卢娄盲赂虏莽職聞氓颅聴莽卢娄盲赂虏茫聙聜 |
| replaceOptions | SearchOptions |  | 氓聟露盲禄聳忙聸驴忙聧垄忙聺隆盲禄露茂录聦氓聦聟忙聥卢忙聵炉氓聬娄茅聹聙猫娄聛氓聦鹿茅聟聧忙聲麓盲赂陋氓聧聲氓聟聝忙聽录茫聙聛氓聦潞氓聢聠氓陇搂氓掳聫氓聠聶茫聙聛忙聵炉氓聬娄盲陆驴莽聰篓忙颅拢氓聢聶茫聙聛忙聵炉氓聬娄氓聦鹿茅聟聧氓聟卢氓录聫茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const replaceCount = sheet.replaceAll('str', 'otherStr');
const replaceCount = sheet.replaceAll('str', 'otherStr', {  matchEntireCell: true });
```

## filter(range)

氓聢聸氓禄潞Filter茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Filter - Filter 氓庐聻盲戮聥茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| range | Range | 忙聵炉 | 氓聦潞氓聼聼茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const range = sheet.getRange('A1:C5');
sheet.filter(range);
```

## getFilter()

猫聨路氓聫聳氓陆聯氓聣聧sheet盲赂聤莽職聞filter氓庐聻盲戮聥茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Filter 茂陆聹 null - Filter 氓庐聻盲戮聥茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const filter = sheet.getFilter();
```

## deleteFilter()

氓聢聽茅聶陇氓陆聯氓聣聧sheet盲赂聤莽職聞filter茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

void -

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
sheet.deleteFilter();
```