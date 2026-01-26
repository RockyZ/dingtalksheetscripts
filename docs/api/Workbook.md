# Workbook

## insertSheet()

忙聫聮氓聟楼氓路楼盲陆聹猫隆篓茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 忙聳掳氓垄聻氓聤聽莽職聞氓路楼盲陆聹猫隆篓

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.insertSheet();
```

## insertSheet(name)

忙聫聮氓聟楼氓路楼盲陆聹猫隆篓茂录聦氓鹿露忙聦聡氓庐職氓聬聧氓颅聴茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 忙聳掳氓垄聻氓聤聽莽職聞氓路楼盲陆聹猫隆篓

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| name | string | 忙聵炉 | 忙聦聡氓庐職忙聳掳氓路楼盲陆聹猫隆篓氓聬聧氓颅聴 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.insertSheet('A new sheet');
```

## insertSheet(index)

忙聫聮氓聟楼氓路楼盲陆聹猫隆篓茂录聦氓鹿露忙聦聡氓庐職盲陆聧莽陆庐茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 忙聳掳氓垄聻氓聤聽莽職聞氓路楼盲陆聹猫隆篓

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| index | number | 忙聵炉 | 忙聦聡氓庐職忙聳掳氓路楼盲陆聹猫隆篓盲陆聧莽陆庐茂录聦盲禄聨 0 氓录聙氓搂聥茂录聦氓聹篓index 盲鹿聥氓聬聨忙聫聮氓聟楼氓路楼盲陆聹猫隆篓茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.insertSheet(0);
```

## insertSheet(name, index)

忙聫聮氓聟楼氓路楼盲陆聹猫隆篓茂录聦氓鹿露忙聦聡氓庐職氓聬聧氓颅聴氓聮聦盲陆聧莽陆庐茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 忙聳掳氓垄聻氓聤聽莽職聞氓路楼盲陆聹猫隆篓

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.insertSheet('hello world', 0);
```

## moveSheet(sheet, targetIndex)

莽搂禄氓聤篓 Worksheet 氓聢掳忙聦聡氓庐職盲陆聧莽陆庐茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Workbook - 氓陆聯氓聣聧猫隆篓忙聽录氓庐聻盲戮聥

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| sheet | Sheet | 忙聵炉 | 猫娄聛莽搂禄氓聤篓莽職聞莽聸庐忙聽聡氓路楼盲陆聹猫隆篓 |
| targetIndex | number | 忙聵炉 | 莽聸庐忙聽聡莽職聞莽麓垄氓录聲盲陆聧莽陆庐 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const sheet = Workbook.getSheetByIndex(0);
Workbook.moveSheet(sheet, 1);
```

## deleteSheet(name)

氓聢聽茅聶陇氓路楼盲陆聹猫隆篓茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Workbook - 氓陆聯氓聣聧猫隆篓忙聽录氓庐聻盲戮聥

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| name | Sheet | string | 忙聵炉 | 氓路楼盲陆聹猫隆篓氓炉鹿猫卤隆茫聙聛ID忙聢聳氓聬聧氓颅聴 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.deleteSheetByName('Sheet1');
```

## setActiveSheet(sheet)

忙驴聙忙麓禄氓路楼盲陆聹猫隆篓茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 忙驴聙忙麓禄莽職聞氓路楼盲陆聹猫隆篓

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| sheet | Sheet | 忙聵炉 | 忙聝鲁忙驴聙忙麓禄莽職聞氓路楼盲陆聹猫隆篓氓炉鹿猫卤隆茫聙聛ID忙聢聳氓聬聧氓颅聴 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const sheet = Workbook.getSheetByName('Sheet1');
Workbook.setActiveSheet(sheet);
```

## getSheets()

猫聨路氓聫聳忙聣聙忙聹聣莽職聞氓路楼盲陆聹猫隆篓茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet[] - 忙聣聙忙聹聣莽職聞氓路楼盲陆聹猫隆篓

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const sheets = Workbook.getSheets();
```

## getSheetCount()

猫聨路氓聫聳氓路楼盲陆聹猫隆篓莽職聞忙聲掳茅聡聫茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

number - 氓路楼盲陆聹猫隆篓莽職聞忙聲掳茅聡聫

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const sheetCount = Workbook.getSheetCount();

// 莽颅聣盲禄路盲潞聨
const sheets = Workbook.getSheets();
const sheetCount = sheets.length;
```

## getActiveSheet()

猫聨路氓聫聳氓陆聯氓聣聧忙驴聙忙麓禄莽職聞氓路楼盲陆聹猫隆篓茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet - 氓陆聯氓聣聧忙驴聙忙麓禄莽職聞氓路楼盲陆聹猫隆篓

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const sheet = Workbook.getActiveSheet();
sheet.getRange(0, 0, 2, 2).setBackgroundColor('yellow');
```

## getSheet(key)

忙聽鹿忙聧庐 ID 忙聢聳猫聙聟 name 猫聨路氓聫聳氓路楼盲陆聹猫隆篓茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Sheet | null - 忙聼楼猫炉垄氓聢掳莽職聞氓路楼盲陆聹猫隆篓茫聙聜null猫隆篓莽陇潞忙虏隆忙聹聣忙聼楼氓聢掳茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| key | string | 忙聵炉 | 氓路楼盲陆聹猫隆篓莽職聞 ID 忙聢聳猫聙聟 name |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const sheet1 = Workbook.getSheet('id1');
const sheet2 = Workbook.getSheet('Sheet2');
```

## getRange(a1Notation)

猫聨路氓聫聳氓陆聯氓聣聧氓路楼盲陆聹猫隆篓莽職聞茅聙聣氓聦潞猫聦聝氓聸麓茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 茅聙聣氓聦潞

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| a1Notation | string | 忙聵炉 | 茅聙聣氓聦潞氓聹掳氓聺聙 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const range = Workbook.getRange('A1:B4');

// 莽颅聣盲禄路盲潞聨
const range = Workbook.getActiveSheet().getRange('A1:B4');
```

## getRange(row, column, rowCount, columnCount)

猫聨路氓聫聳氓陆聯氓聣聧氓路楼盲陆聹猫隆篓莽職聞茅聙聣氓聦潞猫聦聝氓聸麓茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 茅聙聣氓聦潞

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| row | number | 忙聵炉 | 猫隆聦茂录聦盲禄聨 0 氓录聙氓搂聥 |
| column | number | 忙聵炉 | 氓聢聴茂录聦盲禄聨 0 氓录聙氓搂聥 |
| rowCount | number | 忙聵炉 | 猫隆聦忙聲掳茂录聦盲禄聨 1 氓录聙氓搂聥 |
| columnCount | number | 忙聵炉 | 氓聢聴忙聲掳茂录聦盲禄聨 1 氓录聙氓搂聥 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const range = Workbook.getRange(0, 0, 2, 2);

// 莽颅聣盲禄路盲潞聨
const range = Workbook.getActiveSheet().getRange(0, 0, 2, 2);
```

## getActiveCell()

猫聨路氓聫聳activeCell茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - activeCell

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const range = Workbook.getActiveCell();

// 莽颅聣盲禄路盲潞聨
const range = Workbook.getActiveSheet().getActiveCell();
```

## getActiveRange()

猫聨路氓聫聳忙驴聙忙麓禄莽職聞茅聙聣氓聦潞茫聙聜氓娄聜忙聻聹忙驴聙忙麓禄盲潞聠氓陇職盲赂陋盲赂聧猫驴聻莽禄颅氓聦潞氓聼聼茂录聦氓聢聶猫驴聰氓聸聻activeCell忙聣聙氓聹篓莽職聞茅聙聣氓聦潞茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range | null - 氓陆聯氓聣聧忙驴聙忙麓禄莽職聞茅聙聣氓聦潞茫聙聜氓娄聜忙聻聹氓陆聯氓聣聧茅聙聣盲赂颅盲潞聠氓聸戮猫隆篓莽颅聣茂录聦氓聢聶猫驴聰氓聸聻null茫聙聜

## getRangeList(a1Notations)

猫聨路氓聫聳氓陇職茅聙聣氓聦潞茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

RangeList | null - 氓陇職茅聙聣氓聦潞

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const rangeList = Workbook.getRangeList(['A1:B10', 'D2:E5']);

// 莽颅聣盲禄路盲潞聨
const rangeList = Workbook.getActiveSheet().getRangeList(['A1:B10', 'D2:E5']);
```

## getActiveRangeList()

猫聨路氓聫聳忙驴聙忙麓禄莽職聞氓陇職茅聙聣氓聦潞茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

RangeList - 忙驴聙忙麓禄莽職聞氓陇職茅聙聣氓聦潞

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const rangeList = Workbook.getActiveRangeList();

// 莽颅聣盲禄路盲潞聨
const rangeList = Workbook.getActiveSheet().getActiveRangeList();
```

## newFilterCriteriaBuilder()

忙聳掳氓禄潞盲赂聙盲赂陋莽颅聸茅聙聣忙聺隆盲禄露猫搂聞氓聢聶莽職聞莽聰聼忙聢聬氓聶篓茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

FilterCriteriaBuilder - 莽颅聸茅聙聣忙聺隆盲禄露猫搂聞氓聢聶莽聰聼忙聢聬氓聶篓

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const criteria = Workbook.newFilterCriteriaBuilder()
 .setVisibleValues(['3', '5'])
  .build();
const filter = Workbook.getActiveSheet().getFilter();
filter.setColumnFilterCriteria(1, criteria);
```