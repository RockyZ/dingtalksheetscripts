# Range

## isEntireSheet()

忙聵炉氓聬娄忙聵炉忙聲麓盲赂陋氓路楼盲陆聹猫隆篓氓聦潞氓聼聼茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

boolean - 忙聵炉氓聬娄忙聵炉忙聲麓盲赂陋氓路楼盲陆聹猫隆篓氓聦潞氓聼聼茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const isEntireSheet = Workbook.getRange('A1').isEntireSheet();
```

## isEntireRow()

忙聵炉氓聬娄忙聵炉忙聲麓猫隆聦茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

boolean - 忙聵炉氓聬娄忙聵炉忙聲麓猫隆聦茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const isEntireRow = Workbook.getRange('A1').isEntireRow();
```

## isEntireColumn()

忙聵炉氓聬娄忙聵炉忙聲麓氓聢聴茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

boolean - 忙聵炉氓聬娄忙聵炉忙聲麓氓聢聴茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const isEntireColumn = Workbook.getRange('A1').isEntireColumn();
```

## isCell()

忙聵炉氓聬娄忙聵炉盲赂聙盲赂陋氓聧聲氓聟聝忙聽录茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

boolean - 忙聵炉氓聬娄忙聵炉盲赂聙盲赂陋氓聧聲氓聟聝忙聽录茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const isCell = Workbook.getRange('A1').isCell();
```

## activate()

忙驴聙忙麓禄茅聙聣氓聦潞茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 氓聦潞氓聼聼茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getRange('A1').activate()
```

## clear()

忙赂聟茅聶陇忙聽路氓录聫氓聮聦忙聲掳忙聧庐茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 氓聦潞氓聼聼茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getRange('A1').clear();
```

## clearFormat()

忙赂聟茅聶陇忙聽路氓录聫茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 氓聦潞氓聼聼茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getRange('A1').clearFormat();
```

## clearData()

忙赂聟茅聶陇忙聲掳忙聧庐茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 氓聦潞氓聼聼茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getRange('A1').clearData();
```

## setBackgroundColor(backgroundColor)

猫庐戮莽陆庐茅聙聣氓聦潞氓聧聲氓聟聝忙聽录猫聝聦忙聶炉茅垄聹猫聣虏茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 氓聦潞氓聼聼茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| backgroundColor | string | null | 忙聵炉 | 猫聝聦忙聶炉茅垄聹猫聣虏, null 猫隆篓莽陇潞忙赂聟茅聶陇茫聙聜盲陆驴莽聰篓氓聧聛氓聟颅猫驴聸氓聢露氓颅聴莽卢娄盲赂虏氓陆垄氓录聫猫隆篓莽陇潞茅垄聹猫聣虏茂录聦盲戮聥氓娄聜 '#ff0000' 猫隆篓莽陇潞莽潞垄猫聣虏茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getRange('A1').setBackgroundColor('#ff0000');
```

## getBackgroundColor()

猫聨路氓聫聳氓路娄盲赂聤猫搂聮氓聧聲氓聟聝忙聽录猫聝聦忙聶炉猫聣虏茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Color - 猫聝聦忙聶炉茅垄聹猫聣虏茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const color = Workbook.getRange('A1').getBackgroundColor();
```

## getBackgroundColors()

猫聨路氓聫聳氓聦潞氓聼聼氓聧聲氓聟聝忙聽录莽職聞猫聝聦忙聶炉猫聣虏茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Array<Array<Color>>> - 猫聝聦忙聶炉茅垄聹猫聣虏茂录聦盲潞聦莽禄麓忙聲掳莽禄聞茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const colors = Workbook.getRange('A1').getBackgroundColors();
```

## setFontColor(fontColor)

猫庐戮莽陆庐氓颅聴盲陆聯茅垄聹猫聣虏茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 氓聦潞氓聼聼茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| fontColor | string | null | 忙聵炉 | 氓颅聴盲陆聯茅垄聹猫聣虏, null 猫隆篓莽陇潞忙赂聟茅聶陇茫聙聜盲陆驴莽聰篓氓聧聛氓聟颅猫驴聸氓聢露氓颅聴莽卢娄盲赂虏氓陆垄氓录聫猫隆篓莽陇潞茅垄聹猫聣虏茂录聦盲戮聥氓娄聜 '#ff0000' 猫隆篓莽陇潞莽潞垄猫聣虏茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getRange('A1').setFontColor('#ff0000');
```

## getFontColor()

猫聨路氓聫聳氓路娄盲赂聤猫搂聮氓聧聲氓聟聝忙聽录氓颅聴盲陆聯茅垄聹猫聣虏茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Color - 氓颅聴盲陆聯茅垄聹猫聣虏茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const color = Workbook.getRange('A1').getFontColor();
```

## getFontColors()

猫聨路氓聫聳氓聦潞氓聼聼氓聧聲氓聟聝忙聽录氓颅聴盲陆聯茅垄聹猫聣虏茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Array<Array<Color>>> - 氓颅聴盲陆聯茅垄聹猫聣虏茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const colors = Workbook.getRange('A1').getFontColors();
```

## setFontFamily(fontFamily)

猫庐戮莽陆庐氓颅聴盲陆聯茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 氓聦潞氓聼聼茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| fontFamily | string | null | 忙聵炉 | 氓颅聴盲陆聯, null 猫隆篓莽陇潞忙赂聟茅聶陇茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getRange('A1').setFontFamily('arial');
```

## getFontFamily()

猫聨路氓聫聳氓路娄盲赂聤猫搂聮氓聧聲氓聟聝忙聽录氓颅聴盲陆聯茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

string - 氓颅聴盲陆聯茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const family = Workbook.getRange('A1').getFontFamily();
```

## getFontFamilies()

猫聨路氓聫聳氓聦潞氓聼聼氓聧聲氓聟聝忙聽录氓颅聴盲陆聯茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Array<Array<string>> - 氓颅聴盲陆聯茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const families = Workbook.getRange('A1').getFontFamilies();
```

## setFontSize(fontSize)

猫庐戮莽陆庐氓颅聴盲陆聯氓陇搂氓掳聫茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 氓聦潞氓聼聼茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| fontSize | number | null | 忙聵炉 | 氓颅聴盲陆聯氓陇搂氓掳聫, null 猫隆篓莽陇潞忙赂聟茅聶陇茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getRange('A1').setFontSize(20);
```

## getFontSize()

猫聨路氓聫聳氓路娄盲赂聤猫搂聮氓聧聲氓聟聝忙聽录氓颅聴盲陆聯氓陇搂氓掳聫茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

number - 氓颅聴盲陆聯氓陇搂氓掳聫茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const size = Workbook.getRange('A1').getFontSize();
```

## getFontSizes()

猫聨路氓聫聳氓聦潞氓聼聼氓聧聲氓聟聝忙聽录氓颅聴盲陆聯氓陇搂氓掳聫茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Array<Array<number>> - 氓颅聴盲陆聯氓陇搂氓掳聫茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const sizes = Workbook.getRange('A1').getFontSizes();
```

## setFontWeight(fontWeight)

猫庐戮莽陆庐氓颅聴盲陆聯莽虏聴盲陆聯茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 氓聦潞氓聼聼茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| fontWeight | 'bold' | 'normal' | null | 忙聵炉 | 氓颅聴盲陆聯莽虏聴莽禄聠, null 猫隆篓莽陇潞忙赂聟茅聶陇茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getRange('A1').setFontWeight('bold');
```

## getFontWeight()

猫聨路氓聫聳氓路娄盲赂聤猫搂聮氓聧聲氓聟聝忙聽录氓颅聴盲陆聯莽虏聴盲陆聯茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

'bold' | 'normal' - 氓颅聴盲陆聯莽虏聴盲陆聯茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const weight = Workbook.getRange('A1').getFontWeight();
```

## getFontWeights()

猫聨路氓聫聳氓聦潞氓聼聼氓聧聲氓聟聝忙聽录氓颅聴盲陆聯莽虏聴盲陆聯茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Array<Array<'bold' | 'normal'>> - 氓颅聴盲陆聯莽虏聴盲陆聯茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const weights = Workbook.getRange('A1').getFontWeights();
```

## setFontStyle(fontStyle)

猫庐戮莽陆庐氓颅聴盲陆聯忙聽路氓录聫茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 氓聦潞氓聼聼茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

|  |  |  |  |
| --- | --- | --- | --- |
| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| fontStyle | 'italic' | 'normal' | null | 忙聵炉 | 氓颅聴盲陆聯忙聽路氓录聫, null 猫隆篓莽陇潞忙赂聟茅聶陇茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getRange('A1').setFontStyle('italic');
```

## getFontStyle()

猫聨路氓聫聳氓路娄盲赂聤猫搂聮氓聧聲氓聟聝忙聽录氓颅聴盲陆聯忙聽路氓录聫茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

'italic' | 'normal' - 氓颅聴盲陆聯忙聽路氓录聫茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const style = Workbook.getRange('A1').getFontStyle();
```

## getFontStyles()

猫聨路氓聫聳氓聦潞氓聼聼氓聧聲氓聟聝忙聽录氓颅聴盲陆聯忙聽路氓录聫茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Array<Array<'italic' | 'normal'>> - 氓颅聴盲陆聯忙聽路氓录聫茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const styles = Workbook.getRange('A1').getFontStyles();
```

## setIndent(indent)

猫庐戮莽陆庐莽录漏猫驴聸茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 氓聦潞氓聼聼茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| indent | number | null | 忙聵炉 | 莽录漏猫驴聸茂录聦氓驴聟茅隆禄盲赂潞忙颅拢忙聲麓忙聲掳茂录聦null 猫隆篓莽陇潞忙赂聟茅聶陇茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getRange('A1').setIndent(3);
```

## getIndent()

猫聨路氓聫聳氓路娄盲赂聤猫搂聮氓聧聲氓聟聝忙聽录莽录漏猫驴聸茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

number - 莽录漏猫驴聸茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const indent = Workbook.getRange('A1').getIndent();
```

## getIndents()

猫聨路氓聫聳氓聦潞氓聼聼氓聧聲氓聟聝忙聽录莽录漏猫驴聸茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Array<Array<number>> - 莽录漏猫驴聸茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const indents = Workbook.getRange('A1').getIndents();
```

## setHorizontalAlignment(alignment)

猫庐戮莽陆庐忙掳麓氓鹿鲁氓炉鹿茅陆聬忙聳鹿氓录聫茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 氓聦潞氓聼聼茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| horizontalAlignment | 'left' | 'center' | 'right' | 'general' | null | 忙聵炉 | 忙掳麓氓鹿鲁氓炉鹿茅陆聬忙聳鹿氓录聫, null 猫隆篓莽陇潞忙赂聟茅聶陇茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getRange('A1').setHorizontalAlignment('left');
```

## getHorizontalAlignment()

猫聨路氓聫聳氓路娄盲赂聤猫搂聮氓聧聲氓聟聝忙聽录忙掳麓氓鹿鲁氓炉鹿茅陆聬忙聳鹿氓录聫茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

'left' | 'center' | 'right' | 'general' - 忙掳麓氓鹿鲁氓炉鹿茅陆聬忙聳鹿氓录聫茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const alignment = Workbook.getRange('A1').getHorizontalAlignment();
```

## getHorizontalAlignments()

猫聨路氓聫聳氓聦潞氓聼聼氓聧聲氓聟聝忙聽录忙掳麓氓鹿鲁氓炉鹿茅陆聬忙聳鹿氓录聫茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Array<Array<'left' | 'center' | 'right' | 'general'>> - 忙掳麓氓鹿鲁氓炉鹿茅陆聬忙聳鹿氓录聫茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const alignments = Workbook.getRange('A1').getHorizontalAlignments();
```

## setVerticalAlignment(alignment)

猫庐戮莽陆庐氓聻聜莽聸麓氓炉鹿茅陆聬忙聳鹿氓录聫茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 氓聦潞氓聼聼茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

|  |  |  |  |
| --- | --- | --- | --- |
| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| verticalAlignment | 'top' | 'middle' | 'bottom' | null | 忙聵炉 | 氓颅聴盲陆聯氓聢聽茅聶陇莽潞驴, null 猫隆篓莽陇潞忙赂聟茅聶陇茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getRange('A1').setVerticalAlignments('top');
```

## getVerticalAlignment()

猫聨路氓聫聳氓路娄盲赂聤猫搂聮氓聧聲氓聟聝忙聽录氓聻聜莽聸麓氓炉鹿茅陆聬忙聳鹿氓录聫茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

'top' | 'middle' | 'bottom' - 氓聻聜莽聸麓氓炉鹿茅陆聬忙聳鹿氓录聫茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const alignment = Workbook.getRange('A1').getVerticalAlignment();
```

## getVerticalAlignments()

猫聨路氓聫聳氓聦潞氓聼聼氓聧聲氓聟聝忙聽录氓聻聜莽聸麓氓炉鹿茅陆聬忙聳鹿氓录聫茫聙聜茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Array<Array<'top' | 'middle' | 'bottom'>> - 氓聻聜莽聸麓氓炉鹿茅陆聬忙聳鹿氓录聫茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const alignments = Workbook.getRange('A1').getVerticalAlignments();
```

## setWordWrap(wordWrap)

猫庐戮莽陆庐忙聧垄猫隆聦忙聳鹿氓录聫茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 氓聦潞氓聼聼茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| wordWrap | 'overflow' | 'clip' | 'autoWrap' | null | 忙聵炉 | 忙聧垄猫隆聦忙聳鹿氓录聫, null 猫隆篓莽陇潞忙赂聟茅聶陇茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getRange('A1').setWordWrap('overflow');
```

## getWordWrap()

猫聨路氓聫聳氓路娄盲赂聤猫搂聮氓聧聲氓聟聝忙聽录忙聧垄猫隆聦忙聳鹿氓录聫茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

'overflow' | 'clip' | 'autoWrap' - 忙聧垄猫隆聦忙聳鹿氓录聫茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const wordWrap = Workbook.getRange('A1').getWordWrap();
```

## getWordWraps()

猫聨路氓聫聳氓聦潞氓聼聼氓聧聲氓聟聝忙聽录忙聧垄猫隆聦忙聳鹿氓录聫茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Array<Array<'overflow' | 'clip' | 'autoWrap'>> - 忙聧垄猫隆聦忙聳鹿氓录聫茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const wordWraps = Workbook.getRange('A1').getWordWraps();
```

## setBorder(null)

忙赂聟茅聶陇猫戮鹿忙隆聠茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 氓聦潞氓聼聼茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
workbook.getActiveRange().setBorder(null);
```

## setBorder(type, color, style)

猫庐戮莽陆庐猫戮鹿忙隆聠茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 氓聦潞氓聼聼茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| type | BorderType | 忙聵炉 | 猫戮鹿忙隆聠莽卤禄氓聻聥茫聙聜 |
| color | string | 忙聵炉 | 猫戮鹿忙隆聠茅垄聹猫聣虏茫聙聜 |
| style | 'none' |  'dotted'|  'dashed'|  'solid'|  'medium'|  'thick'|  'double'|  'hair'|  'dashDotDot'|  'dashDot'|  'mediumDashDotDot'|  'slantDashDot'|  'mediumDashDot'|  'mediumDashed' | 忙聵炉 | 猫戮鹿忙隆聠忙聽路氓录聫茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getActiveRange().setBorder({ top: true, bottom: true },  '#ff00ff', 'solid');
```

## decreaseDecimal()

氓聡聫氓掳聭氓掳聫忙聲掳盲陆聧茂录聦氓娄聜忙聻聹 activeCell 氓聹篓 Range 氓聠聟茂录聦盲禄楼 activeCell 盲赂潞氓聫聜莽聟搂茂录聦氓聬娄氓聢聶盲禄楼氓路娄盲赂聤猫搂聮氓聧聲氓聟聝忙聽录盲赂潞氓聫聜莽聟搂茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 氓聦潞氓聼聼茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getRange('A1').decreaseDecimal();
```

## increaseDecimal()

氓聡聫氓掳聭氓掳聫忙聲掳盲陆聧茂录聦氓娄聜忙聻聹 activeCell 氓聹篓 Range 氓聠聟茂录聦盲禄楼 activeCell 盲赂潞氓聫聜莽聟搂茂录聦氓聬娄氓聢聶盲禄楼氓路娄盲赂聤猫搂聮氓聧聲氓聟聝忙聽录盲赂潞氓聫聜莽聟搂茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 氓聦潞氓聼聼茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getRange('A1').increaseDecimal();
```

## setNumberFormat(numberFormat)

猫庐戮莽陆庐忙聲掳氓颅聴忙聽录氓录聫茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 氓聦潞氓聼聼茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| numberFormat | string | 忙聵炉 | 忙聲掳氓颅聴忙聽录氓录聫茂录聦猫炉娄忙聝聟氓聫聜猫聙聝[忙聲掳氓颅聴忙聽录氓录聫](/document/orgapp-client/number-format#topic-2193105)茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getActiveRange().setNumberFormat('0.00');
```

## getNumberFormat()

猫聨路氓聫聳氓路娄盲赂聤猫搂聮氓聧聲氓聟聝忙聽录忙聲掳氓颅聴忙聽录氓录聫茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

string - 忙聲掳氓颅聴忙聽录氓录聫茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const numberFormat = Workbook.getActiveRange().getNumberFormat();
```

## getNumberFormats()

猫聨路氓聫聳氓聦潞氓聼聼氓聧聲氓聟聝忙聽录忙聲掳氓颅聴忙聽录氓录聫茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Array<Array<string>> - 忙聲掳氓颅聴忙聽录氓录聫茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const numberFormats = Workbook.getActiveRange().getNumberFormats();
```

## getFormula()

猫聨路氓聫聳氓路娄盲赂聤猫搂聮氓聧聲氓聟聝忙聽录氓聟卢氓录聫猫隆篓猫戮戮氓录聫茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

string - 氓聟卢氓录聫猫隆篓猫戮戮氓录聫茂录聦忙聴聽氓聟卢氓录聫盲赂潞莽漏潞氓颅聴莽卢娄盲赂虏茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const formula = Workbook.getActiveRange().getFormula();
```

## getFormulas()

猫聨路氓聫聳氓聦潞氓聼聼氓聧聲氓聟聝忙聽录氓聟卢氓录聫猫隆篓猫戮戮氓录聫茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Array<Array<string>> - 氓聟卢氓录聫猫隆篓猫戮戮氓录聫茂录聦忙聴聽氓聟卢氓录聫盲赂潞莽漏潞氓颅聴莽卢娄盲赂虏茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const formulas = Workbook.getActiveRange().getFormulas();
```

## getDisplayValue()

猫聨路氓聫聳氓路娄盲赂聤猫搂聮氓聧聲氓聟聝忙聽录忙聵戮莽陇潞氓聙录茫聙聜忙聵戮莽陇潞氓聙录忙聦聡氓聼潞盲潞聨忙聲掳氓颅聴忙聽录氓录聫莽颅聣猫庐隆莽庐聴氓聬聨莽職聞氓聙录茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

string - 氓聧聲氓聟聝忙聽录忙聵戮莽陇潞氓聙录茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const value = Workbook.getActiveRange().getDisplayValue();
```

## getDisplayValues()

猫聨路氓聫聳氓聦潞氓聼聼氓聧聲氓聟聝忙聽录忙聵戮莽陇潞氓聙录茫聙聜忙聵戮莽陇潞氓聙录忙聦聡氓聼潞盲潞聨忙聲掳氓颅聴忙聽录氓录聫莽颅聣猫庐隆莽庐聴氓聬聨莽職聞氓聙录茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Array<Array<string>> - 氓聧聲氓聟聝忙聽录忙聵戮莽陇潞氓聙录茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const formulas = Workbook.getActiveRange().getFormulas();
```

## getValue()

猫聨路氓聫聳氓路娄盲赂聤猫搂聮氓聧聲氓聟聝忙聽录氓聙录茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

string | number | boolean | null - 氓聧聲氓聟聝忙聽录氓聙录茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const value = Workbook.getActiveRange().getValue();
```

## getDisplayValues()

猫聨路氓聫聳氓聦潞氓聼聼氓聧聲氓聟聝忙聽录氓聙录茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Array<Array<string | number | boolean | null>> - 氓聧聲氓聟聝忙聽录氓聙录茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const values = Workbook.getActiveRange().getValues();
```

## setValues(values, options)

猫庐戮莽陆庐氓聙录茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 氓聦潞氓聼聼茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| value | Array<Array<string | boolean | number | null>> | 忙聵炉 | 氓聙录茂录聦盲录聽null 猫隆篓莽陇潞忙赂聟茅聶陇茫聙聜 |
| options | SetValueOptions | 氓聬娄 | 猫庐戮莽陆庐氓聙录莽職聞茅聟聧莽陆庐茅聙聣茅隆鹿茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getRange('A1:B1').setValues([
  [1, 2],
]);

Workbook.getRange('A1:B1').setValues([
  [1, 2],
], { parseType: 'raw' });
```

## merge()

氓聬聢氓鹿露氓聧聲氓聟聝忙聽录茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 氓聦潞氓聼聼茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getActiveRange().merge();
```

## unmerge()

氓聫聳忙露聢氓聬聢氓鹿露氓聧聲氓聟聝忙聽录茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 氓聦潞氓聼聼茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getActiveRange().unmerge();
```

## getMergedRanges()

猫聨路氓聫聳氓聬聢氓鹿露氓聧聲氓聟聝忙聽录氓聦潞氓聼聼茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range[] - 氓聬聢氓鹿露氓聧聲氓聟聝忙聽录氓聦潞氓聼聼茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const ranges = Workbook.getActiveRange().getMergedRanges();
```

## getA1Notation()

猫聨路氓聫聳氓聦潞氓聼聼氓聹掳氓聺聙茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

string - 氓聦潞氓聼聼氓颅聴莽卢娄盲赂虏氓聹掳氓聺聙茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const notation = Workbook.getActiveRange().getA1Notation();
```

## getRow()

猫聨路氓聫聳盲禄聨0氓录聙氓搂聥莽職聞猫隆聦氓聫路茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

number - 猫隆聦茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const row = Workbook.getActiveRange().getRow();
```

## getColumn()

猫聨路氓聫聳盲禄聨0氓录聙氓搂聥莽職聞氓聢聴氓聫路茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

number - 氓聢聴茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const column = Workbook.getActiveRange().getColumn();
```

## getRowCount()

猫聨路氓聫聳猫隆聦忙聲掳茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

number - 猫隆聦忙聲掳茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const count = Workbook.getActiveRange().getRowCount();
```

## getColumnCount()

猫聨路氓聫聳氓聢聴忙聲掳茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

number - 氓聢聴忙聲掳茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const count = Workbook.getActiveRange().getColumnCount();
```

## getHyperlink()

猫聨路氓聫聳氓聧聲氓聟聝忙聽录茅聯戮忙聨楼茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Hyperlink | null - 氓聧聲氓聟聝忙聽录茅聯戮忙聨楼, null 猫隆篓莽陇潞忙聴聽茅聯戮忙聨楼茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const link = Workbook.getActiveRange().getHyperlink();
```

## sort(field)

忙聨聮氓潞聫氓陆聯氓聣聧氓聦潞氓聼聼茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 氓聦潞氓聼聼茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| field | SortField | number | 忙聵炉 | 忙聨聮氓潞聫猫搂聞氓聢聶 茂陆聹 莽聸赂氓炉鹿盲潞聨氓陆聯氓聣聧 Range 茅娄聳氓聢聴莽職聞氓聛聫莽搂禄茅聡聫茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
// 氓聼潞盲潞聨氓陆聯氓聣聧茅聙聣氓聦潞盲赂颅莽職聞莽卢卢0氓聢聴猫驴聸猫隆聦茅聙聠氓潞聫忙聨聮氓潞聫
Workbook.getActiveRange().sort({  column: 0,  ascending: false });

// 氓聼潞盲潞聨氓陆聯氓聣聧茅聙聣氓聦潞盲赂颅莽職聞莽卢卢0氓聢聴猫驴聸猫隆聦茅隆潞氓潞聫忙聨聮氓潞聫
Workbook.getActiveRange().sort(0);
```

## find(text, findOptions)

忙聽鹿忙聧庐忙聦聡氓庐職莽職聞忙聺隆盲禄露忙聼楼忙聣戮氓聦鹿茅聟聧莽禄聶氓庐職氓颅聴莽卢娄盲赂虏莽職聞莽卢卢盲赂聙盲赂陋Range茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 氓聦潞氓聼聼茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| text | string | 忙聵炉 | 猫娄聛忙聼楼忙聣戮莽職聞氓颅聴莽卢娄盲赂虏茫聙聜 |
| findOptions | SearchOptions | 氓聬娄 | 氓聟露盲禄聳忙聬聹莽麓垄忙聺隆盲禄露茂录聦氓聦聟忙聥卢忙聬聹莽麓垄忙聵炉氓聬娄茅聹聙猫娄聛氓聦鹿茅聟聧忙聲麓盲赂陋氓聧聲氓聟聝忙聽录茫聙聛氓聦潞氓聢聠氓陇搂氓掳聫氓聠聶茫聙聛忙聵炉氓聬娄盲陆驴莽聰篓忙颅拢氓聢聶猫隆篓猫戮戮氓录聫茫聙聛忙聵炉氓聬娄氓聦鹿茅聟聧氓聟卢氓录聫茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const range = Workbook.getRange('A1:C10');
const findRange1 = range.find('1');
const findRange2 = range.find('1', {  matchEntireCell: true });
```

## findNext(text, findOptions)

忙聽鹿忙聧庐忙聦聡氓庐職莽職聞忙聺隆盲禄露忙聼楼忙聣戮氓聦鹿茅聟聧莽禄聶氓庐職氓颅聴莽卢娄盲赂虏盲赂聥盲赂聙盲赂陋Range茂录聦忙聬聹莽麓垄猫聦聝氓聸麓盲赂潞盲禄楼氓陆聯氓聣聧Range氓路娄盲赂聤猫搂聮莽職聞Cell盲鹿聥氓聬聨氓录聙氓搂聥莽職聞忙聲麓盲赂陋氓路楼盲陆聹猫隆篓茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 氓聦潞氓聼聼茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| text | string | 忙聵炉 | 猫娄聛忙聼楼忙聣戮莽職聞氓颅聴莽卢娄盲赂虏茫聙聜 |
| findOptions | SearchOptions | 氓聬娄 | 氓聟露盲禄聳忙聬聹莽麓垄忙聺隆盲禄露茂录聦氓聦聟忙聥卢忙聬聹莽麓垄忙聵炉氓聬娄茅聹聙猫娄聛氓聦鹿茅聟聧忙聲麓盲赂陋氓聧聲氓聟聝忙聽录茫聙聛氓聦潞氓聢聠氓陇搂氓掳聫氓聠聶茫聙聛忙聵炉氓聬娄盲陆驴莽聰篓忙颅拢氓聢聶茫聙聛忙聵炉氓聬娄氓聦鹿茅聟聧氓聟卢氓录聫茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const range = Workbook.getRange('A2');
const findRange = range.findNext('1');
```

## findPrevious(text, findOptions)

忙聽鹿忙聧庐忙聦聡氓庐職莽職聞忙聺隆盲禄露忙聼楼忙聣戮氓聦鹿茅聟聧莽禄聶氓庐職氓颅聴莽卢娄盲赂虏盲赂聤盲赂聙盲赂陋Range茂录聦忙聬聹莽麓垄猫聦聝氓聸麓盲赂潞盲禄楼氓陆聯氓聣聧Range氓路娄盲赂聤猫搂聮莽職聞Cell盲鹿聥氓聣聧莽職聞忙聲麓盲赂陋氓路楼盲陆聹猫隆篓茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - 氓聦潞氓聼聼茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| text | string | 忙聵炉 | 猫娄聛忙聼楼忙聣戮莽職聞氓颅聴莽卢娄盲赂虏茫聙聜 |
| findOptions | SearchOptions | 氓聬娄 | 氓聟露盲禄聳忙聬聹莽麓垄忙聺隆盲禄露茂录聦氓聦聟忙聥卢忙聬聹莽麓垄忙聵炉氓聬娄茅聹聙猫娄聛氓聦鹿茅聟聧忙聲麓盲赂陋氓聧聲氓聟聝忙聽录茫聙聛氓聦潞氓聢聠氓陇搂氓掳聫氓聠聶茫聙聛忙聵炉氓聬娄盲陆驴莽聰篓忙颅拢氓聢聶茫聙聛忙聵炉氓聬娄氓聦鹿茅聟聧氓聟卢氓录聫茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const range = Workbook.getRange('A2');
const findRange = range.findPrevious('1');
```

## replaceAll(text, replaceText, replaceOptions)

忙聽鹿忙聧庐忙聦聡氓庐職莽職聞忙聺隆盲禄露忙聼楼忙聣戮氓鹿露忙聸驴忙聧垄莽禄聶氓庐職莽職聞氓颅聴莽卢娄盲赂虏茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

number - 忙聣搂猫隆聦莽職聞忙聸驴忙聧垄忙聲掳茂录聦氓聫炉猫聝陆盲赂聨氓聦鹿茅聟聧氓聧聲氓聟聝忙聽录莽職聞忙聲掳莽聸庐盲赂聧氓聬聦茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| text | string | 忙聵炉 | 猫娄聛忙聼楼忙聣戮莽職聞氓颅聴莽卢娄盲赂虏茫聙聜 |
| replaceText | string | 忙聵炉 | 忙聸驴忙聧垄氓聨聼氓搂聥氓颅聴莽卢娄盲赂虏莽職聞氓颅聴莽卢娄盲赂虏茫聙聜 |
| replaceOptions | SearchOptions |  | 氓聟露盲禄聳忙聸驴忙聧垄忙聺隆盲禄露茂录聦氓聦聟忙聥卢忙聵炉氓聬娄茅聹聙猫娄聛氓聦鹿茅聟聧忙聲麓盲赂陋氓聧聲氓聟聝忙聽录茫聙聛氓聦潞氓聢聠氓陇搂氓掳聫氓聠聶茫聙聛忙聵炉氓聬娄盲陆驴莽聰篓忙颅拢氓聢聶茫聙聛忙聵炉氓聬娄氓聦鹿茅聟聧氓聟卢氓录聫茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const range = Workbook.getRange('A1:C10');
const replaceCount = range.replaceAll('str', 'otherStr');
```

## insertCheckboxes()

忙聫聮氓聟楼氓陇聧茅聙聣忙隆聠茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - Range茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const range = Workbook.getRange(2, 2, 2, 2);
range.insertCheckboxes();
```

## deleteCheckboxes()

忙赂聟茅聶陇氓陇聧茅聙聣忙隆聠茂录聦value 盲赂聧盲录職忙赂聟茅聶陇茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - Range茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const range = Workbook.getRange(2, 2, 2, 2);
range.deleteCheckboxes();
```

## getCheckedState()

猫驴聰氓聸聻氓路娄盲赂聤猫搂聮氓聧聲氓聟聝忙聽录忙聵炉氓聬娄氓聥戮茅聙聣茂录聦null猫隆篓莽陇潞氓聧聲氓聟聝忙聽录盲赂聧忙聵炉氓陇聧茅聙聣忙隆聠茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

boolean | null - 忙聵炉氓聬娄氓聥戮茅聙聣茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const state = Workbook.getRange(2, 2, 2, 2).getCheckedState();
```

## getCheckedStates()

猫驴聰氓聸聻氓聦潞氓聼聼氓聧聲氓聟聝忙聽录忙聵炉氓聬娄氓聥戮茅聙聣茂录聦null猫隆篓莽陇潞氓聧聲氓聟聝忙聽录盲赂聧忙聵炉氓陇聧茅聙聣忙隆聠茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Array<Array<boolean | null>> - 氓聦潞氓聼聼氓聧聲氓聟聝忙聽录氓聥戮茅聙聣莽職聞莽聤露忙聙聛茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const states = Workbook.getRange(2, 2, 2, 2).getCheckedStates();
```

## setCheckedState(checked)

氓聥戮茅聙聣/氓聫聳忙露聢氓聥戮茅聙聣氓陇聧茅聙聣忙隆聠茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - Range茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| checked | boolean | 忙聵炉 | 忙聵炉氓聬娄氓聥戮茅聙聣茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getRange(2, 2, 2, 2).setCheckedState(true);
```

## setCheckedStates(checkeds)

氓聥戮茅聙聣/氓聫聳忙露聢氓聥戮茅聙聣氓陇聧茅聙聣忙隆聠茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - Range茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| checkeds | Array<Array<boolean | null>> | 忙聵炉 | 忙聵炉氓聬娄氓聥戮茅聙聣茂录聦null 猫隆篓莽陇潞猫路鲁猫驴聡茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getRange(2, 2, 2, 2).setCheckedStates([[true, false], [false, null]]);
```

## insertDropdownLists(options)

猫庐戮莽陆庐盲赂聥忙聥聣忙隆聠茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - Range茫聙聜

### 氓聫聜忙聲掳猫炉麓忙聵聨

| 氓聫聜忙聲掳 | 莽卤禄氓聻聥 | 忙聵炉氓聬娄氓驴聟盲录聽 | 猫炉麓忙聵聨 |
| --- | --- | --- | --- |
| options | DropdownListOption[] | 忙聵炉 | 盲赂聥忙聥聣茅聙聣茅隆鹿茫聙聜 |

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
const range = Workbook.getActiveRange();
range.insertDropdownLists([
    {
        value: '莽炉庐莽聬聝',
        color: '#ff0000',
    },
    {
        value: '猫露鲁莽聬聝',
        color: '#00ff00',
    },
    {
        value: '忙聨聮莽聬聝',
        color: '#0000ff',
    },
])
```

## deleteDropdownLists()

忙赂聟茅聶陇盲赂聥忙聥聣忙隆聠茂录聦value 盲赂聧盲录職忙赂聟茅聶陇茫聙聜

### 猫驴聰氓聸聻莽禄聯忙聻聹

Range - Range茫聙聜

### 莽陇潞盲戮聥盲禄拢莽聽聛

```
Workbook.getActiveRange().deleteDropdownLists();
```