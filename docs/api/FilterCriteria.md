# FilterCriteria

## copy()

æ ¹æ®å½å FilterCriteria çææ°ç FilterCriteriaBuilderã

### è¿åç»æ

FilterCriteriaBuilder - æ°çFilterCriteriaBuilderã

### ç¤ºä¾ä»£ç 

```
const builder = criteria.copy();
```

## getVisibleValues()

è·åè®¾ç½®çå¯è§å¼ã

### è¿åç»æ

string[] | null - å¯è§å¼ã

### ç¤ºä¾ä»£ç 

```
builder.getVisibleValues();
```

## getVisibleFontColor()

è·åè®¾ç½®çå¯è§å­ä½é¢è²ã

### è¿åç»æ

Color ï½ null - å¯è§å­ä½é¢è²ã

### ç¤ºä¾ä»£ç 

```
builder.getVisibleFontColor();
```

## getVisibleBackgroundColor()

è·åè®¾ç½®çå¯è§ååæ ¼é¢è²ã

### è¿åç»æ

Color ï½ null - å¯è§ååæ ¼é¢è²ã

### ç¤ºä¾ä»£ç 

```
builder.getVisibleBackgroundColor();
```

## getVisibleConditions()

è·åè®¾ç½®çç­éæ¡ä»¶ã

### è¿åç»æ

FilterCondition[] | null - ç­éæ¡ä»¶åè¡¨ã

### ç¤ºä¾ä»£ç 

```
builder.getVisibleConditions();
```

## getVisibleConditionsOperator()

è·åè®¾ç½®çä¸¤ä¸ªç­éæ¡ä»¶å³ç³»ã

### è¿åç»æ

'and' | 'or' | null - ç­éæ¡ä»¶å³ç³»ã

### ç¤ºä¾ä»£ç 

```
builder.getVisibleConditionsOperator();
```