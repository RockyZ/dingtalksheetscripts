# SortField

## API描述

排序字段。用于Filter.sort() API中指定排序方式。

## 属性说明

| 属性 | 类型 | 是否必填 | 说明 |
| --- | --- | --- | --- |
| column | number | 是 | 表示与range第一列的偏移量。 |
| ascending | boolean | 否 | 指定是否以升序完成排序。   * false：降序，默认值。 * true：升序。 |