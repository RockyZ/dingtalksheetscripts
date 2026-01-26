# FilterCondition

## API描述

筛选条件。用于FilterCriteriaBuilder或者FilterCriteria的相关API中指定筛选条件。

## 属性说明

| 属性 | 类型 | 是否必传 | 说明 |
| --- | --- | --- | --- |
| operator | String | 是 | 筛选条件的运算符，可选以下值   * equal：相等。 * not-equal: 不相等。 * contains：包含。 * not-contains：不包含。 * starts-with：从...开始。 * not-starts-with：不是从...开始。 * ends-with：以...结束。 * not-ends-with：不是以...结束。 * greater：大于。 * greater-equal：大于等于。 * less：小于。 * less-equal：小于等于。 |
| value | string | number | 是 | 筛选条件的值。 |