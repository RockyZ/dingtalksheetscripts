# SetValueOptions

## API描述

设置值的选项。用于 Range.setValue()或者Range.setValues() API中指定设置值时的选项。

## 属性说明

| 属性 | 类型 | 是否必填 | 说明 |
| --- | --- | --- | --- |
| parseType | String | 是 | 值的解析类型。   * **raw**：表示设置值的时候不解析。 * **useEntered**：表示策略和用户输入一致，例如输入 100%，会被解析成值为1，数字格式为%。 |