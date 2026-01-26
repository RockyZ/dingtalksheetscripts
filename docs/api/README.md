# 钉钉表格 API 文档

本目录包含钉钉表格脚本的所有API对象的详细文档。

## API 对象列表

### 核心对象

- [Workbook](Workbook.md) - 工作簿（顶级对象）
- [Sheet](Sheet.md) - 工作表
- [Range](Range.md) - 单个选区
- [RangeList](RangeList.md) - 多个选区

### 数据和格式选项

- [SetValueOptions](SetValueOptions.md) - 设置值的选项
- [DropdownListOption](DropdownListOption.md) - 下拉列表选项
- [BorderType](BorderType.md) - 边框类型
- [Color](Color.md) - 颜色

### 筛选和搜索

- [Filter](Filter.md) - 筛选
- [FilterCriteria](FilterCriteria.md) - 筛选规则
- [FilterCriteriaBuilder](FilterCriteriaBuilder.md) - 筛选规则构造器
- [FilterCondition](FilterCondition.md) - 筛选条件
- [SearchOptions](SearchOptions.md) - 搜索选项

### 其他功能

- [Hyperlink](Hyperlink.md) - 超链接
- [SortField](SortField.md) - 排序字段

## 文档说明

所有API文档均由钉钉官方HTML文档转换而来，包含：

- **方法说明**：每个方法的功能描述
- **参数列表**：详细的参数说明，包括类型、是否必传等
- **返回值说明**：方法返回值的类型和说明
- **示例代码**：实际使用示例

## 文档获取和生成

### 从官方文档生成

本项目的API文档是从钉钉官方OSS文档自动转换生成的。如需更新或重新生成文档，可以使用以下方法：

#### 1. 准备工具

确保已安装 `markitdown` 工具：

```bash
pip install markitdown
```

#### 2. 下载并转换文档

使用以下命令下载HTML文档并转换为Markdown格式：

```bash
# 进入API文档目录
cd docs/api

# 下载HTML文档
curl -o <ObjectName>.html https://icms-document.oss-cn-beijing.aliyuncs.com/zh-CN/dingtalk/orgapp/topics/<ObjectName>.html

# 转换为Markdown
markitdown <ObjectName>.html -o <ObjectName>.md
```

#### 3. 示例

以Filter API为例：

```bash
# 下载Filter API文档
curl -o Filter-api.html https://icms-document.oss-cn-beijing.aliyuncs.com/zh-CN/dingtalk/orgapp/topics/Filter-api.html

# 转换为Markdown
markitdown Filter-api.html -o Filter.md
```

以Hyperlink API为例：

```bash
# 下载Hyperlink API文档
curl -o Hyperlink-script.html https://icms-document.oss-cn-beijing.aliyuncs.com/zh-CN/dingtalk/orgapp/topics/Hyperlink-script.html

# 转换为Markdown
markitdown Hyperlink-script.html -o Hyperlink.md
```

### 官方文档URL规则

钉钉表格脚本API文档的OSS地址遵循以下规则：

```
https://icms-document.oss-cn-beijing.aliyuncs.com/zh-CN/dingtalk/orgapp/topics/{对象名}.html
```

常见对象的URL后缀：
- `Filter-api.html` - Filter对象
- `Hyperlink-script.html` - Hyperlink对象
- `RangeList.html` - RangeList对象
- `Range.html` - Range对象
- `Sheet.html` - Sheet对象
- `Workbook.html` - Workbook对象

### 注意事项

1. **文件命名**：不同API对象的HTML文件名可能有不同的后缀（如`-api`、`-script`等），需要参考实际可用的URL。
2. **格式修复**：转换后的Markdown文件可能需要手动调整格式，特别是表格中的复杂内容。
3. **内容验证**：转换完成后，建议对照官方文档验证内容的完整性和准确性。
4. **临时文件清理**：转换完成后记得删除临时的HTML文件。

## 快速开始

请参考项目根目录的 [README.md](../../README.md) 了解基本使用方式。
