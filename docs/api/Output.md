# Output

Output 对象提供了在脚本执行面板中输出信息的方法，用于调试和向用户展示执行结果。

## 方法

### log(message)

在输出面板打印普通信息。

**参数：**
- `message` (any) - 要打印的信息，可以是字符串、数字、对象等任何类型

**返回值：**
- `void`

**示例：**

```javascript
// 输出字符串
Output.log("脚本开始执行");

// 输出数字
const count = 42;
Output.log(`处理了 ${count} 条记录`);

// 输出对象
const data = { name: "张三", age: 25 };
Output.log(data);

// 输出数组
const items = ["苹果", "香蕉", "橙子"];
Output.log(`水果列表：${items.join(", ")}`);

// 输出计算结果
const sum = 10 + 20;
Output.log(`计算结果：${sum}`);
```

---

### error(message)

在输出面板打印错误信息，背景色为红色，用于突出显示错误或警告信息。

**参数：**
- `message` (any) - 要打印的错误信息

**返回值：**
- `void`

**示例：**

```javascript
// 输出错误信息
Output.error("数据验证失败！");

// 输出异常信息
try {
  const sheet = Workbook.getSheet("不存在的表");
} catch (e) {
  Output.error(`获取工作表失败：${e.message}`);
}

// 输出警告信息
const value = sheet.getRange('A1').getValue();
if (value === null || value === "") {
  Output.error("警告：A1 单元格为空");
}

// 输出详细错误
const errors = [];
if (!userName) errors.push("用户名不能为空");
if (!userEmail) errors.push("邮箱不能为空");

if (errors.length > 0) {
  Output.error("表单验证失败：");
  errors.forEach(err => Output.error(`- ${err}`));
}
```

## 使用场景

### 1. 调试信息输出

```javascript
Output.log("开始处理数据...");

const sheet = Workbook.getActiveSheet();
const range = sheet.getRange('A1:B10');
const values = range.getValues();

Output.log(`读取了 ${values.length} 行数据`);

for (let i = 0; i < values.length; i++) {
  Output.log(`第 ${i + 1} 行：${values[i][0]}, ${values[i][1]}`);
}

Output.log("数据处理完成");
```

### 2. 进度提示

```javascript
const totalRows = 100;

for (let i = 0; i < totalRows; i++) {
  // 处理数据...
  
  if (i % 10 === 0) {
    const progress = Math.round((i / totalRows) * 100);
    Output.log(`处理进度：${progress}%`);
  }
}

Output.log("✓ 所有数据处理完成");
```

### 3. 错误处理和验证

```javascript
async function validateAndImport() {
  try {
    const files = await Input.filesAsync("请上传数据文件", {
      multiple: false,
      accept: ".csv"
    });
    
    if (!files || files.length === 0) {
      Output.error("错误：未选择文件");
      return;
    }
    
    const content = await files[0].text();
    
    if (content.trim() === "") {
      Output.error("错误：文件内容为空");
      return;
    }
    
    const rows = content.split('\n');
    
    if (rows.length < 2) {
      Output.error("错误：文件至少需要包含表头和一行数据");
      return;
    }
    
    Output.log(`✓ 文件验证通过，共 ${rows.length} 行数据`);
    
    // 导入数据...
    const sheet = Workbook.getActiveSheet();
    const data = rows.map(row => row.split(','));
    sheet.getRange(0, 0, data.length, data[0].length).setValues(data);
    
    Output.log("✓ 数据导入成功");
    
  } catch (error) {
    Output.error(`导入失败：${error.message}`);
  }
}

await validateAndImport();
```

### 4. 结果汇总

```javascript
const sheet = Workbook.getActiveSheet();
const range = sheet.getRange('A1:C100');
const values = range.getValues();

let successCount = 0;
let errorCount = 0;
const errors = [];

for (let i = 0; i < values.length; i++) {
  const [name, email, phone] = values[i];
  
  if (!email || !email.includes('@')) {
    errorCount++;
    errors.push(`第 ${i + 1} 行：邮箱格式错误`);
  } else {
    successCount++;
  }
}

Output.log("=== 数据验证结果 ===");
Output.log(`✓ 验证通过：${successCount} 条`);

if (errorCount > 0) {
  Output.error(`✗ 验证失败：${errorCount} 条`);
  Output.error("错误详情：");
  errors.forEach(err => Output.error(`  ${err}`));
} else {
  Output.log("✓ 所有数据验证通过");
}
```

## 使用技巧

1. **使用表情符号**：使用 ✓、✗、⚠️ 等符号让输出更直观
2. **格式化输出**：使用模板字符串和换行符让输出更易读
3. **分级输出**：普通信息用 `log()`，错误和警告用 `error()`
4. **进度反馈**：在长时间运行的脚本中定期输出进度信息
5. **调试信息**：开发时多用 `log()` 输出中间结果，便于排查问题

## 完整示例

```javascript
async function processDataWithLogging() {
  Output.log("=== 开始数据处理 ===");
  
  try {
    // 获取用户输入
    const sheetName = await Input.textAsync("请输入工作表名称：");
    Output.log(`目标工作表：${sheetName}`);
    
    // 获取工作表
    const sheet = Workbook.getSheet(sheetName);
    if (!sheet) {
      Output.error(`错误：找不到工作表 "${sheetName}"`);
      return;
    }
    
    Output.log("✓ 工作表获取成功");
    
    // 读取数据
    const range = sheet.getUsedRange();
    const values = range.getValues();
    Output.log(`读取了 ${values.length} 行数据`);
    
    // 处理数据
    let processedCount = 0;
    let errorCount = 0;
    
    for (let i = 1; i < values.length; i++) {
      try {
        // 数据处理逻辑...
        processedCount++;
        
        if (i % 10 === 0) {
          Output.log(`处理进度：${i}/${values.length}`);
        }
      } catch (error) {
        errorCount++;
        Output.error(`第 ${i + 1} 行处理失败：${error.message}`);
      }
    }
    
    // 输出结果
    Output.log("=== 处理完成 ===");
    Output.log(`✓ 成功处理：${processedCount} 条`);
    
    if (errorCount > 0) {
      Output.error(`✗ 处理失败：${errorCount} 条`);
    } else {
      Output.log("✓ 所有数据处理成功");
    }
    
  } catch (error) {
    Output.error(`脚本执行失败：${error.message}`);
  }
}

await processDataWithLogging();
```
