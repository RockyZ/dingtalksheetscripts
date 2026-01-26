# Input

Input 对象提供了与用户交互的方法，用于获取用户输入、选择和文件上传。所有方法都是异步的，需要使用 `await` 关键字。

## 方法

### textAsync(label)

显示一个输入框，等待用户输入文本后执行后续操作。

**参数：**
- `label` (string) - 输入框的标题/提示文本

**返回值：**
- `Promise<string>` - 用户输入的文本内容

**示例：**

```javascript
// 获取用户输入的名称
const userName = await Input.textAsync("请输入您的名称：");
Output.log(`您好，${userName}！`);

// 获取数值输入
const age = await Input.textAsync("请输入您的年龄：");
const ageNumber = parseInt(age);
Output.log(`您的年龄是：${ageNumber}`);
```

---

### selectAsync(label, options)

显示一组选项按钮，等待用户点击选择后执行后续操作。

**参数：**
- `label` (string) - 选择器的标题
- `options` (Array<string | SelectOption>) - 选项数组，可以是字符串数组或 SelectOption 对象数组

**返回值：**
- `Promise<string>` - 用户选择的选项值

**示例：**

```javascript
// 使用字符串数组作为选项
const color = await Input.selectAsync("请选择颜色：", [
  "红色",
  "绿色",
  "蓝色"
]);
Output.log(`您选择了：${color}`);

// 使用 SelectOption 对象（带值和标签）
const action = await Input.selectAsync("请选择操作：", [
  { text: "新增数据", value: "add" },
  { text: "更新数据", value: "update" },
  { text: "删除数据", value: "delete" }
]);

if (action === "add") {
  Output.log("执行新增操作");
} else if (action === "update") {
  Output.log("执行更新操作");
} else if (action === "delete") {
  Output.log("执行删除操作");
}
```

---

### filesAsync(label, option)

显示一个文件上传按钮，等待用户上传文件后执行后续操作。

**参数：**
- `label` (string) - 文件上传按钮的描述文本
- `option` (FilesOption) - 文件上传配置选项

**返回值：**
- `Promise<any[]>` - 上传的文件数组

**FilesOption 配置：**
- `multiple` (boolean) - 是否允许多文件上传
- `fileType` (string) - 接受的文件类型，如 ".xlsx,.csv" 或 "image/*"

**示例：**

```javascript
// 上传单个 Excel 文件
const files = await Input.filesAsync("请上传 Excel 文件", {
  multiple: false,
  fileType: ".xlsx,.xls"
});

if (files && files.length > 0) {
  const file = files[0];
  Output.log(`文件名：${file.name}`);
  Output.log(`文件大小：${file.size} 字节`);
  
  // 读取文件内容
  const content = await file.text();
  // 处理文件内容...
}

// 上传多个图片文件
const imageFiles = await Input.filesAsync("请上传图片", {
  multiple: true,
  fileType: "image/*"
});

Output.log(`共上传了 ${imageFiles.length} 个图片文件`);
for (const file of imageFiles) {
  Output.log(`- ${file.name} (${file.size} 字节)`);
}

// 上传 CSV 文件并解析
const csvFiles = await Input.filesAsync("请上传 CSV 文件", {
  multiple: false,
  fileType: ".csv"
});

if (csvFiles && csvFiles.length > 0) {
  const csvContent = await csvFiles[0].text();
  const rows = csvContent.split('\n').map(row => row.split(','));
  
  // 将 CSV 数据填充到表格
  const sheet = Workbook.getActiveSheet();
  sheet.getRange(0, 0, rows.length, rows[0].length).setValues(rows);
  
  Output.log(`成功导入 ${rows.length} 行数据`);
}
```

## 使用注意事项

1. **异步调用**：所有 Input 方法都返回 Promise，必须使用 `await` 关键字，且必须在 `async` 函数中调用
2. **用户交互**：这些方法会暂停脚本执行，等待用户完成交互后才继续
3. **错误处理**：建议使用 try-catch 捕获可能的用户取消操作或其他错误
4. **文件处理**：上传的文件对象提供 `.text()`、`.arrayBuffer()` 等方法读取内容

## 完整示例

```javascript
async function importData() {
  try {
    // 选择导入方式
    const importType = await Input.selectAsync("选择导入方式：", [
      { text: "从文件导入", value: "file" },
      { text: "手动输入", value: "manual" }
    ]);
    
    const sheet = Workbook.getActiveSheet();
    
    if (importType === "file") {
      // 文件导入
      const files = await Input.filesAsync("请上传数据文件", {
        multiple: false,
        fileType: ".csv,.txt"
      });
      
      if (files && files.length > 0) {
        const content = await files[0].text();
        const rows = content.split('\n').map(row => row.split(','));
        sheet.getRange(0, 0, rows.length, rows[0].length).setValues(rows);
        Output.log(`成功导入 ${rows.length} 行数据`);
      }
    } else {
      // 手动输入
      const name = await Input.textAsync("请输入名称：");
      const value = await Input.textAsync("请输入数值：");
      
      sheet.getRange('A1').setValue(name);
      sheet.getRange('B1').setValue(value);
      Output.log("数据已添加到表格");
    }
  } catch (error) {
    Output.error(`导入失败：${error.message}`);
  }
}

await importData();
```
