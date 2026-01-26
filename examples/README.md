# 钉钉表格脚本示例

本目录包含各种钉钉表格脚本的实用示例代码。

## 示例文件列表

### [fetch-demo.js](fetch-demo.js)

展示如何使用 `fetch` API 获取网络数据并在表格中展示。

**包含的示例：**

1. **fetchJSONExample()** - 从公开API获取JSON数据
   - 获取用户列表并填充到表格
   - 设置表头样式和边框
   - 数据格式化和展示

2. **fetchTodoList()** - 获取待办事项列表
   - 条件格式化（已完成/未完成）
   - 根据状态设置不同的背景色和字体色

3. **fetchWithPostRequest()** - POST 请求示例
   - 发送 JSON 数据
   - 处理响应并展示结果

4. **fetchRandomImage()** - 获取图片信息
   - 处理包含URL的数据
   - 设置超链接样式

5. **fetchWithErrorHandling()** - 错误处理示例
   - HTTP 状态码检查
   - 完整的错误处理流程
   - 错误信息展示

6. **fetchMultipleRequests()** - 批量请求示例
   - 循环处理多个请求
   - 实时进度更新
   - 成功/失败状态标记

## 使用方法

### 方式1：直接运行

1. 打开钉钉表格
2. 点击菜单栏的「脚本」→「脚本编辑器」
3. 复制示例文件中的任意一个函数
4. 在脚本末尾调用该函数，例如：
   ```javascript
   fetchJSONExample();
   ```
5. 点击「运行」按钮

### 方式2：自定义按钮

1. 复制示例代码到脚本编辑器
2. 在钉钉表格中插入按钮
3. 将按钮关联到对应的函数
4. 点击按钮即可运行

### 方式3：定时执行

1. 复制示例代码到脚本编辑器
2. 设置定时触发器
3. 配置执行时间（如每天早上9点）
4. 脚本将自动定时运行

## 注意事项

### API 地址

示例中使用的是 JSONPlaceholder（https://jsonplaceholder.typicode.com/）这个公开的测试API。在实际使用时，请替换为您自己的API地址。

### CORS 跨域

某些API可能有跨域限制，确保您的API服务器配置了正确的CORS响应头：
```
Access-Control-Allow-Origin: *
Access-Control-Allow-Methods: GET, POST, PUT, DELETE
Access-Control-Allow-Headers: Content-Type, Authorization
```

### 请求频率

注意API的请求频率限制，避免在短时间内发送大量请求。可以在循环中添加延迟：

```javascript
// 简单的延迟函数
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// 在请求之间添加延迟
for (let i = 0; i < items.length; i++) {
  await fetch(...);
  await sleep(1000); // 延迟1秒
}
```

### 错误处理

始终使用 try-catch 包裹 fetch 请求，并提供友好的错误提示：

```javascript
try {
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`HTTP ${response.status}: ${response.statusText}`);
  }
  const data = await response.json();
  // 处理数据...
} catch (error) {
  Output.log(`错误：${error.message}`);
  // 在表格中标记错误
}
```

### 安全性

- **不要在脚本中硬编码敏感信息**（如API密钥、密码等）
- 如需使用API密钥，可以通过 `Input.textAsync()` 让用户输入
- 或者使用环境变量（如果钉钉表格支持）

## 常见问题

### Q: fetch 请求失败怎么办？

A: 检查以下几点：
1. 网络连接是否正常
2. API地址是否正确
3. API是否支持CORS跨域
4. 查看控制台的错误信息
5. 检查API是否需要认证

### Q: 如何传递认证信息？

A: 在 fetch 的 headers 中添加认证信息：
```javascript
const response = await fetch(url, {
  headers: {
    'Authorization': 'Bearer YOUR_TOKEN',
    'Content-Type': 'application/json'
  }
});
```

### Q: 如何处理大量数据？

A: 建议分批处理：
```javascript
const batchSize = 100;
for (let i = 0; i < totalData.length; i += batchSize) {
  const batch = totalData.slice(i, i + batchSize);
  // 处理这一批数据
  sheet.getRange(...).setValues(batch);
}
```

### Q: 可以上传文件吗？

A: 钉钉表格脚本的 fetch API 支持基本的 HTTP 请求，对于文件上传，可能需要使用 FormData 或 Base64 编码。

## 更多资源

- [钉钉表格脚本官方文档](https://open.dingtalk.com/document/orgapp/overview-of-dingtalk-scripts)
- [Fetch API MDN 文档](https://developer.mozilla.org/zh-CN/docs/Web/API/Fetch_API)
- [项目主 README](../README.md)
- [API 对象文档](../docs/api/)

## 贡献示例

欢迎提交更多实用的示例代码！如果您有好的使用案例，请：

1. Fork 本项目
2. 添加您的示例代码
3. 在本 README 中添加说明
4. 提交 Pull Request

---

如有问题或建议，欢迎提出 Issue！
