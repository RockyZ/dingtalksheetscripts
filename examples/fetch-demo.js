/**
 * 钉钉表格脚本 - Fetch API 使用示例
 * 
 * 本文件包含多个使用 fetch 获取网络数据的示例
 * 可以直接复制到钉钉表格脚本编辑器中运行
 */

// ============================================
// 示例1：从公开API获取数据并填充到表格
// ============================================
async function fetchJSONExample() {
  const sheet = Workbook.getActiveSheet();
  
  try {
    // 使用公开的JSONPlaceholder API作为示例
    const response = await fetch('https://jsonplaceholder.typicode.com/users');
    const users = await response.json();
    
    // 设置表头
    const headers = [["ID", "姓名", "用户名", "邮箱", "城市"]];
    sheet.getRange('A1:E1').setValues(headers);
    sheet.getRange('A1:E1').setFontWeight('bold');
    sheet.getRange('A1:E1').setBackgroundColor('#4285F4');
    sheet.getRange('A1:E1').setFontColor('#FFFFFF');
    
    // 填充数据
    const rows = users.map(user => [
      user.id,
      user.name,
      user.username,
      user.email,
      user.address.city
    ]);
    
    sheet.getRange(1, 0, rows.length, 5).setValues(rows);
    
    // 设置边框和对齐
    sheet.getRange(0, 0, rows.length + 1, 5).setBorder(
      'all', 'thin', '#E0E0E0'
    );
    
    Output.log(`✅ 成功导入 ${rows.length} 条用户数据`);
  } catch (error) {
    Output.log(`❌ 获取数据失败：${error.message}`);
  }
}

// ============================================
// 示例2：获取待办事项列表
// ============================================
async function fetchTodoList() {
  const sheet = Workbook.getActiveSheet();
  
  try {
    const response = await fetch('https://jsonplaceholder.typicode.com/todos?_limit=10');
    const todos = await response.json();
    
    // 设置表头
    sheet.getRange('A1:C1').setValues([["ID", "任务", "状态"]]);
    sheet.getRange('A1:C1').setFontWeight('bold');
    
    // 填充数据
    const rows = todos.map(todo => [
      todo.id,
      todo.title,
      todo.completed ? '✓ 已完成' : '○ 未完成'
    ]);
    
    sheet.getRange(1, 0, rows.length, 3).setValues(rows);
    
    // 根据状态设置不同颜色
    for (let i = 0; i < rows.length; i++) {
      const statusCell = sheet.getRange(`C${i + 2}`);
      if (todos[i].completed) {
        statusCell.setFontColor('#34A853'); // 绿色
        statusCell.setBackgroundColor('#E8F5E9');
      } else {
        statusCell.setFontColor('#EA4335'); // 红色
        statusCell.setBackgroundColor('#FFEBEE');
      }
    }
    
    Output.log(`✅ 成功导入 ${rows.length} 条待办事项`);
  } catch (error) {
    Output.log(`❌ 获取数据失败：${error.message}`);
  }
}

// ============================================
// 示例3：POST 请求示例
// ============================================
async function fetchWithPostRequest() {
  const sheet = Workbook.getActiveSheet();
  
  try {
    // 创建新的帖子
    const newPost = {
      title: '钉钉表格脚本测试',
      body: '这是一个使用 fetch 发送 POST 请求的示例',
      userId: 1
    };
    
    const response = await fetch('https://jsonplaceholder.typicode.com/posts', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(newPost)
    });
    
    const result = await response.json();
    
    // 在表格中显示结果
    const data = [
      ["属性", "值"],
      ["ID", result.id],
      ["标题", result.title],
      ["内容", result.body],
      ["用户ID", result.userId],
      ["状态", "创建成功 ✓"]
    ];
    
    sheet.getRange('A1:B6').setValues(data);
    sheet.getRange('A1:A6').setFontWeight('bold');
    
    Output.log('✅ POST 请求成功！');
  } catch (error) {
    Output.log(`❌ 请求失败：${error.message}`);
  }
}

// ============================================
// 示例4：获取随机图片并记录
// ============================================
async function fetchRandomImage() {
  const sheet = Workbook.getActiveSheet();
  
  try {
    const response = await fetch('https://jsonplaceholder.typicode.com/photos?_limit=5');
    const photos = await response.json();
    
    // 设置表头
    sheet.getRange('A1:D1').setValues([["ID", "标题", "缩略图URL", "原图URL"]]);
    sheet.getRange('A1:D1').setFontWeight('bold');
    sheet.getRange('A1:D1').setBackgroundColor('#FF6D00');
    sheet.getRange('A1:D1').setFontColor('#FFFFFF');
    
    // 填充数据
    const rows = photos.map(photo => [
      photo.id,
      photo.title,
      photo.thumbnailUrl,
      photo.url
    ]);
    
    sheet.getRange(1, 0, rows.length, 4).setValues(rows);
    
    // 设置URL列为超链接样式
    sheet.getRange(1, 2, rows.length, 2).setFontColor('#1A73E8');
    sheet.getRange(1, 2, rows.length, 2).setTextUnderline('single');
    
    Output.log(`✅ 成功导入 ${rows.length} 张图片信息`);
  } catch (error) {
    Output.log(`❌ 获取数据失败：${error.message}`);
  }
}

// ============================================
// 示例5：错误处理和状态检查
// ============================================
async function fetchWithErrorHandling() {
  const sheet = Workbook.getActiveSheet();
  
  try {
    const response = await fetch('https://jsonplaceholder.typicode.com/posts/1');
    
    // 检查响应状态
    if (!response.ok) {
      throw new Error(`HTTP 错误！状态码: ${response.status}`);
    }
    
    const post = await response.json();
    
    // 显示详细信息
    const info = [
      ["请求信息", ""],
      ["状态码", response.status],
      ["状态文本", response.statusText],
      ["", ""],
      ["数据内容", ""],
      ["ID", post.id],
      ["标题", post.title],
      ["内容", post.body],
      ["用户ID", post.userId]
    ];
    
    sheet.getRange('A1:B9').setValues(info);
    sheet.getRange('A1').setFontWeight('bold').setFontSize(14);
    sheet.getRange('A5').setFontWeight('bold').setFontSize(14);
    
    Output.log('✅ 请求成功，状态码：' + response.status);
  } catch (error) {
    Output.log(`❌ 错误：${error.message}`);
    
    // 在表格中记录错误
    sheet.getRange('A1:B2').setValues([
      ["错误类型", "网络请求失败"],
      ["错误信息", error.message]
    ]);
    sheet.getRange('A1:B2').setBackgroundColor('#FFEBEE');
    sheet.getRange('A1:A2').setFontWeight('bold');
  }
}

// ============================================
// 示例6：批量请求和进度显示
// ============================================
async function fetchMultipleRequests() {
  const sheet = Workbook.getActiveSheet();
  
  // 设置表头
  sheet.getRange('A1:C1').setValues([["请求编号", "状态", "数据"]]);
  sheet.getRange('A1:C1').setFontWeight('bold');
  
  const requestIds = [1, 2, 3, 4, 5];
  
  for (let i = 0; i < requestIds.length; i++) {
    try {
      Output.log(`正在获取第 ${i + 1}/${requestIds.length} 条数据...`);
      
      const response = await fetch(`https://jsonplaceholder.typicode.com/posts/${requestIds[i]}`);
      const post = await response.json();
      
      // 更新进度
      const row = i + 1;
      sheet.getRange(`A${row + 1}:C${row + 1}`).setValues([[
        requestIds[i],
        '✓ 成功',
        post.title
      ]]);
      
      sheet.getRange(`B${row + 1}`).setFontColor('#34A853');
    } catch (error) {
      const row = i + 1;
      sheet.getRange(`A${row + 1}:C${row + 1}`).setValues([[
        requestIds[i],
        '✗ 失败',
        error.message
      ]]);
      
      sheet.getRange(`B${row + 1}`).setFontColor('#EA4335');
    }
  }
  
  Output.log('✅ 批量请求完成！');
}

// ============================================
// 使用说明
// ============================================
/*
如何使用这些示例：

1. 将上面任意一个函数复制到钉钉表格脚本编辑器中
2. 在脚本最后添加调用代码，例如：
   fetchJSONExample();
3. 运行脚本即可看到效果

注意事项：
- 所有示例使用的都是公开的测试API（JSONPlaceholder）
- 实际使用时请替换为您自己的API地址
- 确保API支持CORS跨域请求
- 添加适当的错误处理
- 考虑API的请求频率限制
*/
