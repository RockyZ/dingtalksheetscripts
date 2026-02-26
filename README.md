# dingtalk-sheet-scripts

钉钉表格脚本开发参考仓库，包含：

- `SKILL.md`：可被 Agent 直接加载的技能定义（已 SKILL 化）
- `docs/api/`：钉钉表格脚本 API 参考文档
- `examples/`：可直接复用的脚本示例
- `scripts/update_api_docs.py`：API 文档更新脚本

## 安装技能

本仓库已经 SKILL 化，可通过以下命令安装：

```bash
npx skills add RockyZ/dingtalk-sheet-scripts
```

## 使用说明

1. 安装后在 Agent 中启用 `dingtalk-sheet-scripts` 技能
2. 按需查阅 `docs/api/` 中的对象文档（如 `Workbook`、`Sheet`、`Range`）
3. 参考 `examples/fetch-demo.js` 快速开始编写脚本

## 相关文档

- API 索引：`docs/api/README.md`
- 示例说明：`examples/README.md`
- 技能定义：`SKILL.md`
