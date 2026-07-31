# 参与贡献

感谢你愿意改进 WPS 演示自动填色。提交代码前，请先搜索现有 Issue，避免重复工作。较大的功能或算法调整建议先创建 Issue，说明使用场景、预期行为和实现思路。

## 本地开发

1. Fork 并克隆仓库。
2. 安装 Node.js 和 WPS 演示 Windows 桌面版。
3. 安装依赖：

   ```powershell
   npm install
   ```

4. 运行自动化测试：

   ```powershell
   npm test
   ```

5. 生成生产构建：

   ```powershell
   npm run build
   ```

6. 如需在 WPS 中联调，运行：

   ```powershell
   wpsjs debug
   ```

## 提交要求

- 一个 Pull Request 尽量只解决一个明确问题。
- 保持现有代码风格，不提交 `node_modules/`、`dist/`、日志或测试演示文稿。
- 涉及几何算法的修改，请补充或更新 `tests/geometry.test.js`。
- 涉及 WPS API 交互的修改，请补充或更新 `tests/wps-mock.test.js`。
- 修复缺陷时，请说明复现步骤、根因和验证方法；如果涉及界面，请附截图。
- 提交前确保 `npm test` 和 `npm run build` 均通过。

## Commit 与 Pull Request

Commit 信息建议使用简洁的祈使句，例如：

```text
fix: correct ellipse arc direction
feat: add reusable color swatches
docs: clarify network selection workflow
```

Pull Request 应填写模板中的变更说明、验证结果和检查清单。提交即表示你同意以本项目的 [MIT License](LICENSE) 发布你的贡献。

## 行为规范

参与项目即表示你同意遵守 [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md)。

