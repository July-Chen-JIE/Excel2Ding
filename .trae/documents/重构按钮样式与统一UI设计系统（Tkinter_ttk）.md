## 问题诊断
- Windows 默认 `vista/xpnative` 主题忽略 `ttk.Button` 的 `background`，导致按钮呈现“灰底、蓝边、白字”。
- 现有 `style.configure('TButton', background=...)` 在该主题下不生效；需要切换到支持背景绘制的主题或引入现代主题库。

## 设计方案（符合国际规范）
1. 主题与设计系统
- 启用 `clam` 主题（原生 ttk 支持背景色），集中封装“设计系统”样式（颜色、字体、间距、状态）到单一初始化函数。
- 定义语义化按钮样式：`Primary.TButton`、`Secondary.TButton`、`Danger.TButton`、`Info.TButton`，统一 `active/pressed/disabled` 映射。
- 保持色彩与 `ui_config.py` 的品牌配置一致（主色、次级色、成功/警告/危险），统一字体层级与 8px 倍数间距。

2. 组件统一
- 按钮统一套用语义样式（如“本周/本月/恢复默认”为 `Primary.TButton`；浏览类为 `Secondary.TButton`；处理为 `Primary.TButton`；退出为 `Danger.TButton`；配置为 `Info.TButton`）。
- 输入框、标签、分组框（卡片）保持卡片化白底与统一内外边距；滚动条配浅灰色统一视觉。

3. 动效与响应式
- 保留淡入动画与页签平滑过渡；为不同宽度断点调整字体（720/1024 px），兼顾可读性与空间利用。
- 维持滚动区域防抖与 Canvas 宽度自适应，减少回流重绘。

4. 兼容性与可选增强
- 若 `clam` 不可用，自动回退到当前主题并提示可选择集成 `ttkbootstrap`（更现代、跨平台一致的语义按钮样式）。
- 未来可将 `build_styles()` 迁移到 `ui_config.py` 导出 `apply_design_system(style)`，实现样式与业务彻底解耦。

## 具体改动（代码级）
- 在 `create_gui` 顶部：`style = ttk.Style(); style.theme_use('clam')`（存在时启用，否则回退）。
- 新增 `build_styles(style)`：集中定义 `Primary/Secondary/Danger/Info` 样式，并 `map` 出 `active/pressed/disabled` 背景与前景。
- 应用样式到按钮：
  - 日期快捷：`week_btn`、`month_btn`、`default_btn` 使用 `style='Primary.TButton'`（Excel2Ding_1.6_20251127.py:1166–1173）。
  - 文件浏览：`browse_input_btn`、`browse_output_btn` 使用 `style='Secondary.TButton'`（Excel2Ding_1.6_20251127.py:1124, 1131）。
  - 底部操作：`process_btn`→`Primary.TButton`，`exit_btn`→`Danger.TButton`，`config_btn`→`Info.TButton`（Excel2Ding_1.6_20251127.py:1257–1262）。
- 保持输入框占位、遮罩加载、Toast 提示等交互不变，仅提升色彩与对比度。

## 验证步骤
- 启动应用后检查按钮填充为品牌主色、文字为白色、悬停/按压态明显；不再出现灰底蓝边。
- 缩放窗口验证断点字体调整；滚动区域平滑且不卡顿。
- 跑一次处理流程，检查遮罩与进度条样式与成功提示颜色一致性。

## 用户测试与迭代
- 使用“反馈”页签收集视觉与交互体验意见；汇总 `app.log` 中的满意度与文字反馈。
- 若需要跨平台一致的更现代主题，下一步可引入 `ttkbootstrap` 并将语义样式迁移（保留 `ui_config` 颜色与命名）。

## 结构优化
- 将样式初始化独立为 `build_styles(style)`（短期）与可选 `ui_config.apply_design_system(style)`（中期）；业务逻辑与样式彻底分离，符合可维护性与规范化要求。
