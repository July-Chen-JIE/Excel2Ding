
# Excel2Ding 特做数据整理工具

![应用图标](Excel2Ding.ico)  
专为飞书特做数据设计的自动化处理工具，提供一站式数据清洗解决方案  
**当前版本**：v2.1 | **最后更新**：2025-04-14 | **许可证**：MIT

---

## 📌 核心功能
1. **智能数据清洗**  
   ✅ 自动识别多种日期格式（含Excel序列号、文本日期）  
   ✅ 动态列名匹配机制，兼容字段名称变体  
   ✅ 自动过滤无效数据行，保留合规记录

2. **交互式操作界面**  
   ✅ 可视化日期范围选择（支持周/月快捷设置）  
   ✅ 实时进度条反馈处理状态  
   ✅ 自动化输出路径生成（含时间戳防覆盖）

3. **企业级数据规范**  
   ✅ 生成标准化周报格式（自动添加当前周字段）  
   ✅ 智能列宽调整（自适应内容长度）  
   ✅ 输出文件自动包含：对接人、创建时间、项目信息等7大核心字段

---

## 🚀 快速入门

### 步骤1：数据预处理
1. **获取原始数据**  
   联系飞书管理员**黄漫**，获取包含【产品设计部】与【智慧会议BG】部门的特做数据Excel文件

2. **格式调整要求**：  
   - 删除非数据相关的工作表（仅保留一个待处理Sheet）  
   - **关键步骤**：  
     ```excel
     1. 选中时间列 → 右键"设置单元格格式" → 选择"文本"
     2. 或使用数据分列功能：数据 → 分列 → 第三步选择"文本"格式
     ```
   - 清除首行说明信息（保留标准表头）

### 步骤2：环境配置
```bash
# 创建虚拟环境（推荐）
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate     # Windows

# 安装依赖
pip install -r requirements.txt
```

### 步骤3：运行工具
![GUI界面演示](gui_preview.png)  
1. 双击运行 `特做数据自动整理.exe`
2. 按照GUI指引：  
   - 选择输入文件（支持.xls/.xlsx）  
   - 设置时间范围
   - 指定输出路径（默认为输入路径）  
3. 点击【开始处理】观察进度条状态

---

## ⚙️ 技术架构
### 数据处理流程
[![](https://mermaid.ink/img/pako:eNplkctOwkAUhl-lOWs0SlsqXZgooOLalS2Lxo5ilJZUSFQg8V6iqBgvxGjEC0ZjFFwQowL6Ms50-hYOtcaFszon__fPn3NODqZMHYEMM5aWTnITUdXg2BtS8F4V3-3EFqfQfILr6RnkhnPu9QaubTjlLfehRBurhR90uKvmyWUbt_fJ041zVc9zEQVv35OVVVys4NKbu7mb-M_i9hp-fc1zUcU9OqWNBtkvO7V3H4x4kTGFVG7JefXHQDuPZPfaB2IeMMKAF7fSpKV1fNaknzZp1XxgxANGFXxwT46faf3KqVd8adSTxhRyaWN7C5dO6Mchtn-jxzw1rlD7gQ3RnaDeoc_r5LjpA3EPGFecoyoplp3WIbk4Jyf2V-slAQG2x1kd5IyVRQFIISuldVvIda0qZJIohVSQWalr1pwKqlFgnrRmTJpm6tdmmdmZJMjT2vwC67JpXcug6KzGLvSHIENHVsTMGhmQ-8PeFyDnYBFknpd6gyGRF8L8gCD0B_kALIEs8r0DkhCWgkEh1CeKkhQqBGDZC-1jilj4BrpE2v4?type=png)](https://mermaid-live.nodejs.cn/edit#pako:eNplkctOwkAUhl-lOWs0SlsqXZgooOLalS2Lxo5ilJZUSFQg8V6iqBgvxGjEC0ZjFFwQowL6Ms50-hYOtcaFszon__fPn3NODqZMHYEMM5aWTnITUdXg2BtS8F4V3-3EFqfQfILr6RnkhnPu9QaubTjlLfehRBurhR90uKvmyWUbt_fJ041zVc9zEQVv35OVVVys4NKbu7mb-M_i9hp-fc1zUcU9OqWNBtkvO7V3H4x4kTGFVG7JefXHQDuPZPfaB2IeMMKAF7fSpKV1fNaknzZp1XxgxANGFXxwT46faf3KqVd8adSTxhRyaWN7C5dO6Mchtn-jxzw1rlD7gQ3RnaDeoc_r5LjpA3EPGFecoyoplp3WIbk4Jyf2V-slAQG2x1kd5IyVRQFIISuldVvIda0qZJIohVSQWalr1pwKqlFgnrRmTJpm6tdmmdmZJMjT2vwC67JpXcug6KzGLvSHIENHVsTMGhmQ-8PeFyDnYBFknpd6gyGRF8L8gCD0B_kALIEs8r0DkhCWgkEh1CeKkhQqBGDZC-1jilj4BrpE2v4)

<!--
    graph TD
    A[原始Excel] --- B{预处理验证}
    B ---|格式正确| C[动态列匹配]
    B ---|格式异常| D[错误提示]
    C --- E[日期格式转换]
    E --- F[时间范围过滤]
    F --- G[周数计算]
    G --- H[标准化输出]
    H --- I[自动列宽调整]
    I --- J[生成结果文件]
-->



### 关键技术栈
| 模块          | 技术实现                          | 相关文档                         |
|---------------|----------------------------------|---------------------------------|
| 列名清洗       | 正则表达式替换特殊字符            |  数据清洗规范             |
| 日期解析       | Excel序列号转日期算法             |  日期转换方法             |
| 进度反馈       | Tkinter多线程+进度条             |  GUI开发指南              |
| 文件输出       | Openpyxl格式控制                 |  报表生成专利             |

---

## ⚠️ 重要注意事项
1. **文件规范**  
   - 输入文件必须包含名为`Sheet1`的工作表  
   - 时间戳列必须为第3列（C列）且为文本格式

2. **异常处理**  
   - 遇到`列名不匹配`错误时检查[Excel2Ding.py]配置  

3. **性能优化**  
   - 单文件处理建议不超过50万行  
   - 大文件处理时关闭其他Excel进程

---

## 📜 版本更新

### v2.1 [2025-04-14]
- 新增列映射配置功能（支持动态添加/编辑/删除）
- 优化代码架构，引入 ColumnMapper 类管理配置
- 改进UI视觉效果，统一应用配色方案
- 支持产品线过滤和批量更换对接人
- 添加配置持久化存储（JSON格式）
- 完善错误处理和用户提示

### v2.0 [2025-04-11]
- 新增多线程处理加速（性能提升40%）
- 支持.xls与.xlsx格式混合输入
- 修复时间戳精度丢失问题
- 优化UI界面，采用扁平化设计风格

### v1.5 [2025-02-22]
- 首次发布基础版本
- 实现核心清洗逻辑
- 开发GUI操作界面
---

### 下版本优化
   简化一下必要操作
   - 删除非数据相关的工作表（仅保留一个待处理Sheet）  
   - **关键步骤**：  
     ```excel
     1. 选中时间列 → 右键"设置单元格格式" → 选择"文本"
     2. 或使用数据分列功能：数据 → 分列 → 第三步选择"文本"格式
     ```
   - 清除首行说明信息（保留标准表头）

## 🤝 参与贡献
欢迎通过以下方式参与：
1. 联系开发者czj