# 钢材优化系统V3.0 - 导出功能使用指南

## 功能概述

钢材优化系统V3.0现已支持完整的导出功能，可以将优化结果导出为Excel和HTML格式的报告，方便用户保存、分享和打印优化结果。

## 支持的导出格式

### 1. Excel采购清单 (.xlsx)
- **主要用途**：钢材采购和成本核算
- **内容包含**：
  - 详细采购清单（规格、数量、单价、总价）
  - 材料利用率分析
  - 优化结果汇总信息
- **特点**：
  - 双工作表结构：采购清单 + 优化信息
  - 包含成本估算和利用率计算
  - 格式化表格，支持进一步的数据处理
  - 适合采购部门和财务核算

### 2. HTML完整报告 (.html) - (PDF替代方案)
- **主要用途**：技术分析、方案汇报和打印
- **内容包含**：
  - 优化结果总览
  - 钢材采购清单
  - 完整的设计钢材清单
  - 技术说明和使用指南
- **特点**：
  - **完美中文支持**：通过生成HTML，从根本上解决了PDF中文乱码问题。
  - **浏览器原生打印**：用户可在浏览器中打开，利用其强大的打印功能（Ctrl+P）另存为PDF。
  - **格式灵活**：内容适应浏览器窗口，查看方便。
  - **内容详尽**：包含所有技术细节，适合工程师和管理人员。

## 使用方法

### 前端界面操作

1. **执行优化**：
   - 在优化页面完成钢材优化计算
   - 等待优化完成并跳转到结果页面

2. **导出报告**：
   - 在结果页面底部找到导出按钮
   - 选择"导出采购清单(Excel)"获取详细的采购信息
   - 选择"导出完整报告(PDF)"将生成并下载一个HTML文件

3. **文件保存与使用**：
   - 浏览器会自动下载生成的文件
   - Excel文件名：`钢材优化报告_YYYY-MM-DDTHH-MM-SS.xlsx`
   - **HTML文件名**：`钢材优化报告_YYYY-MM-DDTHH-MM-SS.html`
   - **获取PDF**：双击打开下载的HTML文件，在浏览器中使用 `Ctrl + P` 打印功能，选择"另存为PDF"。

### API调用方式

#### Excel采购清单导出API
```javascript
// POST /api/export/excel - 导出采购清单
const response = await fetch('/api/export/excel', {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
  },
  body: JSON.stringify({
    optimizationResult: results,
    exportOptions: {
      format: 'excel',
      includeDetails: true,
      customTitle: '钢材采购清单'
    }
  }),
});
```

#### HTML报告导出API (PDF)
```javascript
// POST /api/export/pdf - 生成并下载HTML报告
const response = await fetch('/api/export/pdf', {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
  },
  body: JSON.stringify({
    optimizationResult: results,
    exportOptions: {
      format: 'pdf', // 此参数保留，但后端已硬编码为HTML生成
      includeDetails: true,
      customTitle: '钢材优化完整报告'
    },
    designSteels: designSteels // 别忘了传入设计钢材数据
  }),
});

// 响应体为一个JSON对象，包含下载链接
const result = await response.json();
// {
//   "success": true,
//   "filename": "钢材优化报告_....html",
//   "downloadUrl": "/api/download/钢材优化报告_....html",
//   "message": "..."
// }
```

## 导出选项配置

### ExportOptions参数说明

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `format` | string | 'excel' | 导出格式：'excel' 或 'pdf'。**注意**：'pdf'选项现在会触发HTML文件生成。 |
| `includeCharts` | boolean | false | 是否包含图表（暂未实现） |
| `includeDetails` | boolean | true | 是否包含详细切割方案 |
| `includeLossRateBreakdown` | boolean | true | 是否包含损耗率分解 |
| `customTitle` | string | undefined | 自定义报告标题 |

## 文件内容详解

### Excel采购清单结构

#### 工作表1：钢材采购清单
- 序号、模数钢材规格、单根长度、采购数量
- 总长度、材料利用率、总金额、备注信息
- 汇总行：总采购成本和整体利用率
- 格式化表格，包含成本估算

#### 工作表2：优化信息
- 优化指标汇总：损耗率、利用率、执行时间等
- 系统生成时间和版本信息

### HTML完整报告结构

1. **优化结果汇总**：关键指标和统计数据。
2. **模数钢材采购清单**：表格化的采购信息。
3. **设计钢材清单**：完整的原始需求列表。
4. **技术说明**：对优化算法、损耗率计算等的解释。
5. **使用指南**：包含如何将HTML打印为PDF的说明。

## 技术实现

### 服务器端依赖
- **ExcelJS**：用于生成Excel文件。
- **fs, path**：Node.js内置模块，用于生成和管理HTML文件。

### 前端实现
- 使用fetch API调用导出接口。
- **对于HTML报告**，解析后端返回的JSON，并动态创建 `<a>` 标签来触发下载。
- 错误处理和用户反馈。

## 错误处理

### 常见错误及解决方案

1. **"缺少优化结果数据"**
   - 确保已完成优化计算
   - 检查优化结果是否正常

2. **"Excel导出失败" / "PDF导出失败"**
   - 检查服务器是否正常运行
   - 查看服务器日志获取详细错误信息

3. **文件下载失败或内容为空**
   - 检查浏览器设置，清除缓存。
   - 确保后端 `server/uploads` 目录有写入权限。
   - 确认前后端API响应格式匹配。

## 性能优化

### 文件大小优化
- Excel文件：通过合理的数据结构减少文件大小
- HTML文件：使用简洁的布局和字体

### 生成速度优化
- 异步处理大量数据
- 流式写入减少内存占用

## 未来计划

1. **图表支持**：在Excel和HTML报告中添加数据图表
2. **模板定制**：支持用户自定义报告模板
3. **批量导出**：支持多个优化结果的批量导出
4. **云存储集成**：支持直接保存到云存储服务

## 更新日志

### v3.0.1 (2025-06-23) - PDF功能重构
- 🚀 **重大重构**：将PDF导出功能从 `jsPDF` 库完全迁移到 **HTML文件生成方案**。
- ✅ **解决中文乱码**：通过生成HTML，完美解决了PDF导出中的中文乱码问题。
- ✨ **简化流程**：用户现在下载HTML文件，并在浏览器中使用打印功能（Ctrl+P）来生成PDF。
- 🗑️ **移除依赖**：删除了 `jsPDF` 和 `jsPDF-AutoTable` 库，简化了后端依赖。
- 📝 **更新文档**：全面更新了与导出功能相关的技术和使用文档。

### v3.0.0 (2025-01-23)
- ✅ 实现Excel采购清单导出功能
- ✅ 实现PDF完整报告导出功能 (基于jsPDF)
- ✅ 明确功能分工：Excel专注采购，PDF专注技术分析
- ✅ 添加成本估算和利用率计算
- ✅ 集成前端导出按钮并更新文案
- ✅ 支持自定义导出选项
- ✅ 完善错误处理机制

---

如有问题或建议，请联系开发团队。 