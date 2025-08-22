# Netlify 异步任务系统部署指南

## 🎯 系统概述

本项目现已实现完整的异步任务系统，解决了原有的"无限轮询"问题，提供真正的任务队列和状态管理。

### 核心特性

- ✅ **真正的异步处理**：立即返回taskId，后台执行优化
- ✅ **持久化存储**：任务状态存储在文件系统中，重启不丢失
- ✅ **实时进度跟踪**：支持进度条和状态消息
- ✅ **自动清理**：24小时后自动清理过期任务
- ✅ **错误处理**：完整的错误捕获和报告机制
- ✅ **任务管理**：支持查询、取消、历史记录

## 📁 文件结构

```
netlify/functions/
├── optimize.js          # 异步优化接口（已修改）
├── task.js             # 单个任务查询接口（新增）
├── tasks.js            # 任务管理接口（已更新）
└── utils/
    └── TaskManager.js  # 任务管理器核心类（新增）
```

## 🔧 API 接口

### 1. 提交优化任务
```http
POST /api/optimize
Content-Type: application/json

{
  "designSteels": [...],
  "moduleSteels": [...],
  "constraints": {...}
}
```

**响应:**
```json
{
  "success": true,
  "taskId": "task_1703123456789_1",
  "message": "优化任务已创建，请通过taskId查询进度",
  "status": "pending"
}
```

### 2. 查询任务状态
```http
GET /api/task/{taskId}
```

**响应:**
```json
{
  "success": true,
  "taskId": "task_1703123456789_1",
  "status": "running",
  "progress": 60,
  "message": "正在计算最优切割方案...",
  "executionTime": 5000,
  "createdAt": "2023-12-21T10:30:00.000Z",
  "updatedAt": "2023-12-21T10:30:05.000Z",
  "results": null
}
```

### 3. 获取任务列表
```http
GET /api/tasks?limit=20&status=completed
```

### 4. 取消任务
```http
DELETE /api/task/{taskId}
```

## 🚀 部署步骤

### 1. 确保文件已创建
确认以下文件已正确创建：
- `netlify/functions/utils/TaskManager.js`
- `netlify/functions/task.js`
- 已修改的 `netlify/functions/optimize.js`
- 已修改的 `netlify/functions/tasks.js`
- 已修改的 `netlify.toml`

### 2. 本地测试
```bash
# 运行测试脚本
node test-async-system.js
```

### 3. 部署到Netlify
```bash
# 提交代码
git add .
git commit -m "实现异步任务系统"
git push origin main

# 或者手动部署
npm run build:netlify
```

### 4. 验证部署
访问以下端点验证系统正常工作：
- `https://your-site.netlify.app/api/health`
- `https://your-site.netlify.app/api/tasks`

## 📊 任务状态流程

```
idle → pending → running → completed
                    ↓
                  failed
                    ↓
                cancelled
```

### 状态说明
- **idle**: 初始状态，未提交任务
- **pending**: 任务已创建，等待处理
- **running**: 任务正在执行，可查看进度
- **completed**: 任务完成，可获取结果
- **failed**: 任务失败，查看错误信息
- **cancelled**: 任务被用户取消

## 🔍 前端集成

前端已有完整的异步处理逻辑，位于 `client/src/hooks/useOptimizationResults.ts`：

```typescript
import { useAsyncOptimization } from './hooks/useOptimizationResults';

function OptimizationPage() {
  const {
    currentTask,
    isPolling,
    submitOptimization,
    cancelTask,
    resetTask,
    isActive,
    hasResults
  } = useAsyncOptimization();

  const handleOptimize = async () => {
    const result = await submitOptimization(optimizationData);
    if (result.success) {
      console.log('任务已提交:', result.taskId);
    }
  };

  return (
    <div>
      {isActive && (
        <div>
          <Progress percent={currentTask.progress} />
          <p>{currentTask.message}</p>
          <Button onClick={cancelTask}>取消任务</Button>
        </div>
      )}
      
      {hasResults && (
        <ResultsDisplay results={currentTask.results} />
      )}
    </div>
  );
}
```

## 🛠️ 故障排查

### 常见问题

1. **任务一直处于pending状态**
   - 检查Netlify Functions日志
   - 确认OptimizationService是否正确导入

2. **文件权限错误**
   - Netlify Functions使用 `/tmp` 目录存储任务数据
   - 确保TaskManager使用正确的文件路径

3. **任务丢失**
   - 任务存储在临时文件中，Netlify重新部署时会清空
   - 这是正常行为，不影响新任务

4. **轮询404错误**
   - 检查 `netlify.toml` 路由配置
   - 确认 `/api/task/:taskId` 指向正确的函数

### 调试命令

```bash
# 查看Netlify Functions日志
netlify dev

# 本地测试任务管理器
node -e "
const TaskManager = require('./netlify/functions/utils/TaskManager');
const tm = new TaskManager();
tm.createOptimizationTask({test: true}).then(console.log);
"
```

## 📈 性能优化建议

1. **任务清理策略**
   - 当前设置24小时自动清理
   - 可根据需要调整 `maxTaskAge` 参数

2. **轮询频率**
   - 前端默认2秒轮询一次
   - 可根据任务复杂度调整频率

3. **存储优化**
   - 考虑使用外部数据库（Redis/MongoDB）
   - 实现任务结果压缩存储

## 🔐 安全考虑

1. **任务ID安全**
   - 使用时间戳+计数器生成，难以猜测
   - 可考虑添加UUID提高安全性

2. **访问控制**
   - 当前无身份验证，所有人都可查询任务
   - 生产环境建议添加认证机制

3. **资源限制**
   - Netlify Functions有执行时间限制
   - 大型优化任务可能需要分片处理

## 📞 技术支持

如遇到问题，请检查：
1. Netlify Functions 部署日志
2. 浏览器开发者工具网络面板
3. 本地测试脚本输出

---

**🎉 恭喜！您的异步任务系统已成功部署！** 