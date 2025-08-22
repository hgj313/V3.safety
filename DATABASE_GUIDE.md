# 钢材优化系统数据库使用指南

## 📖 概述

钢材优化系统V3.0现在集成了轻量级JSON数据库，用于持久化存储系统数据。我们选择了**lowdb**作为数据库解决方案，这是一个基于JSON文件的轻量级数据库。

## 🎯 为什么选择JSON数据库？

### 优点
- ✅ **零配置**：无需安装数据库服务器
- ✅ **跨平台**：在Windows、Mac、Linux上都能正常工作
- ✅ **易于调试**：数据以JSON格式存储，可直接查看和编辑
- ✅ **无编译问题**：避免了SQLite在Windows上的编译问题
- ✅ **轻量级**：适合中小型应用
- ✅ **备份简单**：只需复制JSON文件即可

### 缺点
- ❌ 不适合大量数据（建议<10MB）
- ❌ 不支持复杂的SQL查询
- ❌ 并发性能有限

## 📁 数据库文件结构

```
database/
├── Database.js          # 数据库管理类
├── schema.sql          # SQL表结构（参考用）
├── steel_system.json   # 主数据库文件
└── backups/           # 自动备份目录
    ├── backup_2024-01-20T10-30-00-000Z.json
    ├── backup_2024-01-20T11-00-00-000Z.json
    └── ...
```

## 🗃️ 数据结构

### 主要数据表

#### 1. 设计钢材 (designSteels)
```json
{
  "id": "steel_1705731234567_0",
  "displayId": "D001",
  "length": 3000,
  "quantity": 10,
  "crossSection": 2000,
  "componentNumber": "GJ-001",
  "specification": "HRB400",
  "partNumber": "P-001",
  "createdAt": "2024-01-20T10:30:00.000Z",
  "updatedAt": "2024-01-20T10:30:00.000Z"
}
```

#### 2. 模数钢材 (moduleSteels)
```json
{
  "id": "default_1",
  "name": "12米标准钢材",
  "length": 12000,
  "createdAt": "2024-01-20T10:30:00.000Z",
  "updatedAt": "2024-01-20T10:30:00.000Z"
}
```

#### 3. 优化任务 (optimizationTasks)
```json
{
  "id": "opt_1705731234567_abc123",
  "status": "completed",
  "progress": 100,
  "wasteThreshold": 100,
  "targetLossRate": 5,
  "timeLimit": 30,
  "maxWeldingSegments": 1,
  "results": { /* 优化结果JSON */ },
  "executionTime": 2850,
  "createdAt": "2024-01-20T10:30:00.000Z",
  "updatedAt": "2024-01-20T10:30:00.000Z"
}
```

#### 4. 系统统计 (systemStats)
```json
{
  "totalOptimizations": 0,
  "totalDesignSteels": 0,
  "totalModuleSteels": 3,
  "totalSavedCost": 0,
  "lastUpdated": "2024-01-20T10:30:00.000Z"
}
```

#### 5. 操作日志 (operationLogs)
```json
{
  "id": "1705731234567",
  "operationType": "upload",
  "description": "上传设计钢材文件",
  "details": { /* 详细信息 */ },
  "ipAddress": "127.0.0.1",
  "userAgent": "Mozilla/5.0...",
  "createdAt": "2024-01-20T10:30:00.000Z"
}
```

## 🔧 数据库操作

### 基本操作

#### 初始化数据库
```javascript
const databaseManager = require('./database/Database');
await databaseManager.init();
```

#### 保存设计钢材
```javascript
const steels = [/* 钢材数组 */];
await databaseManager.saveDesignSteels(steels);
```

#### 获取设计钢材
```javascript
const steels = databaseManager.getDesignSteels();
```

#### 记录操作日志
```javascript
await databaseManager.logOperation(
  'upload',
  '上传设计钢材文件',
  { fileName: 'test.xlsx', count: 10 },
  req
);
```

### 数据备份

#### 手动备份
```javascript
const backupPath = await databaseManager.backup();
console.log('备份文件:', backupPath);
```

#### 自动备份
- 系统会在每次重要操作时自动备份
- 保留最新10个备份文件
- 备份文件命名格式：`backup_YYYY-MM-DDTHH-mm-ss-sssZ.json`

### 数据导出/导入

#### 导出数据
```javascript
const exportData = await databaseManager.exportData();
// 可以保存为JSON文件或发送给其他系统
```

#### 导入数据
```javascript
const importData = { /* 导出的数据格式 */ };
const success = await databaseManager.importData(importData);
```

## 📊 数据库监控

### 获取统计信息
```javascript
const stats = databaseManager.getStats();
console.log('数据库统计:', stats);
```

输出示例：
```json
{
  "designSteels": 156,
  "moduleSteels": 3,
  "optimizationTasks": 12,
  "completedTasks": 10,
  "operationLogs": 45,
  "databaseSize": "125.67 KB",
  "lastUpdated": "2024-01-20T10:30:00.000Z"
}
```

## 🛠️ 维护操作

### 1. 查看数据库文件
数据库文件位于 `database/steel_system.json`，可以用任何文本编辑器打开查看。

### 2. 手动编辑数据
⚠️ **注意**：直接编辑JSON文件时请确保：
- JSON格式正确
- 备份原文件
- 重启服务器以重新加载数据

### 3. 清理日志
系统会自动限制日志数量（默认1000条），但可以手动清理：
```javascript
// 清空操作日志
databaseManager.db.data.operationLogs = [];
await databaseManager.save();
```

### 4. 重置数据库
如果需要重置所有数据：
```bash
# 删除数据库文件
rm database/steel_system.json

# 重启服务器，系统会自动创建新的数据库
npm start
```

## 🚀 性能优化建议

### 1. 定期备份
- 建议每天手动备份一次重要数据
- 可以设置定时任务自动备份

### 2. 数据清理
- 定期清理旧的优化任务记录
- 限制操作日志数量

### 3. 文件大小监控
- 当JSON文件超过5MB时考虑数据清理
- 超过10MB时可能影响性能

## 🔄 数据迁移

### 从内存模式迁移到数据库
如果之前使用的是内存模式，现在想保存数据：

1. 确保数据库已初始化
2. 将内存中的数据保存到数据库：
```javascript
// 保存设计钢材
await databaseManager.saveDesignSteels(designSteels);

// 保存模数钢材
await databaseManager.saveModuleSteels(moduleSteels);
```

### 升级到真正的数据库
如果将来需要升级到PostgreSQL或MySQL：

1. 导出当前数据：
```javascript
const data = await databaseManager.exportData();
```

2. 创建SQL脚本将JSON数据转换为SQL INSERT语句
3. 在新数据库中导入数据

## 🐛 故障排除

### 常见问题

#### 1. 数据库文件损坏
```bash
# 从备份恢复
cp database/backups/backup_最新时间.json database/steel_system.json
```

#### 2. 权限问题
```bash
# 确保数据库目录有写权限
chmod 755 database/
chmod 644 database/steel_system.json
```

#### 3. JSON格式错误
使用JSON验证工具检查文件格式，或从备份恢复。

#### 4. 性能问题
- 检查文件大小
- 清理旧数据
- 考虑升级到真正的数据库

## 📞 技术支持

如果遇到数据库相关问题：

1. 查看服务器日志
2. 检查数据库文件是否存在
3. 验证JSON格式
4. 尝试从备份恢复
5. 联系技术支持团队

---

**注意**：这是一个轻量级解决方案，适合中小型应用。如果数据量增长很大，建议升级到专业数据库系统。 