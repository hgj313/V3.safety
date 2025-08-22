# Netlify环境变量配置指南

## 🔧 必需环境变量

在Netlify Dashboard的Site settings → Environment variables中添加以下变量：

### 基础配置
```
NODE_ENV=production
REACT_APP_VERSION=3.0.0
REACT_APP_NAME=钢材采购优化系统V3.0
REACT_APP_API_URL=/.netlify/functions
```

### 数据库配置
```
DATABASE_URL=your_neon_postgres_connection_string
NETLIFY_DATABASE_URL=your_neon_postgres_connection_string
```

### 功能开关
```
REACT_APP_DEBUG=false
REACT_APP_PERFORMANCE_MONITORING=true
REACT_APP_ERROR_REPORTING=true
```

### 文件处理配置
```
REACT_APP_MAX_FILE_SIZE=10
REACT_APP_SUPPORTED_FILE_TYPES=.xlsx,.xls,.csv
```

### 系统限制
```
REACT_APP_MAX_DESIGN_STEELS=1000
REACT_APP_MAX_MODULE_STEELS=100
REACT_APP_TASK_TIMEOUT=300
```

### 缓存配置
```
REACT_APP_STORAGE_PREFIX=hgj_v3_
REACT_APP_CACHE_EXPIRE_HOURS=24
```

### 部署信息
```
REACT_APP_DEPLOY_ENV=production
REACT_APP_BUILD_TIME=
REACT_APP_GIT_HASH=
```

## 📋 配置说明

### 1. 数据库配置
- 需要在Neon PostgreSQL创建数据库
- 获取连接字符串并填入DATABASE_URL
- NETLIFY_DATABASE_URL与DATABASE_URL相同

### 2. 功能开关
- REACT_APP_DEBUG：生产环境建议设为false
- REACT_APP_PERFORMANCE_MONITORING：启用性能监控
- REACT_APP_ERROR_REPORTING：启用错误报告

### 3. 文件处理
- REACT_APP_MAX_FILE_SIZE：最大文件上传大小(MB)
- REACT_APP_SUPPORTED_FILE_TYPES：支持的文件类型

### 4. 系统限制
- REACT_APP_MAX_DESIGN_STEELS：最大设计钢材数量
- REACT_APP_MAX_MODULE_STEELS：最大模数钢材数量
- REACT_APP_TASK_TIMEOUT：任务超时时间(秒)

## 🚀 快速配置

### 方法1：通过Netlify UI配置
1. 登录Netlify Dashboard
2. 选择你的站点
3. 进入Site settings → Environment variables
4. 点击Add variable逐个添加上述变量

### 方法2：通过Netlify CLI配置
```bash
# 安装Netlify CLI
npm install -g netlify-cli

# 登录
netlify login

# 进入项目目录
cd your-project

# 配置环境变量
netlify env:set NODE_ENV production
netlify env:set REACT_APP_VERSION 3.0.0
netlify env:set REACT_APP_API_URL /.netlify/functions
# ... 继续添加其他变量
```

### 方法3：通过环境变量文件批量导入
创建一个临时的环境变量文件 `netlify-env.txt`：
```
NODE_ENV=production
REACT_APP_VERSION=3.0.0
REACT_APP_API_URL=/.netlify/functions
DATABASE_URL=your_connection_string
...
```

然后使用Netlify CLI导入：
```bash
netlify env:import netlify-env.txt
```

## ⚠️ 注意事项

1. **敏感信息**：DATABASE_URL等敏感信息请务必保密
2. **前端变量**：只有REACT_APP_开头的变量才能在前端代码中使用
3. **构建时变量**：NODE_ENV等变量在构建时生效
4. **运行时变量**：Functions中的变量在运行时生效

## 🔍 验证配置

部署后可以通过以下方式验证环境变量配置：

1. 访问健康检查端点：`/.netlify/functions/health`
2. 查看浏览器控制台是否有环境变量相关错误
3. 检查Functions日志是否有数据库连接错误

## 📞 故障排查

如果遇到环境变量相关问题：

1. 确保变量名拼写正确
2. 检查变量值是否有特殊字符需要转义
3. 确认前端变量是否以REACT_APP_开头
4. 查看Netlify Functions日志获取详细错误信息 