# 🚀 Netlify部署前检查清单

## ✅ 部署前必检项目

### 1. 代码和配置文件检查

- [ ] **netlify.toml** 配置文件存在且正确
  - [ ] 构建命令：`npm run build:netlify`
  - [ ] 发布目录：`client/build`
  - [ ] Functions目录：`netlify/functions`
  - [ ] 重定向规则已配置
  - [ ] 环境变量已设置

- [ ] **package.json** 脚本配置正确
  - [ ] `build:netlify` 脚本存在
  - [ ] 所有必要的依赖项已安装
  - [ ] 版本号已更新

- [ ] **前端构建配置**
  - [ ] `client/craco.config.js` 优化配置已启用
  - [ ] 生产环境webpack配置已优化
  - [ ] 代码分割和压缩已启用

### 2. 数据库准备

- [ ] **Neon PostgreSQL数据库**
  - [ ] 数据库已创建
  - [ ] 连接字符串已获取
  - [ ] 数据库初始化脚本已运行
  - [ ] 表结构已创建
  - [ ] 索引已创建
  - [ ] 权限已设置

- [ ] **数据库连接测试**
  ```bash
  # 运行数据库初始化脚本
  psql your_database_url -f database/init-netlify.sql
  ```

### 3. 环境变量配置

- [ ] **必需的环境变量**
  - [ ] `NODE_ENV=production`
  - [ ] `REACT_APP_VERSION=3.0.0`
  - [ ] `REACT_APP_API_URL=/.netlify/functions`
  - [ ] `DATABASE_URL=your_neon_postgres_connection_string`
  - [ ] `NETLIFY_DATABASE_URL=your_neon_postgres_connection_string`

- [ ] **功能配置变量**
  - [ ] `REACT_APP_MAX_FILE_SIZE=10`
  - [ ] `REACT_APP_MAX_DESIGN_STEELS=1000`
  - [ ] `REACT_APP_MAX_MODULE_STEELS=100`
  - [ ] `REACT_APP_TASK_TIMEOUT=300`

### 4. API Functions检查

- [ ] **健康检查API**
  - [ ] `netlify/functions/health.js` 存在
  - [ ] 返回正确的健康状态

- [ ] **优化算法API**
  - [ ] `netlify/functions/optimize.js` 存在
  - [ ] 异步任务处理正确
  - [ ] TaskManager正确配置

- [ ] **文件上传API**
  - [ ] `netlify/functions/upload-design-steels.js` 存在
  - [ ] 文件解析功能正常
  - [ ] 智能字段识别功能正常

- [ ] **统计API**
  - [ ] `netlify/functions/stats.js` 存在
  - [ ] 数据库查询正常

- [ ] **任务管理API**
  - [ ] `netlify/functions/task.js` 存在
  - [ ] `netlify/functions/tasks.js` 存在
  - [ ] 任务状态管理正常

### 5. 性能优化检查

- [ ] **前端性能优化**
  - [ ] 代码分割已启用
  - [ ] 懒加载已配置
  - [ ] 静态资源压缩已启用
  - [ ] 缓存策略已配置

- [ ] **后端性能优化**
  - [ ] 数据库查询已优化
  - [ ] 索引已创建
  - [ ] 连接池已配置
  - [ ] 错误处理已完善

### 6. 安全配置检查

- [ ] **HTTP头部安全**
  - [ ] CORS配置正确
  - [ ] 安全头部已设置
  - [ ] XSS防护已启用

- [ ] **数据安全**
  - [ ] 数据库连接使用SSL
  - [ ] 敏感信息不在代码中
  - [ ] 环境变量正确配置

### 7. 测试和验证

- [ ] **本地构建测试**
  ```bash
  npm run deploy-check
  ```

- [ ] **功能测试**
  - [ ] 文件上传功能正常
  - [ ] 优化算法功能正常
  - [ ] 结果展示功能正常
  - [ ] 导出功能正常

## 🔧 部署前准备步骤

### 步骤1：环境检查
```bash
# 检查Node.js版本
node --version  # 应该是18+

# 检查npm版本
npm --version

# 检查项目依赖
npm audit
```

### 步骤2：本地构建测试
```bash
# 清理缓存
npm run clean

# 运行构建
npm run build:netlify

# 验证构建结果
ls -la client/build/
```

### 步骤3：环境变量准备
```bash
# 创建环境变量文件（用于批量导入）
cp NETLIFY_ENV_VARS.md netlify-env-template.txt

# 编辑文件，填入实际值
# 然后在Netlify Dashboard中配置
```

### 步骤4：数据库初始化
```bash
# 运行数据库初始化脚本
psql $DATABASE_URL -f database/init-netlify.sql
```

## 📋 部署后验证

### 立即验证
1. 访问主页：`https://your-site.netlify.app`
2. 检查健康状态：`https://your-site.netlify.app/.netlify/functions/health`
3. 查看系统统计：`https://your-site.netlify.app/.netlify/functions/stats`

### 性能验证
```bash
# 运行性能检查脚本
npm run performance-check https://your-site.netlify.app
```

### 功能验证
- [ ] 文件上传功能
- [ ] 优化算法功能
- [ ] 结果展示功能
- [ ] 数据导出功能
- [ ] 响应时间检查

## ⚠️ 常见问题解决

### 构建失败
- 检查Node.js版本是否为18+
- 确认所有依赖都已正确安装
- 查看构建日志中的具体错误

### API不工作
- 检查环境变量是否正确设置
- 确认数据库连接字符串正确
- 查看Function日志

### 数据库连接失败
- 验证DATABASE_URL格式正确
- 确认Neon数据库处于活跃状态
- 检查网络连接

### 性能问题
- 启用gzip压缩
- 检查静态资源缓存
- 优化数据库查询

## 📞 部署支持

如果遇到问题，请按以下顺序检查：

1. **查看构建日志**：在Netlify Dashboard中查看详细的构建日志
2. **检查Function日志**：查看API函数的运行日志
3. **验证环境变量**：确保所有必要的环境变量都已设置
4. **测试本地构建**：在本地运行构建命令确认无误
5. **数据库连接**：使用psql或其他工具测试数据库连接

## 🎉 部署完成后

部署成功后，请：

1. 运行完整的性能和功能测试
2. 更新项目文档中的部署URL
3. 配置域名（如果有自定义域名）
4. 设置监控和日志记录
5. 准备用户使用指南

---

**版本**: V3.0.0  
**更新时间**: 2024年  
**适用平台**: Netlify 