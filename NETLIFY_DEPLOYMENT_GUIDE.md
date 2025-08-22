# 钢材采购优化系统 V3.0 - Netlify部署指南

## 📋 部署前准备

### 1. 确保系统环境
- Node.js 18+ 
- npm 或 yarn
- Git
- Netlify CLI (可选)

### 2. 检查项目文件
确保以下文件存在且配置正确：
- `netlify.toml` - Netlify配置文件 ✅
- `package.json` - 已添加构建脚本 ✅
- `netlify/functions/` - API函数目录 ✅
- `.env.example` - 环境变量示例 ✅

## 🚀 部署方式

### 方式一：通过Netlify网站部署 (推荐)

1. **登录Netlify**
   - 访问 [https://www.netlify.com](https://www.netlify.com)
   - 使用GitHub/GitLab账号登录

2. **连接代码仓库**
   - 点击 "New site from Git"
   - 选择你的Git仓库提供商
   - 选择项目仓库

3. **配置构建设置**
   ```
   Build command: npm run build:netlify
   Publish directory: client/build
   Functions directory: netlify/functions
   ```

4. **设置环境变量**
   - 在Site settings > Environment variables 中添加：
   ```
   NODE_ENV=production
   REACT_APP_VERSION=3.0.0
   REACT_APP_API_URL=/.netlify/functions
   ```

5. **部署**
   - 点击 "Deploy site"
   - 等待构建完成

### 方式二：使用Netlify CLI部署

1. **安装Netlify CLI**
   ```bash
   npm install -g netlify-cli
   ```

2. **登录Netlify**
   ```bash
   netlify login
   ```

3. **初始化项目**
   ```bash
   netlify init
   ```

4. **构建项目**
   ```bash
   npm run build:netlify
   ```

5. **部署**
   ```bash
   # 部署到预览环境
   netlify deploy
   
   # 部署到生产环境
   netlify deploy --prod
   ```

### 方式三：通过拖拽部署

1. **本地构建**
   ```bash
   npm run build:netlify
   ```

2. **拖拽部署**
   - 访问 [https://app.netlify.com/drop](https://app.netlify.com/drop)
   - 将 `client/build` 文件夹拖拽到页面

## 🔧 配置说明

### netlify.toml 关键配置

```toml
[build]
  command = "npm run build:netlify"
  functions = "netlify/functions"
  publish = "client/build"

[[redirects]]
  from = "/api/*"
  to = "/.netlify/functions/:splat"
  status = 200
```

### 环境变量配置

在Netlify Dashboard中设置以下环境变量：

| 变量名 | 值 | 说明 |
|--------|-----|------|
| NODE_ENV | production | 生产环境 |
| REACT_APP_VERSION | 3.0.0 | 应用版本 |
| REACT_APP_API_URL | /.netlify/functions | API地址 |

## 📱 API端点说明

部署后，API端点将变为：

| 原端点 | Netlify端点 | 功能 |
|--------|-------------|------|
| /api/health | /.netlify/functions/health | 健康检查 |
| /api/optimize | /.netlify/functions/optimize | 优化算法 |
| /api/upload-design-steels | /.netlify/functions/upload-design-steels | 文件上传 |
| /api/validate-constraints | /.netlify/functions/validate-constraints | 约束验证 |
| /api/stats | /.netlify/functions/stats | 系统统计 |

## 🔍 测试部署

部署完成后进行以下测试：

1. **访问主页**
   ```
   https://your-site-name.netlify.app
   ```

2. **测试API**
   ```bash
   # 健康检查
   curl https://your-site-name.netlify.app/.netlify/functions/health
   
   # 系统统计
   curl https://your-site-name.netlify.app/.netlify/functions/stats
   ```

3. **功能测试**
   - 文件上传功能
   - 优化算法功能
   - 结果导出功能

## ⚠️ 注意事项

### 1. 函数限制
- Netlify Functions有执行时间限制（免费版10秒，付费版15分钟）
- 复杂优化算法可能需要优化或分批处理

### 2. 文件上传
- Netlify Functions目前最适合处理base64编码的文件
- 大文件上传建议使用其他方案（如AWS S3）

### 3. 数据存储
- Netlify Functions是无状态的
- 需要持久化数据请使用外部数据库（如FaunaDB、Supabase等）

### 4. 成本考虑
- 免费版有调用限制
- 复杂应用建议升级到付费计划

## 🔧 故障排查

### 构建失败
1. 检查 `package.json` 中的构建脚本
2. 确保所有依赖都已安装
3. 查看构建日志中的错误信息

### API不工作
1. 检查函数文件是否在正确位置
2. 确保 `netlify.toml` 中的重定向配置正确
3. 查看函数日志

### 前端无法访问API
1. 检查 `REACT_APP_API_URL` 环境变量
2. 确保CORS配置正确
3. 检查网络请求路径

## 📞 技术支持

如果遇到问题，可以：
1. 查看Netlify文档
2. 检查项目的GitHub Issues
3. 联系项目维护者

## 🎉 部署成功！

恭喜！你的钢材采购优化系统V3.0已成功部署到Netlify。

访问地址：`https://your-site-name.netlify.app`

---

**版本信息**
- 系统版本：V3.0.0
- 部署平台：Netlify
- 更新时间：2024年 