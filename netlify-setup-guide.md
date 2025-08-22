# 🚀 Netlify 具体配置操作指南

## 📋 步骤一：构建设置配置

### 情况A：首次部署配置

1. **登录Netlify**
   - 访问 https://www.netlify.com
   - 点击右上角 "Log in" 按钮
   - 选择 GitHub/GitLab/Bitbucket 登录

2. **创建新站点**
   - 登录后点击 "New site from Git" 按钮
   - 选择你的Git提供商（GitHub/GitLab/Bitbucket）
   - 授权Netlify访问你的仓库
   - 在仓库列表中找到并点击你的项目

3. **配置部署设置**
   
   在出现的 "Deploy settings" 页面中：
   
   **Owner**: 会自动选择你的账户
   
   **Branch to deploy**: 选择 `main` 或 `master`
   
   **Build command**: 
   ```
   npm run build:netlify
   ```
   
   **Publish directory**: 
   ```
   client/build
   ```
   
   **点击 "Show advanced" 展开高级设置**:
   
   **Functions directory**: 
   ```
   netlify/functions
   ```

4. **开始部署**
   - 点击 "Deploy site" 按钮
   - 等待部署完成

### 情况B：已部署站点修改设置

1. **进入站点Dashboard**
   - 在Netlify Dashboard中点击你的站点名称

2. **修改构建设置**
   - 点击 "Site settings" 标签
   - 在左侧菜单点击 "Build & deploy"
   - 在 "Build settings" 区域点击 "Edit settings"
   - 修改以下设置：
     - Build command: `npm run build:netlify`
     - Publish directory: `client/build`
     - Functions directory: `netlify/functions`
   - 点击 "Save" 保存

## 📋 步骤二：环境变量配置

### 方法A：通过网站界面设置

1. **进入环境变量页面**
   - 在站点Dashboard中，点击 "Site settings"
   - 在左侧菜单中点击 "Environment variables"

2. **添加环境变量**
   
   **点击 "Add variable" 按钮**，然后添加以下变量：
   
   **变量1：**
   - Key: `NODE_ENV`
   - Value: `production`
   - 点击 "Create variable"
   
   **变量2：**
   - Key: `REACT_APP_VERSION`
   - Value: `3.0.0`
   - 点击 "Create variable"
   
   **变量3：**
   - Key: `REACT_APP_API_URL`
   - Value: `/.netlify/functions`
   - 点击 "Create variable"

### 方法B：通过命令行设置 (如果你使用Netlify CLI)

```bash
# 设置环境变量
netlify env:set NODE_ENV production
netlify env:set REACT_APP_VERSION 3.0.0
netlify env:set REACT_APP_API_URL /.netlify/functions

# 查看所有环境变量
netlify env:list
```

## 🔄 触发重新部署

设置完成后，触发重新部署：

### 方法1：通过界面
- 在站点Dashboard中点击 "Deploys" 标签
- 点击 "Trigger deploy" 按钮
- 选择 "Deploy site"

### 方法2：通过Git推送
- 对仓库进行任何提交并推送
- Netlify会自动触发新的部署

### 方法3：通过CLI
```bash
netlify deploy --prod
```

## 📊 验证配置

### 1. 检查构建日志
- 在 "Deploys" 页面点击最新的部署
- 查看构建日志，确认：
  - npm run build:netlify 命令执行成功
  - client/build 目录生成
  - Functions 正确部署

### 2. 测试环境变量
部署完成后，测试你的站点：
```bash
# 替换 YOUR_SITE_NAME 为实际站点名称
curl https://YOUR_SITE_NAME.netlify.app/.netlify/functions/health
```

### 3. 检查Functions
在站点Dashboard的 "Functions" 标签中，应该能看到：
- health
- optimize
- stats
- tasks
- upload-design-steels
- validate-constraints

## 🔧 常见问题解决

### 问题1：构建失败
- 检查 package.json 中是否有 `build:netlify` 脚本
- 确认构建命令拼写正确
- 查看构建日志中的错误信息

### 问题2：Functions不工作
- 确认 Functions directory 设置为 `netlify/functions`
- 检查函数文件是否在正确位置
- 查看函数日志

### 问题3：环境变量不生效
- 确认变量名拼写正确（注意大小写）
- React环境变量必须以 `REACT_APP_` 开头
- 设置后需要重新部署

## 📞 获取帮助

如果遇到问题：
1. 查看Netlify的构建日志
2. 查看浏览器控制台错误
3. 检查函数日志
4. 参考本项目的其他文档

---

🎉 **配置完成后，你的钢材采购优化系统V3.0就可以在Netlify上正常运行了！** 