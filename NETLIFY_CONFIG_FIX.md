# 🔧 Netlify构建配置修复指南

## ❌ 发现的问题
你的构建命令被错误配置为：
```
Build command: npm run build:netlify Publish directory: client/build
```

## ✅ 正确的配置

### 第1步：分别填写三个字段

**Build command（构建命令）：**
```
npm run build:netlify
```

**Publish directory（发布目录）：**
```
client/build
```

**Functions directory（函数目录）：**
```
netlify/functions
```

### 第2步：具体操作

1. **进入站点设置**
   - Site settings > Build & deploy > Build settings
   - 点击 "Edit settings"

2. **清除并重新填写**
   - **Build command** 框：只填 `npm run build:netlify`
   - **Publish directory** 框：只填 `client/build`
   - **Functions directory** 框：只填 `netlify/functions`

3. **保存设置**
   - 点击 "Save" 保存
   - 点击 "Trigger deploy" > "Deploy site" 重新部署

## 🎯 关键点

- ✅ **每个字段分开填写**
- ✅ **Build command只能是命令，不能包含目录信息**
- ✅ **Publish directory只能是目录路径**
- ✅ **修改后必须重新部署**

## 📊 验证配置正确

修复后，构建日志应该显示：
```
Build command from netlify.toml: npm run build:netlify
Publish directory: client/build
Functions directory: netlify/functions
```

而不是：
```
Build command: npm run build:netlify Publish directory: client/build
``` 