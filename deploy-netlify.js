#!/usr/bin/env node

/**
 * 钢材采购优化系统 V3.0 - Netlify部署辅助脚本
 * 
 * 使用方法：
 * node deploy-netlify.js
 */

const fs = require('fs');
const path = require('path');

console.log('🚀 钢材采购优化系统 V3.0 - Netlify部署准备');
console.log('================================================');

// 检查必要文件
const requiredFiles = [
  'netlify.toml',
  'package.json',
  'client/package.json',
  'netlify/functions/health.js',
  'netlify/functions/optimize.js',
  'netlify/functions/upload-design-steels.js',
  'netlify/functions/validate-constraints.js',
  'netlify/functions/stats.js'
];

console.log('📋 检查必要文件...');
const missingFiles = [];

requiredFiles.forEach(file => {
  if (fs.existsSync(file)) {
    console.log(`  ✅ ${file}`);
  } else {
    console.log(`  ❌ ${file} (缺失)`);
    missingFiles.push(file);
  }
});

if (missingFiles.length > 0) {
  console.log('\n❌ 缺失关键文件！请先创建以下文件：');
  missingFiles.forEach(file => console.log(`  - ${file}`));
  process.exit(1);
}

// 检查package.json中的构建脚本
console.log('\n🔧 检查构建脚本...');
const packageJson = JSON.parse(fs.readFileSync('package.json', 'utf8'));

if (packageJson.scripts && packageJson.scripts['build:netlify']) {
  console.log('  ✅ build:netlify 脚本已配置');
} else {
  console.log('  ❌ 缺少 build:netlify 脚本');
  console.log('  请在package.json中添加: "build:netlify": "npm run install-all && cd client && npm run build"');
  process.exit(1);
}

// 检查Netlify配置
console.log('\n⚙️  检查Netlify配置...');
if (fs.existsSync('netlify.toml')) {
  const netlifyConfig = fs.readFileSync('netlify.toml', 'utf8');
  
  if (netlifyConfig.includes('build:netlify')) {
    console.log('  ✅ netlify.toml 构建命令已配置');
  } else {
    console.log('  ⚠️  netlify.toml 可能需要更新构建命令');
  }
  
  if (netlifyConfig.includes('client/build')) {
    console.log('  ✅ netlify.toml 发布目录已配置');
  } else {
    console.log('  ⚠️  netlify.toml 可能需要更新发布目录');
  }
} else {
  console.log('  ❌ netlify.toml 文件不存在');
  process.exit(1);
}

// 生成部署检查清单
console.log('\n📝 生成部署检查清单...');
const checklist = `
# 钢材采购优化系统 V3.0 - Netlify部署检查清单

## 🔍 部署前检查

### 1. 文件准备 ✅
- [x] netlify.toml 配置文件
- [x] package.json 构建脚本
- [x] Netlify Functions 文件
- [x] 客户端代码

### 2. 环境配置
- [ ] 在Netlify Dashboard设置环境变量：
  - NODE_ENV=production
  - REACT_APP_VERSION=3.0.0
  - REACT_APP_API_URL=/.netlify/functions

### 3. 部署步骤
- [ ] 连接Git仓库到Netlify
- [ ] 配置构建设置：
  - Build command: npm run build:netlify
  - Publish directory: client/build
  - Functions directory: netlify/functions
- [ ] 触发部署

### 4. 部署后测试
- [ ] 访问网站首页
- [ ] 测试健康检查API: /.netlify/functions/health
- [ ] 测试文件上传功能
- [ ] 测试优化算法功能

## 📞 支持
如有问题，请查看 NETLIFY_DEPLOYMENT_GUIDE.md 详细说明。
`;

fs.writeFileSync('NETLIFY_CHECKLIST.md', checklist);
console.log('  ✅ 已生成 NETLIFY_CHECKLIST.md');

// 显示下一步操作
console.log('\n🎉 准备完成！');
console.log('================================================');
console.log('下一步操作：');
console.log('');
console.log('1. 将代码推送到Git仓库');
console.log('2. 登录 https://www.netlify.com');
console.log('3. 点击 "New site from Git"');
console.log('4. 选择你的仓库');
console.log('5. 配置构建设置：');
console.log('   - Build command: npm run build:netlify');
console.log('   - Publish directory: client/build');
console.log('   - Functions directory: netlify/functions');
console.log('6. 设置环境变量：');
console.log('   - NODE_ENV=production');
console.log('   - REACT_APP_VERSION=3.0.0');
console.log('   - REACT_APP_API_URL=/.netlify/functions');
console.log('7. 点击 "Deploy site"');
console.log('');
console.log('📖 详细说明请查看: NETLIFY_DEPLOYMENT_GUIDE.md');
console.log('📋 部署清单请查看: NETLIFY_CHECKLIST.md');
console.log('');
console.log('🚀 祝部署顺利！'); 