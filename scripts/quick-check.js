#!/usr/bin/env node

/**
 * 快速部署检查脚本 - 解决常见问题
 */

const fs = require('fs');
const path = require('path');

// 颜色输出
const colors = {
  green: '\x1b[32m',
  red: '\x1b[31m',
  yellow: '\x1b[33m',
  blue: '\x1b[34m',
  cyan: '\x1b[36m',
  reset: '\x1b[0m'
};

function log(message, color = 'reset') {
  console.log(`${colors[color]}${message}${colors.reset}`);
}

function checkFile(filePath, description) {
  const exists = fs.existsSync(path.resolve(filePath));
  if (exists) {
    log(`✅ ${description}`, 'green');
  } else {
    log(`❌ ${description} - 文件不存在: ${filePath}`, 'red');
  }
  return exists;
}

function main() {
  log('🚀 钢材优化系统V3.0 - 快速部署检查', 'cyan');
  log('='.repeat(50), 'cyan');
  
  let allGood = true;
  
  // 1. 检查关键配置文件
  log('\n📋 检查配置文件:', 'blue');
  allGood &= checkFile('netlify.toml', 'Netlify配置文件');
  allGood &= checkFile('package.json', '根目录package.json');
  allGood &= checkFile('client/package.json', '客户端package.json');
  allGood &= checkFile('client/craco.config.js', 'Craco配置文件');
  
  // 2. 检查Netlify Functions
  log('\n⚡ 检查API Functions:', 'blue');
  const functions = [
    'health.js',
    'optimize.js', 
    'upload-design-steels.js',
    'validate-constraints.js',
    'task.js',
    'tasks.js',
    'stats.js'
  ];
  
  functions.forEach(func => {
    allGood &= checkFile(`netlify/functions/${func}`, func);
  });
  
  // 3. 检查核心代码文件
  log('\n🔧 检查核心代码:', 'blue');
  allGood &= checkFile('netlify/functions/utils/TaskManager.js', 'TaskManager');
  allGood &= checkFile('api/services/OptimizationService.js', 'OptimizationService');
  allGood &= checkFile('core/optimizer/SteelOptimizerV3.js', 'SteelOptimizerV3');
  
  // 4. 检查数据库脚本
  log('\n🗄️ 检查数据库脚本:', 'blue');
  allGood &= checkFile('database/init-netlify.sql', '数据库初始化脚本');
  
  // 5. 检查文档
  log('\n📚 检查部署文档:', 'blue');
  allGood &= checkFile('NETLIFY_ENV_VARS.md', '环境变量指南');
  allGood &= checkFile('NETLIFY_DEPLOYMENT_CHECKLIST.md', '部署检查清单');
  
  // 6. 验证关键配置内容
  log('\n🔍 验证配置内容:', 'blue');
  
  try {
    // 检查netlify.toml关键配置
    const netlifyConfig = fs.readFileSync('netlify.toml', 'utf8');
    const hasCommand = netlifyConfig.includes('command = "npm run build:netlify"');
    const hasPublish = netlifyConfig.includes('publish = "client/build"');
    const hasFunctions = netlifyConfig.includes('directory = "netlify/functions"');
    
    if (hasCommand && hasPublish && hasFunctions) {
      log('✅ netlify.toml 配置正确', 'green');
    } else {
      log('❌ netlify.toml 配置有问题', 'red');
      allGood = false;
    }
    
    // 检查package.json构建脚本
    const packageJson = JSON.parse(fs.readFileSync('package.json', 'utf8'));
    if (packageJson.scripts && packageJson.scripts['build:netlify']) {
      log('✅ 构建脚本配置正确', 'green');
    } else {
      log('❌ 缺少build:netlify脚本', 'red');
      allGood = false;
    }
    
  } catch (error) {
    log(`❌ 配置验证失败: ${error.message}`, 'red');
    allGood = false;
  }
  
  // 结果汇总
  log('\n📊 检查结果:', 'cyan');
  if (allGood) {
    log('🎉 所有检查通过！系统准备就绪', 'green');
    log('\n📝 接下来的步骤:', 'blue');
    log('1. 在Neon创建PostgreSQL数据库', 'blue');
    log('2. 运行: psql DATABASE_URL -f database/init-netlify.sql', 'blue');
    log('3. 在Netlify配置环境变量（参考NETLIFY_ENV_VARS.md）', 'blue');
    log('4. 部署到Netlify', 'blue');
    log('5. 运行性能检查: npm run performance-check YOUR_SITE_URL', 'blue');
  } else {
    log('⚠️ 发现问题，请检查上述失败项目', 'yellow');
    log('\n🔧 常见解决方案:', 'blue');
    log('- 确保所有文件都已提交到Git', 'blue');
    log('- 检查文件路径和名称是否正确', 'blue');
    log('- 运行npm install确保依赖已安装', 'blue');
  }
  
  log('\n📞 如需帮助，请参考:', 'blue');
  log('- NETLIFY_DEPLOYMENT_CHECKLIST.md', 'blue');
  log('- NETLIFY_ENV_VARS.md', 'blue');
  log('- NETLIFY_DEPLOYMENT_GUIDE.md', 'blue');
}

if (require.main === module) {
  main();
}

module.exports = { main }; 