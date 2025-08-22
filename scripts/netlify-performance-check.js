#!/usr/bin/env node

/**
 * Netlify性能监控和验证脚本
 * 用于部署后检查系统性能和功能完整性
 */

const fetch = require('node-fetch');
const fs = require('fs');
const path = require('path');

// 配置
const CONFIG = {
  // 基础URL，部署后需要修改为实际的Netlify URL
  baseUrl: process.env.NETLIFY_URL || 'https://your-site.netlify.app',
  
  // 测试超时时间
  timeout: 30000,
  
  // 重试次数
  retries: 3,
  
  // 测试间隔
  interval: 1000,
  
  // 期望的响应时间阈值（毫秒）
  performanceThresholds: {
    health: 2000,
    stats: 5000,
    optimize: 10000,
    upload: 15000
  }
};

// 颜色输出
const colors = {
  green: '\x1b[32m',
  red: '\x1b[31m',
  yellow: '\x1b[33m',
  blue: '\x1b[34m',
  reset: '\x1b[0m'
};

function log(message, color = 'reset') {
  console.log(`${colors[color]}${message}${colors.reset}`);
}

// 性能测试工具
class PerformanceMonitor {
  constructor() {
    this.results = [];
    this.startTime = Date.now();
  }

  async measureRequest(name, url, options = {}) {
    const start = Date.now();
    
    try {
      const response = await fetch(url, {
        timeout: CONFIG.timeout,
        ...options
      });
      
      const duration = Date.now() - start;
      const result = {
        name,
        url,
        status: response.status,
        duration,
        success: response.ok,
        size: response.headers.get('content-length') || 'unknown',
        timestamp: new Date().toISOString()
      };
      
      this.results.push(result);
      return result;
    } catch (error) {
      const duration = Date.now() - start;
      const result = {
        name,
        url,
        status: 'ERROR',
        duration,
        success: false,
        error: error.message,
        timestamp: new Date().toISOString()
      };
      
      this.results.push(result);
      return result;
    }
  }

  getAverageResponseTime() {
    const successfulRequests = this.results.filter(r => r.success);
    if (successfulRequests.length === 0) return 0;
    
    const totalTime = successfulRequests.reduce((sum, r) => sum + r.duration, 0);
    return Math.round(totalTime / successfulRequests.length);
  }

  getSuccessRate() {
    if (this.results.length === 0) return 0;
    const successCount = this.results.filter(r => r.success).length;
    return Math.round((successCount / this.results.length) * 100);
  }

  generateReport() {
    const totalDuration = Date.now() - this.startTime;
    
    return {
      summary: {
        totalTests: this.results.length,
        successfulTests: this.results.filter(r => r.success).length,
        failedTests: this.results.filter(r => !r.success).length,
        successRate: this.getSuccessRate(),
        averageResponseTime: this.getAverageResponseTime(),
        totalDuration: totalDuration
      },
      details: this.results,
      timestamp: new Date().toISOString()
    };
  }
}

// 测试套件
class TestSuite {
  constructor(baseUrl) {
    this.baseUrl = baseUrl;
    this.monitor = new PerformanceMonitor();
  }

  async runHealthCheck() {
    log('🔍 运行健康检查测试...', 'blue');
    
    const result = await this.monitor.measureRequest(
      'Health Check',
      `${this.baseUrl}/.netlify/functions/health`
    );
    
    if (result.success) {
      log(`✅ 健康检查通过 (${result.duration}ms)`, 'green');
      return true;
    } else {
      log(`❌ 健康检查失败: ${result.error || result.status}`, 'red');
      return false;
    }
  }

  async runStatsCheck() {
    log('📊 运行统计API测试...', 'blue');
    
    const result = await this.monitor.measureRequest(
      'Stats API',
      `${this.baseUrl}/.netlify/functions/stats`
    );
    
    if (result.success) {
      log(`✅ 统计API正常 (${result.duration}ms)`, 'green');
      return true;
    } else {
      log(`❌ 统计API失败: ${result.error || result.status}`, 'red');
      return false;
    }
  }

  async runValidationCheck() {
    log('🔧 运行约束验证测试...', 'blue');
    
    const testData = {
      designSteels: [
        { id: 'test1', length: 6000, quantity: 10, crossSection: 1000 }
      ],
      moduleSteels: [
        { id: 'mod1', name: 'Test Module', length: 12000 }
      ],
      constraints: {
        wasteThreshold: 100,
        targetLossRate: 5,
        timeLimit: 30000,
        maxWeldingSegments: 0
      }
    };
    
    const result = await this.monitor.measureRequest(
      'Validation API',
      `${this.baseUrl}/.netlify/functions/validate-constraints`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(testData)
      }
    );
    
    if (result.success) {
      log(`✅ 约束验证正常 (${result.duration}ms)`, 'green');
      return true;
    } else {
      log(`❌ 约束验证失败: ${result.error || result.status}`, 'red');
      return false;
    }
  }

  async runOptimizationCheck() {
    log('⚡ 运行优化算法测试...', 'blue');
    
    const testData = {
      designSteels: [
        { id: 'test1', length: 6000, quantity: 5, crossSection: 1000, specification: 'Test' }
      ],
      moduleSteels: [
        { id: 'mod1', name: 'Test Module', length: 12000 }
      ],
      constraints: {
        wasteThreshold: 100,
        targetLossRate: 5,
        timeLimit: 30000,
        maxWeldingSegments: 1
      }
    };
    
    const result = await this.monitor.measureRequest(
      'Optimization API',
      `${this.baseUrl}/.netlify/functions/optimize`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(testData)
      }
    );
    
    if (result.success) {
      log(`✅ 优化算法启动成功 (${result.duration}ms)`, 'green');
      return true;
    } else {
      log(`❌ 优化算法失败: ${result.error || result.status}`, 'red');
      return false;
    }
  }

  async runFrontendCheck() {
    log('🌐 运行前端页面测试...', 'blue');
    
    const result = await this.monitor.measureRequest(
      'Frontend',
      this.baseUrl
    );
    
    if (result.success) {
      log(`✅ 前端页面加载正常 (${result.duration}ms)`, 'green');
      return true;
    } else {
      log(`❌ 前端页面加载失败: ${result.error || result.status}`, 'red');
      return false;
    }
  }

  async runPerformanceCheck() {
    log('🚀 运行性能基准测试...', 'blue');
    
    const tests = [
      { name: 'Health Check', url: `${this.baseUrl}/.netlify/functions/health` },
      { name: 'Stats API', url: `${this.baseUrl}/.netlify/functions/stats` },
      { name: 'Frontend', url: this.baseUrl }
    ];
    
    for (const test of tests) {
      const result = await this.monitor.measureRequest(test.name, test.url);
      
      if (result.duration > (CONFIG.performanceThresholds[test.name.toLowerCase().split(' ')[0]] || 5000)) {
        log(`⚠️ ${test.name} 响应时间较慢: ${result.duration}ms`, 'yellow');
      } else {
        log(`✅ ${test.name} 性能良好: ${result.duration}ms`, 'green');
      }
    }
  }

  async runFullTest() {
    log('🔥 开始完整系统测试...', 'blue');
    log(`📍 测试目标: ${this.baseUrl}`, 'blue');
    
    const tests = [
      { name: '前端页面', test: () => this.runFrontendCheck() },
      { name: '健康检查', test: () => this.runHealthCheck() },
      { name: '系统统计', test: () => this.runStatsCheck() },
      { name: '约束验证', test: () => this.runValidationCheck() },
      { name: '优化算法', test: () => this.runOptimizationCheck() },
      { name: '性能基准', test: () => this.runPerformanceCheck() }
    ];
    
    let passedTests = 0;
    
    for (const testItem of tests) {
      try {
        const result = await testItem.test();
        if (result) passedTests++;
        
        // 测试间隔
        await new Promise(resolve => setTimeout(resolve, CONFIG.interval));
      } catch (error) {
        log(`❌ ${testItem.name} 测试异常: ${error.message}`, 'red');
      }
    }
    
    // 生成报告
    const report = this.monitor.generateReport();
    
    log('\n📋 测试结果汇总:', 'blue');
    log(`✅ 通过测试: ${passedTests}/${tests.length}`, 'green');
    log(`📊 成功率: ${report.summary.successRate}%`, 'blue');
    log(`⏱️ 平均响应时间: ${report.summary.averageResponseTime}ms`, 'blue');
    log(`🕐 总测试时间: ${Math.round(report.summary.totalDuration / 1000)}s`, 'blue');
    
    // 保存报告
    this.saveReport(report);
    
    return passedTests === tests.length;
  }

  saveReport(report) {
    try {
      const reportPath = path.join(process.cwd(), 'netlify-performance-report.json');
      fs.writeFileSync(reportPath, JSON.stringify(report, null, 2));
      log(`📄 测试报告已保存: ${reportPath}`, 'blue');
    } catch (error) {
      log(`⚠️ 无法保存测试报告: ${error.message}`, 'yellow');
    }
  }
}

// 主函数
async function main() {
  const args = process.argv.slice(2);
  const baseUrl = args[0] || CONFIG.baseUrl;
  
  log('🚀 Netlify性能监控和验证工具', 'green');
  log('='.repeat(50), 'blue');
  
  if (!baseUrl || baseUrl.includes('your-site')) {
    log('❌ 请提供有效的Netlify URL', 'red');
    log('用法: node netlify-performance-check.js https://your-site.netlify.app', 'yellow');
    process.exit(1);
  }
  
  const suite = new TestSuite(baseUrl);
  
  try {
    const allTestsPassed = await suite.runFullTest();
    
    if (allTestsPassed) {
      log('\n🎉 所有测试通过！系统运行正常', 'green');
      process.exit(0);
    } else {
      log('\n⚠️ 部分测试未通过，请检查系统配置', 'yellow');
      process.exit(1);
    }
  } catch (error) {
    log(`\n❌ 测试过程中发生错误: ${error.message}`, 'red');
    process.exit(1);
  }
}

// 运行主函数
if (require.main === module) {
  main();
}

module.exports = { TestSuite, PerformanceMonitor }; 