/**
 * 统一错误处理器 - V3.0
 * 提供分层错误处理和用户友好的错误信息
 */

const constraintManager = require('../config/ConstraintManager');

class ErrorHandler {
  constructor() {
    this.errorHistory = [];
    this.errorCounts = {};
  }

  /**
   * 分类错误类型
   */
  classifyError(error) {
    const errorString = error.toString().toLowerCase();
    
    if (errorString.includes('constraint') || errorString.includes('约束')) {
      return 'CONSTRAINT_ERROR';
    }
    
    if (errorString.includes('timeout') || errorString.includes('时间')) {
      return 'TIMEOUT_ERROR';
    }
    
    if (errorString.includes('memory') || errorString.includes('内存')) {
      return 'MEMORY_ERROR';
    }
    
    if (errorString.includes('calculation') || errorString.includes('计算')) {
      return 'CALCULATION_ERROR';
    }
    
    if (errorString.includes('data') || errorString.includes('数据')) {
      return 'DATA_ERROR';
    }
    
    if (errorString.includes('algorithm') || errorString.includes('算法')) {
      return 'ALGORITHM_ERROR';
    }
    
    return 'UNKNOWN_ERROR';
  }

  /**
   * 生成用户友好的错误信息
   */
  generateUserFriendlyMessage(error, errorType) {
    const userMessages = {
      'CONSTRAINT_ERROR': {
        message: '约束条件设置有问题',
        suggestions: [
          '请检查废料阈值是否合理（建议100-500mm）',
          '请检查焊接段数设置是否符合实际需求',
          '请确认设计钢材长度没有超出模数钢材能力范围'
        ],
        severity: 'warning'
      },
      'TIMEOUT_ERROR': {
        message: '优化计算超时',
        suggestions: [
          '请尝试增加计算时间限制',
          '考虑减少设计钢材数量',
          '尝试简化约束条件'
        ],
        severity: 'warning'
      },
      'MEMORY_ERROR': {
        message: '系统内存不足',
        suggestions: [
          '请减少设计钢材数量',
          '尝试分批次进行优化',
          '联系技术支持'
        ],
        severity: 'error'
      },
      'CALCULATION_ERROR': {
        message: '计算过程出现错误',
        suggestions: [
          '请检查输入数据的有效性',
          '确认钢材长度和数量都是正数',
          '如果问题持续，请联系技术支持'
        ],
        severity: 'error'
      },
      'DATA_ERROR': {
        message: '输入数据有问题',
        suggestions: [
          '请检查Excel文件格式是否正确',
          '确认所有必需字段都已填写',
          '检查数据中是否有非法字符或空值'
        ],
        severity: 'warning'
      },
      'ALGORITHM_ERROR': {
        message: '优化算法执行失败',
        suggestions: [
          '请尝试调整约束条件',
          '检查设计钢材和模数钢材的匹配性',
          '如果问题持续，请联系技术支持'
        ],
        severity: 'error'
      },
      'UNKNOWN_ERROR': {
        message: '系统出现未知错误',
        suggestions: [
          '请尝试重新提交优化任务',
          '检查网络连接是否正常',
          '如果问题持续，请联系技术支持'
        ],
        severity: 'error'
      }
    };

    return userMessages[errorType] || userMessages['UNKNOWN_ERROR'];
  }

  /**
   * 处理和记录错误
   */
  handleError(error, context = {}) {
    console.error('💥 [ErrorHandler] 捕获到原始异常:', error);
    const errorType = this.classifyError(error);
    const userFriendlyInfo = this.generateUserFriendlyMessage(error, errorType);
    
    // 记录错误统计
    this.errorCounts[errorType] = (this.errorCounts[errorType] || 0) + 1;
    
    // 记录错误历史
    const errorRecord = {
      id: Date.now() + Math.random().toString(36).substr(2, 9),
      timestamp: new Date().toISOString(),
      type: errorType,
      originalError: error.message || error.toString(),
      userMessage: userFriendlyInfo.message,
      suggestions: userFriendlyInfo.suggestions,
      severity: userFriendlyInfo.severity,
      context: context,
      stack: error.stack || 'No stack trace available'
    };
    
    this.errorHistory.push(errorRecord);
    
    // 保持历史记录在合理范围内
    if (this.errorHistory.length > 100) {
      this.errorHistory = this.errorHistory.slice(-50);
    }
    
    // 根据严重程度选择日志级别
    if (userFriendlyInfo.severity === 'error') {
      console.error(`🚨 ${errorType}:`, errorRecord);
    } else {
      console.warn(`⚠️ ${errorType}:`, errorRecord);
    }
    
    return errorRecord;
  }

  /**
   * 验证输入数据
   */
  validateInputData(designSteels, moduleSteels, constraints) {
    const errors = [];
    
    // 验证设计钢材
    if (!Array.isArray(designSteels) || designSteels.length === 0) {
      errors.push({
        field: 'designSteels',
        message: '至少需要一个设计钢材',
        code: 'MISSING_DESIGN_STEELS'
      });
    } else {
      designSteels.forEach((steel, index) => {
        if (!steel.length || steel.length <= 0) {
          errors.push({
            field: `designSteels[${index}].length`,
            message: `设计钢材${index + 1}的长度无效`,
            code: 'INVALID_LENGTH'
          });
        }
        if (!steel.quantity || steel.quantity <= 0) {
          errors.push({
            field: `designSteels[${index}].quantity`,
            message: `设计钢材${index + 1}的数量无效`,
            code: 'INVALID_QUANTITY'
          });
        }
        if (!steel.crossSection || steel.crossSection <= 0) {
          errors.push({
            field: `designSteels[${index}].crossSection`,
            message: `设计钢材${index + 1}的截面面积无效`,
            code: 'INVALID_CROSS_SECTION'
          });
        }
        
        // 🔧 修复：使用约束配置中心的数据限制，消除硬编码
        const dataLimits = constraintManager.getDataLimits('designSteel');
        if (steel.length > dataLimits.maxLength) {
          errors.push({
            field: `designSteels[${index}].length`,
            message: `设计钢材${index + 1}的长度过大（超过${dataLimits.maxLength/1000}米）`,
            code: 'LENGTH_TOO_LARGE'
          });
        }
        if (steel.quantity > dataLimits.maxQuantity) {
          errors.push({
            field: `designSteels[${index}].quantity`,
            message: `设计钢材${index + 1}的数量过大（超过${dataLimits.maxQuantity}根）`,
            code: 'QUANTITY_TOO_LARGE'
          });
        }
      });
    }
    
    // 验证模数钢材
    if (!Array.isArray(moduleSteels) || moduleSteels.length === 0) {
      errors.push({
        field: 'moduleSteels',
        message: '至少需要一种模数钢材',
        code: 'MISSING_MODULE_STEELS'
      });
    } else {
      moduleSteels.forEach((steel, index) => {
        if (!steel.length || steel.length <= 0) {
          errors.push({
            field: `moduleSteels[${index}].length`,
            message: `模数钢材${index + 1}的长度无效`,
            code: 'INVALID_MODULE_LENGTH'
          });
        }
        if (!steel.name || steel.name.trim() === '') {
          errors.push({
            field: `moduleSteels[${index}].name`,
            message: `模数钢材${index + 1}的名称不能为空`,
            code: 'MISSING_MODULE_NAME'
          });
        }
      });
    }
    
    // 🔧 修复：使用约束配置中心的验证限制，消除硬编码
    if (constraints) {
      const validationLimits = constraintManager.getValidationLimits();
      
      if (constraints.wasteThreshold !== undefined) {
        const limits = validationLimits.wasteThreshold;
        if (constraints.wasteThreshold <= 0 || constraints.wasteThreshold > limits.max) {
          errors.push({
            field: 'constraints.wasteThreshold',
            message: `废料阈值应在1-${limits.max}mm范围内`,
            code: 'INVALID_WASTE_THRESHOLD'
          });
        }
      }
      
      if (constraints.targetLossRate !== undefined) {
        const limits = validationLimits.targetLossRate;
        if (constraints.targetLossRate < limits.min || constraints.targetLossRate > limits.max) {
          errors.push({
            field: 'constraints.targetLossRate',
            message: `目标损耗率应在${limits.min}-${limits.max}%范围内`,
            code: 'INVALID_LOSS_RATE'
          });
        }
      }
      
      if (constraints.maxWeldingSegments !== undefined) {
        const limits = validationLimits.maxWeldingSegments;
        if (constraints.maxWeldingSegments < limits.min || constraints.maxWeldingSegments > limits.max) {
          errors.push({
            field: 'constraints.maxWeldingSegments',
            message: `最大焊接段数应在${limits.min}-${limits.max}范围内`,
            code: 'INVALID_WELDING_SEGMENTS'
          });
        }
      }
      
      if (constraints.timeLimit !== undefined) {
        const limits = validationLimits.timeLimit;
        if (constraints.timeLimit < limits.min || constraints.timeLimit > limits.max) {
          const minSeconds = constraintManager.msToSeconds(limits.min);
          const maxSeconds = constraintManager.msToSeconds(limits.max);
          errors.push({
            field: 'constraints.timeLimit',
            message: `时间限制应在${minSeconds}-${maxSeconds}秒范围内`,
            code: 'INVALID_TIME_LIMIT'
          });
        }
      }
    }
    
    return {
      isValid: errors.length === 0,
      errors: errors,
      errorCount: errors.length
    };
  }

  /**
   * 检查系统资源状态
   */
  checkSystemResources() {
    const warnings = [];
    
    // 🔧 修复：使用约束配置中心的系统资源阈值，消除硬编码
    const errorConfig = constraintManager.getErrorConfig();
    
    // 检查内存使用情况
    if (process.memoryUsage) {
      const memoryUsage = process.memoryUsage();
      const heapUsedMB = memoryUsage.heapUsed / 1024 / 1024;
      const heapTotalMB = memoryUsage.heapTotal / 1024 / 1024;
      
      if (heapUsedMB > errorConfig.systemResource.highMemoryThreshold) {
        warnings.push({
          type: 'HIGH_MEMORY_USAGE',
          message: `内存使用率较高: ${heapUsedMB.toFixed(2)}MB/${heapTotalMB.toFixed(2)}MB`,
          recommendation: '建议重启服务或减少并发任务数量'
        });
      }
    }
    
    return {
      status: warnings.length === 0 ? 'healthy' : 'warning',
      warnings: warnings
    };
  }

  /**
   * 获取错误统计信息
   */
  getErrorStats() {
    const totalErrors = this.errorHistory.length;
    const recentErrors = this.errorHistory.filter(
      error => Date.now() - new Date(error.timestamp).getTime() < 24 * 60 * 60 * 1000
    ).length;
    
    return {
      totalErrors: totalErrors,
      recentErrors: recentErrors,
      errorTypes: this.errorCounts,
      lastError: this.errorHistory[this.errorHistory.length - 1] || null,
      errorRate: totalErrors > 0 ? (recentErrors / totalErrors * 100).toFixed(2) + '%' : '0%'
    };
  }

  /**
   * 清空错误历史
   */
  clearErrorHistory() {
    this.errorHistory = [];
    this.errorCounts = {};
    console.log('✅ 错误历史已清空');
  }

  /**
   * 生成错误报告
   */
  generateErrorReport() {
    const stats = this.getErrorStats();
    const systemResources = this.checkSystemResources();
    
    return {
      timestamp: new Date().toISOString(),
      summary: {
        totalErrors: stats.totalErrors,
        recentErrors: stats.recentErrors,
        systemStatus: systemResources.status
      },
      details: {
        errorTypes: stats.errorTypes,
        lastError: stats.lastError,
        systemWarnings: systemResources.warnings,
        recentErrorHistory: this.errorHistory.slice(-10)
      },
      recommendations: this.generateRecommendations(stats, systemResources)
    };
  }

  /**
   * 生成建议
   */
  generateRecommendations(stats, systemResources) {
    const recommendations = [];
    
    if (stats.recentErrors > 10) {
      recommendations.push('最近错误频率较高，建议检查输入数据质量');
    }
    
    if (stats.errorTypes['CONSTRAINT_ERROR'] > 5) {
      recommendations.push('约束条件错误较多，建议优化默认约束设置');
    }
    
    if (stats.errorTypes['TIMEOUT_ERROR'] > 3) {
      recommendations.push('超时错误较多，建议增加默认计算时间限制');
    }
    
    if (systemResources.warnings.length > 0) {
      recommendations.push('系统资源存在警告，建议检查服务器性能');
    }
    
    if (recommendations.length === 0) {
      recommendations.push('系统运行状况良好，无特殊建议');
    }
    
    return recommendations;
  }
}

module.exports = ErrorHandler; 