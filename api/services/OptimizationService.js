/**
 * 优化服务 - V3.0 API层
 * 提供统一的优化接口，支持模块化调用
 */

const SteelOptimizerV3 = require('../../core/optimizer/SteelOptimizerV3');
const ErrorHandler = require('../../core/utils/ErrorHandler');
const constraintManager = require('../../core/config/ConstraintManager');
const { 
  DesignSteel, 
  ModuleSteel, 
  OptimizationConstraints
} = require('../types');

class OptimizationService {
  constructor() {
    this.activeOptimizers = new Map(); // 存储活跃的优化器实例
    this.optimizationHistory = [];
    // 🔧 统一架构：损耗率计算已整合到StatisticsCalculator中
    this.errorHandler = new ErrorHandler(); // 统一错误处理器
  }

  /**
   * 执行钢材优化
   * 主要API入口点
   */
  async optimizeSteel(requestData) {
    try {
      // 1. 🔧 使用ErrorHandler进行完善的输入验证
      const designSteels = this.createDesignSteels(requestData.designSteels || []);
      const moduleSteels = this.createModuleSteels(requestData.moduleSteels || []);
      const constraints = this.createConstraints(requestData.constraints);

      const validationResult = this.errorHandler.validateInputData(designSteels, moduleSteels, constraints);
      if (!validationResult.isValid) {
        console.warn('📋 输入数据验证失败:', validationResult.errors);
        return {
          success: false,
          error: '输入数据验证失败',
          errorType: 'VALIDATION_ERROR',
          details: validationResult.errors,
          suggestions: this.generateValidationSuggestions(validationResult.errors)
        };
      }

      // 2. 数据对象已在验证步骤中创建

      // 3. 生成优化ID并创建优化器
      const optimizationId = this.generateOptimizationId();
      const optimizer = new SteelOptimizerV3(designSteels, moduleSteels, constraints);
      
      // 4. 存储优化器实例
      this.activeOptimizers.set(optimizationId, {
        optimizer,
        startTime: Date.now(),
        status: 'running'
      });

      // 5. 执行优化
      const optimizationOutput = await optimizer.optimize();

      // 6. 更新状态和处理结果
      const optimizerInfo = this.activeOptimizers.get(optimizationId);
      if (optimizerInfo) {
        optimizerInfo.status = optimizationOutput.success ? 'completed' : 'failed';
        optimizerInfo.result = optimizationOutput; // 保存原始输出
      }
      
      // 7. 🔧 修复：直接使用算法层的完整结果，不再重复计算
      // 移除aggregateResults调用，避免数据不一致

      // 8. 记录优化历史
      this.recordOptimizationHistory({
        id: optimizationId,
        input: requestData,
        result: optimizationOutput.result, // 直接使用算法层结果
        timestamp: new Date().toISOString()
      });

      // 9. 清理完成的优化器
      setTimeout(() => {
        this.activeOptimizers.delete(optimizationId);
      }, 300000);

      // 10. 构建最终响应 - 确保与前端数据契约一致
      if (optimizationOutput.success) {
        // 成功时，直接返回包含solutions和completeStats的result对象
        return {
          success: true,
          optimizationId: optimizationId,
          // 直接将优化器产出的result作为顶层结果
          ...optimizationOutput.result, 
          // 将executionTime也合并到结果中
          executionTime: Date.now() - optimizerInfo.startTime,
        };
      } else {
        // 失败时，返回统一的错误结构
        return {
          success: false,
          optimizationId: optimizationId,
          error: optimizationOutput.error,
          executionTime: Date.now() - optimizerInfo.startTime,
        };
      }

    } catch (error) {
      // 🔧 使用ErrorHandler处理异常
      const errorRecord = this.errorHandler.handleError(error, {
        operation: 'optimizeSteel',
        requestData: {
          designSteelsCount: requestData.designSteels?.length || 0,
          moduleSteelsCount: requestData.moduleSteels?.length || 0,
          constraints: requestData.constraints
        }
      });

      return {
        success: false,
        error: errorRecord.userMessage,
        errorType: errorRecord.type,
        suggestions: errorRecord.suggestions,
        errorId: errorRecord.id,
        severity: errorRecord.severity,
        stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
      };
    }
  }

  /**
   * 🗑️ 已移除：aggregateResults方法
   * 原因：避免API层重复计算，确保数据一致性
   * 现在直接使用算法层(ResultBuilder)的完整计算结果
   */

  /**
   * 🔧 新增：生成验证错误的用户友好建议
   */
  generateValidationSuggestions(errors) {
    const suggestions = [];
    const errorTypes = new Set(errors.map(err => err.code));

    if (errorTypes.has('MISSING_DESIGN_STEELS')) {
      suggestions.push('请上传包含设计钢材数据的Excel文件');
    }
    
    if (errorTypes.has('MISSING_MODULE_STEELS')) {
      suggestions.push('请确保添加至少一种模数钢材规格');
    }
    
    if (errorTypes.has('INVALID_LENGTH') || errorTypes.has('INVALID_MODULE_LENGTH')) {
      suggestions.push('请检查钢材长度是否为正数且单位正确（单位：mm）');
    }
    
    if (errorTypes.has('INVALID_QUANTITY')) {
      suggestions.push('请确认钢材数量为正整数');
    }
    
    if (errorTypes.has('INVALID_CROSS_SECTION')) {
      suggestions.push('请检查截面面积数据是否正确（单位：mm²）');
    }
    
    if (errorTypes.has('LENGTH_TOO_LARGE')) {
      const maxLength = constraintManager.getDataLimits('designSteel').maxLength;
      suggestions.push(`钢材长度过大，请检查数据单位是否正确（建议最大${maxLength/1000}米）`);
    }
    
    if (errorTypes.has('QUANTITY_TOO_LARGE')) {
      const maxQuantity = constraintManager.getDataLimits('designSteel').maxQuantity;
      suggestions.push(`钢材数量过多，建议分批次进行优化（单次建议不超过${maxQuantity}根）`);
    }
    
    if (errorTypes.has('INVALID_WASTE_THRESHOLD')) {
      const limits = constraintManager.getValidationLimits('wasteThreshold');
      suggestions.push(`废料阈值建议设置在${limits.recommended.min}-${limits.recommended.max}mm范围内`);
    }
    
    if (errorTypes.has('INVALID_WELDING_SEGMENTS')) {
      const limits = constraintManager.getValidationLimits('maxWeldingSegments');
      suggestions.push(`焊接段数建议设置在${limits.recommended.min}-${limits.recommended.max}段范围内，1段表示不允许焊接`);
    }

    if (suggestions.length === 0) {
      suggestions.push('请检查输入数据的完整性和格式正确性');
    }

    return suggestions;
  }

  /**
   * 验证约束W是否可行
   * 在正式优化前进行预检查
   */
  async validateWeldingConstraints(requestData) {
    try {
      const designSteels = this.createDesignSteels(requestData.designSteels);
      const moduleSteels = this.createModuleSteels(requestData.moduleSteels);
      const constraints = this.createConstraints(requestData.constraints);

      // 创建临时优化器进行约束验证
      const optimizer = new SteelOptimizerV3(designSteels, moduleSteels, constraints);
      const validation = optimizer.constraintValidator.validateAllConstraints(
        designSteels, 
        moduleSteels, 
        constraints
      );

      return {
        success: true,
        validation: validation,
        suggestions: validation.suggestions || [],
        warnings: validation.warnings || []
      };

    } catch (error) {
      return {
        success: false,
        error: `约束验证失败: ${error.message}`
      };
    }
  }

  /**
   * 获取优化进度
   */
  getOptimizationProgress(optimizationId) {
    const optimizerInfo = this.activeOptimizers.get(optimizationId);
    
    if (!optimizerInfo) {
      return {
        success: false,
        error: '优化ID不存在或已过期'
      };
    }

    const runningTime = Date.now() - optimizerInfo.startTime;
    
    return {
      success: true,
      status: optimizerInfo.status,
      runningTime: runningTime,
      result: optimizerInfo.result || null
    };
  }

  /**
   * 取消优化
   */
  cancelOptimization(optimizationId) {
    const optimizerInfo = this.activeOptimizers.get(optimizationId);
    
    if (!optimizerInfo) {
      return {
        success: false,
        error: '优化ID不存在或已完成'
      };
    }

    // 标记为取消
    optimizerInfo.status = 'cancelled';
    this.activeOptimizers.delete(optimizationId);

    return {
      success: true,
      message: '优化已取消'
    };
  }

  /**
   * 验证输入数据
   */
  validateInput(requestData) {
    const errors = [];
    const suggestions = [];

    // 检查设计钢材
    if (!requestData.designSteels || !Array.isArray(requestData.designSteels)) {
      errors.push('设计钢材数据缺失或格式错误');
    } else if (requestData.designSteels.length === 0) {
      errors.push('至少需要一个设计钢材');
    } else {
      requestData.designSteels.forEach((steel, index) => {
        if (!steel.length || steel.length <= 0) {
          errors.push(`设计钢材${index + 1}: 长度无效`);
        }
        if (!steel.quantity || steel.quantity <= 0) {
          errors.push(`设计钢材${index + 1}: 数量无效`);
        }
        if (!steel.crossSection || steel.crossSection <= 0) {
          errors.push(`设计钢材${index + 1}: 截面面积无效`);
        }
      });
    }

    // 检查模数钢材
    if (!requestData.moduleSteels || !Array.isArray(requestData.moduleSteels)) {
      errors.push('模数钢材数据缺失或格式错误');
    } else if (requestData.moduleSteels.length === 0) {
      errors.push('至少需要一种模数钢材');
    } else {
      requestData.moduleSteels.forEach((steel, index) => {
        if (!steel.length || steel.length <= 0) {
          errors.push(`模数钢材${index + 1}: 长度无效`);
        }
        if (!steel.name || steel.name.trim() === '') {
          errors.push(`模数钢材${index + 1}: 名称不能为空`);
        }
      });
    }

    // 检查约束条件
    if (requestData.constraints) {
      const c = requestData.constraints;
      if (c.wasteThreshold !== undefined && c.wasteThreshold <= 0) {
        errors.push('废料阈值必须大于0');
      }
      if (c.expectedLossRate !== undefined && (c.expectedLossRate < 0 || c.expectedLossRate > 50)) {
        errors.push('期望损耗率必须在0-50%之间');
      }
      if (c.weldingSegments !== undefined && c.weldingSegments < 1) {
        errors.push('允许焊接段数必须≥1');
      }
    }

    // 生成建议
    if (requestData.designSteels && requestData.moduleSteels) {
      const maxDesignLength = Math.max(...requestData.designSteels.map(s => s.length));
      const maxModuleLength = Math.max(...requestData.moduleSteels.map(s => s.length));
      
      if (maxDesignLength > maxModuleLength) {
        const W = requestData.constraints?.weldingSegments || 1;
        if (W === 1) {
          suggestions.push({
            type: 'weldingConstraintConflict',
            message: '检测到可能的焊接约束冲突',
            recommendation: '建议在正式优化前运行约束验证'
          });
        }
      }
    }

    return {
      isValid: errors.length === 0,
      errors: errors,
      suggestions: suggestions
    };
  }

  /**
   * 创建设计钢材对象
   */
  createDesignSteels(designSteelsData) {
    return designSteelsData.map((steel, index) => new DesignSteel({
      id: steel.id || `design_${index + 1}`,
      length: steel.length,
      quantity: steel.quantity,
      crossSection: steel.crossSection,
      displayId: steel.displayId || '',
      componentNumber: steel.componentNumber || '',
      specification: steel.specification || '',
      partNumber: steel.partNumber || ''
    }));
  }

  /**
   * 创建模数钢材对象
   */
  createModuleSteels(moduleSteelsData) {
    return moduleSteelsData.map((steel, index) => new ModuleSteel({
      id: steel.id || `module_${index + 1}`,
      name: steel.name,
      length: steel.length
    }));
  }

  /**
   * 创建约束对象
   */
  createConstraints(constraintsData = {}) {
    // 🔧 修复：使用约束配置中心的默认值，消除硬编码
    const defaults = constraintManager.getDefaultConstraints();
    
    return new OptimizationConstraints({
      wasteThreshold: constraintsData.wasteThreshold || defaults.wasteThreshold,
      targetLossRate: constraintsData.targetLossRate || defaults.targetLossRate,
      timeLimit: constraintsData.timeLimit || defaults.timeLimit,
      maxWeldingSegments: constraintsData.maxWeldingSegments || defaults.maxWeldingSegments
    });
  }

  /**
   * 生成优化ID
   */
  generateOptimizationId() {
    const timestamp = Date.now();
    const random = Math.random().toString(36).substr(2, 9);
    return `opt_${timestamp}_${random}`;
  }

  /**
   * 记录优化历史
   */
  recordOptimizationHistory(record) {
    this.optimizationHistory.push(record);
    
    // 保持历史记录数量在合理范围内
    if (this.optimizationHistory.length > 100) {
      this.optimizationHistory = this.optimizationHistory.slice(-50);
    }
  }

  /**
   * 获取优化历史
   */
  getOptimizationHistory(limit = 20) {
    return {
      success: true,
      history: this.optimizationHistory
        .slice(-limit)
        .reverse()
        .map(record => ({
          id: record.id,
          timestamp: record.timestamp,
          success: record.result?.success || false,
          executionTime: record.result?.result?.executionTime || 0,
          totalLossRate: record.result?.result?.totalLossRate || 0,
          totalModuleUsed: record.result?.result?.totalModuleUsed || 0
        }))
    };
  }

  /**
   * 获取系统统计
   */
  getSystemStats() {
    const activeCount = this.activeOptimizers.size;
    const totalOptimizations = this.optimizationHistory.length;
    const successfulOptimizations = this.optimizationHistory.filter(
      record => record.result?.success
    ).length;

    return {
      success: true,
      stats: {
        activeOptimizations: activeCount,
        totalOptimizations: totalOptimizations,
        successfulOptimizations: successfulOptimizations,
        successRate: totalOptimizations > 0 ? 
          ((successfulOptimizations / totalOptimizations) * 100).toFixed(2) : 0,
        averageExecutionTime: this.calculateAverageExecutionTime()
      }
    };
  }

  /**
   * 计算平均执行时间
   */
  calculateAverageExecutionTime() {
    const successfulRecords = this.optimizationHistory.filter(
      record => record.result?.success
    );
    
    if (successfulRecords.length === 0) return 0;
    
    const totalTime = successfulRecords.reduce(
      (sum, record) => sum + (record.result?.result?.executionTime || 0), 
      0
    );
    
    return Math.round(totalTime / successfulRecords.length);
  }

  /**
   * 清理过期的优化器
   */
  cleanupExpiredOptimizers() {
    const now = Date.now();
    const expireTime = 300000; // 5分钟
    
    for (const [id, info] of this.activeOptimizers.entries()) {
      if (now - info.startTime > expireTime) {
        this.activeOptimizers.delete(id);
      }
    }
  }

  /**
   * 获取活跃优化器列表
   */
  getActiveOptimizers() {
    const result = [];
    
    for (const [id, info] of this.activeOptimizers.entries()) {
      result.push({
        id: id,
        status: info.status,
        startTime: info.startTime,
        runningTime: Date.now() - info.startTime
      });
    }
    
    return {
      success: true,
      activeOptimizers: result
    };
  }

  /**
   * 运行优化过程，支持进度回调。
   * 这是为了适应异步后台任务而设计的包装器方法。
   * @param {object} optimizationData - 包含设计钢材、模数钢材和约束的数据。
   * @param {function} progressCallback - 用于报告进度的回调函数。
   * @returns {Promise<object>} - 返回完整的优化结果。
   */
  async run(optimizationData, progressCallback) {
    try {
      console.log('✅ service.run() 已被调用，优化正式开始');
      await progressCallback(15, '参数验证和初始化...');
      
      // 调用现有的核心优化方法
      const results = await this.optimizeSteel(optimizationData);
      
      // 注意：这里的进度更新只是为了完成流程，
      // 因为optimizeSteel是同步的，它完成后，整个任务就完成了。
      // 如果optimizeSteel内部有更多步骤，可以在其中多次调用progressCallback。
      
      console.log('✅ service.run() 即将返回结果');
      return results;

    } catch (error) {
      console.error('❌ 在 service.run() 内部捕获到错误:', error);
      // 将错误向上抛出，以便后台工作者可以捕获并记录它
      throw error;
    }
  }
}

module.exports = OptimizationService; 