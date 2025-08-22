/**
 * 约束管理器 - V3.0 统一约束管理
 * 提供类型安全的约束配置访问接口，是系统中唯一的约束配置访问入口
 */

const {
  DEFAULT_CONSTRAINTS,
  VALIDATION_LIMITS,
  DATA_LIMITS,
  STANDARD_STEEL_LENGTHS,
  DEFAULT_MODULE_LENGTHS,
  REMAINDER_CONFIG,
  PERFORMANCE_CONFIG,
  ERROR_CONFIG,
  SCENARIO_CONFIGS,
  CONSTRAINT_DESCRIPTIONS,
  UNIT_CONVERSIONS,
  ENV_OVERRIDES
} = require('./ConstraintConfig');

class ConstraintManager {
  constructor() {
    this.initializeConstraints();
  }

  /**
   * 初始化约束配置，应用环境变量覆盖
   */
  initializeConstraints() {
    this.currentConstraints = { ...DEFAULT_CONSTRAINTS };
    
    // 应用环境变量覆盖
    Object.keys(ENV_OVERRIDES).forEach(key => {
      if (ENV_OVERRIDES[key] !== null) {
        this.currentConstraints[key] = ENV_OVERRIDES[key];
        console.log(`🔧 约束配置覆盖: ${key} = ${ENV_OVERRIDES[key]} (来源: 环境变量)`);
      }
    });

    console.log('📋 约束管理器初始化完成:', this.currentConstraints);
  }

  /**
   * 获取默认约束条件
   * @param {string} scenario - 可选的场景名称
   * @returns {Object} 约束条件对象
   */
  getDefaultConstraints(scenario = null) {
    if (scenario && SCENARIO_CONFIGS[scenario]) {
      const scenarioConfig = SCENARIO_CONFIGS[scenario];
      console.log(`🎯 使用场景化约束配置: ${scenario} - ${scenarioConfig.description}`);
      return {
        wasteThreshold: scenarioConfig.wasteThreshold,
        targetLossRate: scenarioConfig.targetLossRate,
        timeLimit: scenarioConfig.timeLimit,
        maxWeldingSegments: scenarioConfig.maxWeldingSegments
      };
    }
    
    return { ...this.currentConstraints };
  }

  /**
   * 获取约束验证限制
   * @param {string} constraintKey - 约束键名
   * @returns {Object} 验证限制对象
   */
  getValidationLimits(constraintKey = null) {
    if (constraintKey) {
      return VALIDATION_LIMITS[constraintKey] || null;
    }
    return VALIDATION_LIMITS;
  }

  /**
   * 验证单个约束条件
   * @param {string} key - 约束键名
   * @param {*} value - 约束值
   * @returns {Object} 验证结果
   */
  validateConstraint(key, value) {
    const limits = VALIDATION_LIMITS[key];
    if (!limits) {
      return {
        isValid: false,
        error: `未知的约束条件: ${key}`,
        code: 'UNKNOWN_CONSTRAINT'
      };
    }

    // 检查数据类型
    if (typeof value !== 'number' || isNaN(value)) {
      return {
        isValid: false,
        error: `${key} 必须是有效数字`,
        code: 'INVALID_TYPE'
      };
    }

    // 检查范围
    if (value < limits.min || value > limits.max) {
      return {
        isValid: false,
        error: `${key} 必须在 ${limits.min}-${limits.max} 范围内`,
        code: 'OUT_OF_RANGE',
        limits: limits
      };
    }

    // 检查是否在推荐范围内
    const isRecommended = value >= limits.recommended.min && value <= limits.recommended.max;
    
    return {
      isValid: true,
      isRecommended: isRecommended,
      recommendedRange: limits.recommended,
      warning: !isRecommended ? `建议将${key}设置在 ${limits.recommended.min}-${limits.recommended.max} 范围内以获得最佳效果` : null
    };
  }

  /**
   * 验证完整的约束条件对象
   * @param {Object} constraints - 约束条件对象
   * @returns {Object} 验证结果
   */
  validateConstraints(constraints) {
    const results = {
      isValid: true,
      errors: [],
      warnings: [],
      fieldResults: {}
    };

    Object.keys(DEFAULT_CONSTRAINTS).forEach(key => {
      if (constraints.hasOwnProperty(key)) {
        const validation = this.validateConstraint(key, constraints[key]);
        results.fieldResults[key] = validation;
        
        if (!validation.isValid) {
          results.isValid = false;
          results.errors.push({
            field: key,
            message: validation.error,
            code: validation.code,
            limits: validation.limits
          });
        } else if (validation.warning) {
          results.warnings.push({
            field: key,
            message: validation.warning
          });
        }
      }
    });

    return results;
  }

  /**
   * 获取数据验证限制
   * @param {string} dataType - 数据类型 ('designSteel' | 'moduleSteel')
   * @returns {Object} 数据验证限制
   */
  getDataLimits(dataType = null) {
    if (dataType) {
      return DATA_LIMITS[dataType] || null;
    }
    return DATA_LIMITS;
  }

  /**
   * 获取标准钢材长度
   * @returns {Array} 标准长度数组
   */
  getStandardSteelLengths() {
    return [...STANDARD_STEEL_LENGTHS];
  }

  /**
   * 获取默认模数钢材长度
   * @returns {Array} 默认模数钢材长度数组
   */
  getDefaultModuleLengths() {
    return [...DEFAULT_MODULE_LENGTHS];
  }

  /**
   * 获取余料管理配置
   * @returns {Object} 余料管理配置
   */
  getRemainderConfig() {
    return { ...REMAINDER_CONFIG };
  }

  /**
   * 获取性能配置
   * @returns {Object} 性能配置
   */
  getPerformanceConfig() {
    return { ...PERFORMANCE_CONFIG };
  }

  /**
   * 获取错误处理配置
   * @returns {Object} 错误处理配置
   */
  getErrorConfig() {
    return { ...ERROR_CONFIG };
  }

  /**
   * 获取约束描述
   * @param {string} constraintKey - 约束键名
   * @returns {string} 约束描述
   */
  getConstraintDescription(constraintKey) {
    return CONSTRAINT_DESCRIPTIONS[constraintKey] || '';
  }

  /**
   * 获取所有约束描述
   * @returns {Object} 所有约束描述
   */
  getAllConstraintDescriptions() {
    return { ...CONSTRAINT_DESCRIPTIONS };
  }

  /**
   * 获取可用的场景配置
   * @returns {Object} 场景配置
   */
  getAvailableScenarios() {
    return Object.keys(SCENARIO_CONFIGS).map(key => ({
      key: key,
      name: key,
      description: SCENARIO_CONFIGS[key].description,
      constraints: {
        wasteThreshold: SCENARIO_CONFIGS[key].wasteThreshold,
        targetLossRate: SCENARIO_CONFIGS[key].targetLossRate,
        timeLimit: SCENARIO_CONFIGS[key].timeLimit,
        maxWeldingSegments: SCENARIO_CONFIGS[key].maxWeldingSegments
      }
    }));
  }

  /**
   * 时间单位转换：毫秒转秒（用于前端显示）
   * @param {number} ms - 毫秒
   * @returns {number} 秒
   */
  msToSeconds(ms) {
    return UNIT_CONVERSIONS.time.msToSeconds(ms);
  }

  /**
   * 时间单位转换：秒转毫秒（用于后端处理）
   * @param {number} seconds - 秒
   * @returns {number} 毫秒
   */
  secondsToMs(seconds) {
    return UNIT_CONVERSIONS.time.secondsToMs(seconds);
  }

  /**
   * 长度单位转换：毫米转米
   * @param {number} mm - 毫米
   * @returns {number} 米
   */
  mmToM(mm) {
    return UNIT_CONVERSIONS.length.mmToM(mm);
  }

  /**
   * 长度单位转换：米转毫米
   * @param {number} m - 米
   * @returns {number} 毫米
   */
  mToMm(m) {
    return UNIT_CONVERSIONS.length.mToMm(m);
  }

  /**
   * 创建前端安全的约束对象（时间转换为秒）
   * @param {Object} constraints - 后端约束对象（时间为毫秒）
   * @returns {Object} 前端约束对象（时间为秒）
   */
  createFrontendConstraints(constraints) {
    return {
      wasteThreshold: constraints.wasteThreshold,
      targetLossRate: constraints.targetLossRate,
      timeLimit: this.msToSeconds(constraints.timeLimit),
      maxWeldingSegments: constraints.maxWeldingSegments
    };
  }

  /**
   * 创建后端约束对象（时间转换为毫秒）
   * @param {Object} frontendConstraints - 前端约束对象（时间为秒）
   * @returns {Object} 后端约束对象（时间为毫秒）
   */
  createBackendConstraints(frontendConstraints) {
    return {
      wasteThreshold: frontendConstraints.wasteThreshold,
      targetLossRate: frontendConstraints.targetLossRate,
      timeLimit: this.secondsToMs(frontendConstraints.timeLimit),
      maxWeldingSegments: frontendConstraints.maxWeldingSegments
    };
  }

  /**
   * 生成推荐的模数钢材长度
   * @param {number} requiredLength - 所需最小长度
   * @param {number} maxCount - 最大返回数量
   * @returns {Array} 推荐长度数组
   */
  generateRecommendedLengths(requiredLength, maxCount = 3) {
    const recommended = [];
    
    // 添加刚好满足要求的长度
    recommended.push(requiredLength);
    
    // 添加标准长度中大于等于要求的
    STANDARD_STEEL_LENGTHS.forEach(length => {
      if (length >= requiredLength && !recommended.includes(length)) {
        recommended.push(length);
      }
    });
    
    // 按长度排序并限制数量
    return recommended.sort((a, b) => a - b).slice(0, maxCount);
  }

  /**
   * 检查焊接约束冲突
   * @param {Array} designSteels - 设计钢材数组
   * @param {Array} moduleSteels - 模数钢材数组
   * @param {Object} constraints - 约束条件
   * @returns {Object} 冲突检查结果
   */
  checkWeldingConstraintConflict(designSteels, moduleSteels, constraints) {
    if (!designSteels.length || !moduleSteels.length) {
      return { hasConflict: false };
    }

    const maxModuleLength = Math.max(...moduleSteels.map(m => m.length));
    const conflictSteels = designSteels.filter(d => d.length > maxModuleLength);
    
    if (conflictSteels.length > 0 && constraints.maxWeldingSegments === 1) {
      const maxDesignLength = Math.max(...conflictSteels.map(s => s.length));
      const requiredSegments = Math.ceil(maxDesignLength / maxModuleLength);
      
      return {
        hasConflict: true,
        conflictCount: conflictSteels.length,
        maxConflictLength: maxDesignLength,
        maxModuleLength: maxModuleLength,
        requiredSegments: requiredSegments,
        suggestions: [
          {
            type: 'addLongerModule',
            title: '添加更长的模数钢材',
            description: `建议添加长度≥${maxDesignLength}mm的模数钢材`,
            recommendedLengths: this.generateRecommendedLengths(maxDesignLength)
          },
          {
            type: 'increaseWelding',
            title: '增加允许焊接段数',
            description: `建议将最大焊接段数调整为${requiredSegments}段以上`,
            recommendedValue: requiredSegments
          }
        ]
      };
    }

    return { hasConflict: false };
  }

  /**
   * 重置为默认配置
   */
  resetToDefaults() {
    this.initializeConstraints();
    console.log('🔄 约束配置已重置为默认值');
  }

  /**
   * 获取当前配置摘要
   * @returns {Object} 配置摘要
   */
  getConfigSummary() {
    return {
      currentConstraints: { ...this.currentConstraints },
      hasEnvironmentOverrides: Object.values(ENV_OVERRIDES).some(val => val !== null),
      environmentOverrides: Object.fromEntries(
        Object.entries(ENV_OVERRIDES).filter(([, value]) => value !== null)
      ),
      availableScenarios: Object.keys(SCENARIO_CONFIGS),
      configVersion: '3.0.0',
      lastInitialized: new Date().toISOString()
    };
  }
}

// 创建单例实例
const constraintManager = new ConstraintManager();

module.exports = constraintManager; 