/**
 * 约束配置中心 - V3.0 统一配置
 * 消除系统中所有硬编码约束值，提供集中化配置管理
 */

/**
 * 基础约束默认值
 * 注意：时间单位统一使用毫秒，前端显示时转换为秒
 */
const DEFAULT_CONSTRAINTS = {
  wasteThreshold: 100,        // 废料阈值 (mm) - 余料长度小于此值时视为废料
  targetLossRate: 5,          // 目标损耗率 (%) - 优化目标参考值（软约束）
  timeLimit: 30000,           // 计算时间限制 (ms) - 算法最大运行时间
  maxWeldingSegments: 1       // 最大焊接段数 (段) - 单根设计钢材允许的最大焊接段数
};

/**
 * 约束验证限制
 * 定义各约束条件的合理范围
 */
const VALIDATION_LIMITS = {
  wasteThreshold: {
    min: 1,               // 最小废料阈值 (mm)
    max: 5000,            // 最大废料阈值 (mm)
    recommended: {
      min: 50,            // 推荐最小值
      max: 500            // 推荐最大值
    }
  },
  targetLossRate: {
    min: 0,               // 最小损耗率 (%)
    max: 100,             // 最大损耗率 (%)
    recommended: {
      min: 2,             // 推荐最小值
      max: 15             // 推荐最大值
    }
  },
  timeLimit: {
    min: 1000,            // 最小时间限制 (ms) = 1秒
    max: 300000,          // 最大时间限制 (ms) = 5分钟
    recommended: {
      min: 5000,          // 推荐最小值 = 5秒
      max: 60000          // 推荐最大值 = 1分钟
    }
  },
  maxWeldingSegments: {
    min: 1,               // 最小焊接段数
    max: 10,              // 最大焊接段数
    recommended: {
      min: 1,             // 推荐最小值
      max: 3              // 推荐最大值
    }
  }
};

/**
 * 数据验证限制
 * 用于输入数据的合理性检查
 */
const DATA_LIMITS = {
  designSteel: {
    maxLength: 50000,           // 设计钢材最大长度 (mm)
    maxQuantity: 10000,         // 设计钢材最大数量
    maxCrossSection: 100000     // 最大截面面积 (mm²)
  },
  moduleSteel: {
    minLength: 1000,            // 模数钢材最小长度 (mm)
    maxLength: 50000            // 模数钢材最大长度 (mm)
  }
};

/**
 * 标准钢材长度配置
 * 用于约束验证时的推荐长度生成
 */
const STANDARD_STEEL_LENGTHS = [6000, 9000, 12000, 15000, 18000];

/**
 * 默认模数钢材长度配置
 * 用于SpecificationModuleSteelPool的默认可用长度
 */
const DEFAULT_MODULE_LENGTHS = [12000, 10000, 8000, 6000];

/**
 * 余料管理配置
 */
const REMAINDER_CONFIG = {
  idGeneration: {
    letterLimit: 50,            // 单字母余料ID数量限制
    maxLetters: 26              // 最大字母数量（a-z）
  },
  defaultWasteThreshold: 100,   // 余料管理器默认废料阈值（与DEFAULT_CONSTRAINTS.wasteThreshold保持一致）
  binarySearchEnabled: true,    // 是否启用二分查找优化
  dynamicAlgorithmSelection: true, // 是否启用动态算法选择
  algorithmThresholds: {
    smallScalePoolSize: 20,     // 小规模问题的池大小阈值
    maxSegmentsForDP: 2         // 动态规划算法的最大段数阈值
  }
};

/**
 * 性能配置
 */
const PERFORMANCE_CONFIG = {
  parallelOptimization: {
    maxActiveOptimizers: 10,    // 最大并发优化器数量
    optimizerTimeout: 300000,   // 优化器超时时间 (ms)
    cleanupInterval: 60000,     // 清理间隔 (ms)
    historyLimit: 100           // 历史记录数量限制
  },
  iterationLimits: {
    maxIterationsPerDemand: 100,  // 每个需求的最大迭代次数倍数
    maxTotalIterations: 1000000   // 总的最大迭代次数（安全上限）
  }
};

/**
 * 错误处理配置
 */
const ERROR_CONFIG = {
  validation: {
    maxErrorHistorySize: 100,   // 最大错误历史记录数
    recentErrorWindow: 86400000 // 最近错误时间窗口 (ms) = 24小时
  },
  systemResource: {
    highMemoryThreshold: 1000,  // 高内存使用阈值 (MB)
    maxFileSize: 10485760       // 最大文件大小 (bytes) = 10MB
  }
};

/**
 * 场景化约束配置
 * 支持不同业务场景的预设约束组合
 */
const SCENARIO_CONFIGS = {
  // 高精度场景：桥梁、重要结构
  'precision': {
    wasteThreshold: 50,
    targetLossRate: 2,
    timeLimit: 60000,           // 1分钟
    maxWeldingSegments: 1,
    description: '高精度场景 - 适用于桥梁、重要结构等对精度要求极高的项目'
  },
  
  // 标准场景：一般建筑
  'standard': {
    wasteThreshold: 100,
    targetLossRate: 5,
    timeLimit: 30000,           // 30秒
    maxWeldingSegments: 2,
    description: '标准场景 - 适用于一般建筑项目，平衡精度和效率'
  },
  
  // 经济场景：成本优先
  'economic': {
    wasteThreshold: 200,
    targetLossRate: 8,
    timeLimit: 15000,           // 15秒
    maxWeldingSegments: 3,
    description: '经济场景 - 适用于成本敏感项目，优先考虑效率'
  },
  
  // 快速场景：时间优先
  'fast': {
    wasteThreshold: 150,
    targetLossRate: 6,
    timeLimit: 10000,           // 10秒
    maxWeldingSegments: 2,
    description: '快速场景 - 适用于时间紧急的项目，快速给出优化结果'
  }
};

/**
 * 约束描述配置
 * 用于用户界面的帮助文本
 */
const CONSTRAINT_DESCRIPTIONS = {
  wasteThreshold: '当余料长度小于此值时，将被视为废料无法再次利用',
  targetLossRate: '算法优化时的目标损耗率，作为参考值（不是强制要求）',
  timeLimit: '算法计算的最大允许时间，超时后返回当前最优解',
  maxWeldingSegments: '单根设计钢材允许的最大焊接段数，1段表示不允许焊接（V3核心功能）'
};

/**
 * 单位转换配置
 */
const UNIT_CONVERSIONS = {
  time: {
    // 前端显示单位：秒，后端处理单位：毫秒
    msToSeconds: (ms) => Math.round(ms / 1000),
    secondsToMs: (seconds) => Math.round(seconds * 1000)
  },
  length: {
    // 统一使用毫米作为基础单位
    mmToM: (mm) => mm / 1000,
    mToMm: (m) => m * 1000
  }
};

/**
 * 环境变量覆盖支持
 * 支持通过环境变量覆盖默认配置
 */
const ENV_OVERRIDES = {
  wasteThreshold: process.env.HGJ_WASTE_THRESHOLD ? parseInt(process.env.HGJ_WASTE_THRESHOLD) : null,
  targetLossRate: process.env.HGJ_TARGET_LOSS_RATE ? parseFloat(process.env.HGJ_TARGET_LOSS_RATE) : null,
  timeLimit: process.env.HGJ_TIME_LIMIT ? parseInt(process.env.HGJ_TIME_LIMIT) : null,
  maxWeldingSegments: process.env.HGJ_MAX_WELDING_SEGMENTS ? parseInt(process.env.HGJ_MAX_WELDING_SEGMENTS) : null
};

module.exports = {
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
}; 