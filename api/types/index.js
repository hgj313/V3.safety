/**
 * 钢材采购优化系统 V3.0 - 核心类型定义
 * 包含新的余料系统、约束W和模块化架构
 */

// ==================== 基础数据类型 ====================

/**
 * 设计钢材类型
 */
class DesignSteel {
  constructor({
    id,
    length,
    quantity,
    crossSection,
    displayId = '',
    componentNumber = '',
    specification = '',
    partNumber = ''
  }) {
    this.id = id;
    this.length = length;
    this.quantity = quantity;
    this.crossSection = crossSection;
    this.displayId = displayId;
    this.componentNumber = componentNumber;
    this.specification = specification;
    this.partNumber = partNumber;
  }
}

/**
 * 模数钢材类型
 */
class ModuleSteel {
  constructor({ id, name, length }) {
    this.id = id;
    this.name = name;
    this.length = length;
  }
}

/**
 * 余料类型 - V3.0新设计
 * 支持伪余料/真余料/废料的状态管理
 */
class RemainderV3 {
  constructor({
    id,
    length,
    type = 'pending', // 'pseudo' | 'real' | 'waste' | 'pending'
    isConsumed = false,
    sourceChain = [],
    crossSection,
    createdAt = new Date().toISOString(),
    consumedAt = null,
    originalLength = null,
    parentId = null
  }) {
    this.id = id;
    this.length = length;
    this.type = type; // 余料类型：伪余料/真余料/废料/待定
    this.isConsumed = isConsumed; // 是否被后续使用
    this.sourceChain = sourceChain; // 来源链
    this.crossSection = crossSection;
    this.createdAt = createdAt;
    this.consumedAt = consumedAt;
    this.originalLength = originalLength; // 原始长度
    this.parentId = parentId; // 父余料ID
  }

  /**
   * 标记为伪余料（已被使用）
   */
  markAsPseudo() {
    this.type = 'pseudo';
    this.isConsumed = true;
    this.consumedAt = new Date().toISOString();
  }

  /**
   * 标记为真余料（未使用，计入损耗）
   */
  markAsReal() {
    this.type = 'real';
    this.isConsumed = false;
  }

  /**
   * 标记为废料
   */
  markAsWaste() {
    this.type = 'waste';
  }
}

// ==================== 约束系统 ====================

/**
 * 优化约束条件
 * 🔧 V3.1 修复：与约束配置中心保持一致，统一参数名称
 */
class OptimizationConstraints {
  constructor({
    wasteThreshold = 100,
    targetLossRate = 5,
    timeLimit = 30000,
    maxWeldingSegments = 1,
    // 保留V3动态焊接约束参数以保持兼容性
    weldingSegments = null,      // 向后兼容
    maxIterations = 1000,
    allowDynamicWelding = true,
    minWeldingSegments = 1,
    weldingCostPerSegment = 0.1,
    weldingTimePerSegment = 5
  } = {}) {
    // 🔧 修复：统一约束参数命名，与约束配置中心保持一致
    this.wasteThreshold = wasteThreshold;
    this.targetLossRate = targetLossRate;
    this.timeLimit = timeLimit;
    this.maxWeldingSegments = maxWeldingSegments;
    
    // 向后兼容：如果传入了weldingSegments，映射到maxWeldingSegments
    if (weldingSegments !== null) {
      this.maxWeldingSegments = weldingSegments;
    }
    this.weldingSegments = this.maxWeldingSegments; // 保持兼容性
    
    // V3动态焊接约束（保留但不再是主要参数）
    this.allowDynamicWelding = allowDynamicWelding;
    this.minWeldingSegments = minWeldingSegments;
    this.weldingCostPerSegment = weldingCostPerSegment;
    this.weldingTimePerSegment = weldingTimePerSegment;
    this.maxIterations = maxIterations;
    
    // 🔧 修复：移除严格的验证逻辑，使其更灵活
    // 验证逻辑现在由约束配置中心统一处理
    this.validateBasicConstraints();
  }

  /**
   * V3.1 修复：基础约束验证（更宽松的验证逻辑）
   * 严格验证由约束配置中心统一处理
   */
  validateBasicConstraints() {
    // 🔧 修复：只进行最基本的验证，避免过于严格的限制
    if (this.wasteThreshold <= 0) {
      throw new Error(`废料阈值${this.wasteThreshold}必须大于0`);
    }
    
    if (this.maxWeldingSegments < 1) {
      throw new Error(`最大焊接段数${this.maxWeldingSegments}不能小于1`);
    }
    
    if (this.timeLimit <= 0) {
      throw new Error(`时间限制${this.timeLimit}必须大于0`);
    }
    
    // 不再抛出严格的范围错误，而是给出警告
    if (this.maxWeldingSegments > 10) {
      console.warn(`⚠️ 焊接段数${this.maxWeldingSegments}较大，可能影响优化效果`);
    }
  }

  /**
   * V3新增：动态设置焊接段数
   */
  setWeldingSegments(segments) {
    if (segments < this.minWeldingSegments || segments > this.maxWeldingSegments) {
      throw new Error(`焊接段数${segments}超出允许范围[${this.minWeldingSegments}, ${this.maxWeldingSegments}]`);
    }
    
    this.weldingSegments = segments;
    console.log(`🔧 动态设置焊接约束: 允许${segments}段焊接`);
  }

  /**
   * V3新增：计算焊接成本
   */
  calculateWeldingCost(segments) {
    if (segments <= 1) return 0;
    return (segments - 1) * this.weldingCostPerSegment;
  }

  /**
   * V3新增：计算焊接时间成本
   */
  calculateWeldingTime(segments) {
    if (segments <= 1) return 0;
    return (segments - 1) * this.weldingTimePerSegment;
  }

  /**
   * V3新增：获取焊接约束信息
   */
  getWeldingConstraintInfo() {
    return {
      current: this.weldingSegments,
      min: this.minWeldingSegments,
      max: this.maxWeldingSegments,
      allowDynamic: this.allowDynamicWelding,
      costPerSegment: this.weldingCostPerSegment,
      timePerSegment: this.weldingTimePerSegment
    };
  }

  /**
   * 验证焊接段数约束的可行性
   * 检查设计钢材长度是否超出模数钢材能力范围
   */
  validateWeldingConstraint(designSteels, moduleSteels) {
    const maxModuleLength = Math.max(...moduleSteels.map(m => m.length));
    const conflictSteels = designSteels.filter(d => d.length > maxModuleLength);
    
    if (conflictSteels.length > 0 && this.maxWeldingSegments === 1) {
      return {
        isValid: false,
        conflicts: conflictSteels,
        maxModuleLength,
        suggestions: [
          {
            type: 'addLongerModule',
            description: `建议添加长度≥${Math.max(...conflictSteels.map(s => s.length))}mm的模数钢材`,
            requiredLength: Math.max(...conflictSteels.map(s => s.length))
          },
          {
            type: 'increaseWelding',
            description: `建议将最大焊接段数增加到≥${Math.ceil(Math.max(...conflictSteels.map(s => s.length)) / maxModuleLength)}段`,
            requiredSegments: Math.ceil(Math.max(...conflictSteels.map(s => s.length)) / maxModuleLength)
          }
        ]
      };
    }

    return { isValid: true, conflicts: [], suggestions: [] };
  }

  /**
   * 获取约束条件的友好显示名称
   */
  getDisplayNames() {
    return {
      wasteThreshold: '废料阈值',
      weldingSegments: '焊接段数',
      maxIterations: '最大迭代次数',
      timeLimit: '计算时间限制',
      allowDynamicWelding: '动态焊接约束',
      maxWeldingSegments: '最大焊接段数',
      minWeldingSegments: '最小焊接段数',
      weldingCostPerSegment: '每段焊接的成本系数',
      weldingTimePerSegment: '每段焊接的时间成本(秒)'
    };
  }

  /**
   * 获取约束条件的详细说明
   */
  getDescriptions() {
    return {
      wasteThreshold: '余料长度小于此值时视为废料，无法再次利用',
      weldingSegments: '单根设计钢材允许的焊接段数，1段表示不允许焊接',
      maxIterations: '算法最大迭代次数，超过后返回当前最优解',
      timeLimit: '算法计算的最大允许时间，超时后返回当前最优解',
      allowDynamicWelding: '是否允许动态调整焊接段数',
      maxWeldingSegments: '单根设计钢材允许的最大焊接段数，1段表示不允许焊接',
      minWeldingSegments: '单根设计钢材允许的最小焊接段数，1段表示不允许焊接',
      weldingCostPerSegment: '每段焊接的成本系数，影响焊接成本',
      weldingTimePerSegment: '每段焊接的时间成本(秒)，影响焊接时间'
    };
  }
}

// ==================== 切割和优化结果 ====================

/**
 * 切割详情类型 - V3.0增强
 */
class CuttingDetail {
  constructor({
    sourceType, // 'module' | 'remainder'
    sourceId,
    sourceLength,
    moduleType = '',
    moduleLength = 0,
    designId,
    length,
    quantity,
    remainderInfo = null, // 余料信息
    weldingCount = 1      // 焊接段数
  }) {
    this.sourceType = sourceType;
    this.sourceId = sourceId;
    this.sourceLength = sourceLength;
    this.moduleType = moduleType;
    this.moduleLength = moduleLength;
    this.designId = designId;
    this.length = length;
    this.quantity = quantity;
    this.remainderInfo = remainderInfo;
    this.weldingCount = weldingCount;
  }
}

/**
 * CuttingPlan 切割计划类型 - V3.0 增强
 * 注意：newRemainders 字段只允许 type !== 'waste' 的余料对象，禁止混入废料对象！
 */
class CuttingPlan {
  constructor({
    sourceType,
    sourceId,
    sourceDescription,
    sourceLength,
    moduleType = '',
    moduleLength = 0,
    cuts = [],
    newRemainders = [],
    pseudoRemainders = [], // 伪余料
    realRemainders = [],   // 真余料
    waste = 0,
    usedRemainders = []
  }) {
    this.sourceType = sourceType;
    this.sourceId = sourceId;
    this.sourceDescription = sourceDescription;
    this.sourceLength = sourceLength;
    this.moduleType = moduleType;
    this.moduleLength = moduleLength;
    this.cuts = cuts;
    this.newRemainders = newRemainders;
    this.pseudoRemainders = pseudoRemainders;
    this.realRemainders = realRemainders;
    this.waste = waste;
    this.usedRemainders = usedRemainders;
  }
}

/**
 * 优化解决方案 - V3.0增强
 */
class OptimizationSolution {
  constructor({
    cuttingPlans = [],
    totalModuleUsed = 0,
    totalWaste = 0,
    totalPseudoRemainder = 0,  // 伪余料总长度
    totalRealRemainder = 0,    // 真余料总长度
    details = [],
    lossRateBreakdown = {}     // 损耗率分解
  }) {
    this.cuttingPlans = cuttingPlans;
    this.totalModuleUsed = totalModuleUsed;
    this.totalWaste = totalWaste;
    this.totalPseudoRemainder = totalPseudoRemainder;
    this.totalRealRemainder = totalRealRemainder;
    this.details = details;
    this.lossRateBreakdown = lossRateBreakdown;
  }
}

/**
 * 优化结果类型 - V3.0重构
 */
class OptimizationResult {
  constructor({
    solutions = {},
    totalLossRate = 0,
    totalModuleUsed = 0,
    totalWaste = 0,
    totalPseudoRemainder = 0,
    totalRealRemainder = 0,
    totalMaterial = 0,
    executionTime = 0,
    lossRateValidation = null, // 损耗率验证结果
    constraintValidation = null // 约束验证结果
  }) {
    this.solutions = solutions;
    this.totalLossRate = totalLossRate;
    this.totalModuleUsed = totalModuleUsed;
    this.totalWaste = totalWaste;
    this.totalPseudoRemainder = totalPseudoRemainder;
    this.totalRealRemainder = totalRealRemainder;
    this.totalMaterial = totalMaterial;
    this.executionTime = executionTime;
    this.lossRateValidation = lossRateValidation;
    this.constraintValidation = constraintValidation;
  }
}

// ==================== 损耗率计算 ====================
// 🔧 统一架构重构：损耗率计算已整合到StatisticsCalculator中
// 这里保留注释作为架构变更记录

/*
 * 📋 架构变更记录 - V3.1统一计算器重构
 * 
 * 原LossRateCalculator类已被删除，所有损耗率计算功能已整合到StatisticsCalculator中：
 * - calculateSpecificationLossRate() -> StatisticsCalculator.calculateSpecificationLossRate()
 * - calculateTotalLossRate() -> StatisticsCalculator.calculateTotalLossRate()  
 * - validateLossRateCalculation() -> StatisticsCalculator.validateLossRateCalculation()
 * - calculateTotalModuleMaterial() -> 直接使用统计结果，不再需要独立计算
 * 
 * 优势：
 * ✅ 单一数据源，消除架构冲突
 * ✅ 统一精度控制和错误处理
 * ✅ 消除"幽灵调用点"问题
 * ✅ 完美的架构统一性
 */

// ==================== 导出定义 ====================

module.exports = {
  // 基础类型
  DesignSteel,
  ModuleSteel,
  RemainderV3,
  
  // 约束系统
  OptimizationConstraints,
  
  // 切割和结果
  CuttingDetail,
  CuttingPlan,
  OptimizationSolution,
  OptimizationResult,
  
  // 🔧 统一架构：LossRateCalculator已整合到StatisticsCalculator中
  // 计算器相关功能请使用 StatisticsCalculator
  
  // 常量
  REMAINDER_TYPES: {
    PENDING: 'pending',
    PSEUDO: 'pseudo',
    REAL: 'real',
    WASTE: 'waste'
  },
  
  SOURCE_TYPES: {
    MODULE: 'module',
    REMAINDER: 'remainder'
  }
}; 