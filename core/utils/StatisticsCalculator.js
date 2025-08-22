/**
 * 统计计算器 - V3.0 单一数据权威源
 * 所有统计计算都在这里完成，其他地方只引用结果
 */

const { REMAINDER_TYPES } = require('../../api/types/index');

class StatisticsCalculator {
  constructor() {
    this.PRECISION = 4; // 统一数值精度
  }

  /**
   * 🎯 主计算方法：基于最终状态计算所有统计数据
   * 这是唯一的统计计算源，确保数据一致性
   */
  calculateAllStatistics(solutions, remainderManager) {
    console.log('🧮 StatisticsCalculator: 开始统一统计计算...');
    
    // 1. 确保所有余料状态已最终确定
    const finalRemainderStats = remainderManager.finalizeRemainders();
    
    // 2. 分组计算各规格的统计数据
    const specificationStats = {};
    let globalStats = {
      totalModuleCount: 0,
      totalModuleLength: 0,
      totalDesignLength: 0,
      totalWaste: 0,
      totalRealRemainder: 0,
      totalPseudoRemainder: 0,
      overallLossRate: 0,
      materialUtilizationRate: 0
    };

    Object.entries(solutions).forEach(([groupKey, solution]) => {
      const specStats = this.calculateSpecificationStatistics(solution, groupKey, remainderManager);
      specificationStats[groupKey] = specStats;
      
      // 累加到全局统计
      globalStats.totalModuleCount += specStats.moduleUsed;
      globalStats.totalModuleLength += specStats.totalMaterial;
      globalStats.totalDesignLength += specStats.designLength;
      globalStats.totalWaste += specStats.waste;
      globalStats.totalRealRemainder += specStats.realRemainder;
      globalStats.totalPseudoRemainder += specStats.pseudoRemainder;
    });

    // 3. 计算全局损耗率和利用率
    if (globalStats.totalModuleLength > 0) {
      globalStats.overallLossRate = parseFloat(
        ((globalStats.totalWaste + globalStats.totalRealRemainder) / globalStats.totalModuleLength * 100).toFixed(this.PRECISION)
      );
      globalStats.materialUtilizationRate = parseFloat(
        (globalStats.totalDesignLength / globalStats.totalModuleLength * 100).toFixed(this.PRECISION)
      );
    }

    // 4. 数据一致性验证
    const consistencyCheck = this.validateDataConsistency(globalStats, specificationStats, finalRemainderStats);

    console.log('✅ StatisticsCalculator: 统计计算完成');
    console.log(`📊 全局汇总: 模数${globalStats.totalModuleCount}根, 损耗率${globalStats.overallLossRate}%, 利用率${globalStats.materialUtilizationRate}%`);

    return {
      globalStats,
      specificationStats,
      remainderStats: finalRemainderStats,
      consistencyCheck
    };
  }

  /**
   * 🔢 计算单个规格的统计数据
   */
  calculateSpecificationStatistics(solution, groupKey, remainderManager) {
    let moduleUsed = 0;
    let totalMaterial = 0;
    let designLength = 0;
    let waste = 0;
    let realRemainder = 0;
    let pseudoRemainder = 0;
    
    const usedModuleIds = new Set();

    console.log(`🔍 计算 ${groupKey} 规格统计...`);

    // 遍历切割计划进行精确统计
    solution.cuttingPlans?.forEach(plan => {
      // 1. 统计模数钢材使用
      if (plan.sourceType === 'module' && plan.sourceId) {
        if (!usedModuleIds.has(plan.sourceId)) {
          totalMaterial += plan.sourceLength;
          usedModuleIds.add(plan.sourceId);
        }
      }

      // 2. 统计设计长度（成品）
      plan.cuts?.forEach(cut => {
        designLength += cut.length * cut.quantity;
      });

      // 3. 统计废料（只统计plan级别的废料，避免重复）
      if (plan.waste && plan.waste > 0) {
        waste += plan.waste;
      }
    });

    moduleUsed = usedModuleIds.size;

    // 4. 从余料管理器获取该规格的最终余料统计
    const remainderSpecStats = remainderManager.getSpecificationStatistics(groupKey);
    realRemainder = remainderSpecStats.realLength;
    pseudoRemainder = remainderSpecStats.pseudoRemainders; // 数量，不是长度

    // 5. 计算规格级损耗率和利用率
    const lossRate = totalMaterial > 0 
      ? parseFloat(((waste + realRemainder) / totalMaterial * 100).toFixed(this.PRECISION))
      : 0;
    
    const utilization = totalMaterial > 0 
      ? parseFloat((designLength / totalMaterial * 100).toFixed(this.PRECISION))
      : 0;

    const stats = {
      groupKey,
      moduleUsed,
      totalMaterial,
      designLength,
      waste,
      realRemainder,
      pseudoRemainder,
      lossRate,
      utilization,
      cuttingPlansCount: solution.cuttingPlans?.length || 0
    };

    console.log(`  📋 ${groupKey}: ${moduleUsed}根模数, 废料${waste}mm, 真余料${realRemainder}mm, 损耗率${lossRate}%`);

    return stats;
  }

  /**
   * 🔬 数据一致性验证
   */
  validateDataConsistency(globalStats, specificationStats, remainderStats) {
    const errors = [];
    const warnings = [];

    // 验证总数是否等于各部分之和
    let sumModuleCount = 0;
    let sumTotalMaterial = 0;
    let sumWaste = 0;
    let sumRealRemainder = 0;

    Object.values(specificationStats).forEach(specStats => {
      sumModuleCount += specStats.moduleUsed;
      sumTotalMaterial += specStats.totalMaterial;
      sumWaste += specStats.waste;
      sumRealRemainder += specStats.realRemainder;
    });

    // 检查数值一致性（允许小数点误差）
    const tolerance = 0.01;
    
    if (Math.abs(globalStats.totalModuleCount - sumModuleCount) > tolerance) {
      errors.push(`模数钢材数量不一致: 全局${globalStats.totalModuleCount} vs 规格求和${sumModuleCount}`);
    }

    if (Math.abs(globalStats.totalModuleLength - sumTotalMaterial) > tolerance) {
      errors.push(`模数钢材长度不一致: 全局${globalStats.totalModuleLength} vs 规格求和${sumTotalMaterial}`);
    }

    if (Math.abs(globalStats.totalWaste - sumWaste) > tolerance) {
      errors.push(`废料统计不一致: 全局${globalStats.totalWaste} vs 规格求和${sumWaste}`);
    }

    if (Math.abs(globalStats.totalRealRemainder - sumRealRemainder) > tolerance) {
      errors.push(`真余料统计不一致: 全局${globalStats.totalRealRemainder} vs 规格求和${sumRealRemainder}`);
    }

    // 检查余料管理器统计一致性
    if (Math.abs(remainderStats.totalWaste - sumWaste) > tolerance) {
      warnings.push(`余料管理器废料统计与切割计划不一致: 管理器${remainderStats.totalWaste} vs 计划${sumWaste}`);
    }

    // 检查损耗率合理性
    if (globalStats.overallLossRate < 0 || globalStats.overallLossRate > 100) {
      errors.push(`损耗率超出合理范围: ${globalStats.overallLossRate}%`);
    }

    if (globalStats.overallLossRate > 50) {
      warnings.push(`损耗率偏高: ${globalStats.overallLossRate}%，请检查优化配置`);
    }

    // 检查材料利用率合理性
    if (globalStats.materialUtilizationRate > 100) {
      errors.push(`材料利用率超过100%: ${globalStats.materialUtilizationRate}%，存在计算错误`);
    }

    console.log(`🔍 数据一致性检查: ${errors.length}个错误, ${warnings.length}个警告`);
    
    if (errors.length > 0) {
      console.error('❌ 发现数据一致性错误:');
      errors.forEach(error => console.error(`  - ${error}`));
    }
    
    if (warnings.length > 0) {
      console.warn('⚠️ 数据一致性警告:');
      warnings.forEach(warning => console.warn(`  - ${warning}`));
    }

    return {
      isConsistent: errors.length === 0,
      errors,
      warnings,
      checkTime: new Date().toISOString(),
      tolerance
    };
  }

  /**
   * 🎨 构建图表数据
   */
  buildChartData(specificationStats) {
    const colors = ['#0088FE', '#00C49F', '#FFBB28', '#FF8042', '#8884D8'];
    const lossRateData = [];
    const pieData = [];

    Object.entries(specificationStats).forEach(([groupKey, specStats], index) => {
      const [specification, crossSectionStr] = groupKey.split('_');
      const displayName = `${specification}(${crossSectionStr}mm²)`;

      lossRateData.push({
        specification: displayName,
        groupKey: groupKey,
        lossRate: specStats.lossRate,
        moduleUsed: specStats.moduleUsed,
        waste: specStats.waste,
        realRemainder: specStats.realRemainder,
        pseudoRemainder: specStats.pseudoRemainder,
        utilization: specStats.utilization
      });

      pieData.push({
        name: displayName,
        value: specStats.moduleUsed,
        fill: colors[index % colors.length]
      });
    });

    return { lossRateData, pieData };
  }

  /**
   * 🔢 构建模数钢材使用统计
   */
  buildModuleUsageStats(solutions) {
    const usage = {};
    let grandTotalCount = 0;
    let grandTotalLength = 0;

    Object.entries(solutions).forEach(([groupKey, solution]) => {
      const [specification, crossSectionStr] = groupKey.split('_');
      
      const moduleStats = {};
      let groupTotalCount = 0;
      let groupTotalLength = 0;

      // 统计该规格下的模数钢材使用情况
      const usedModules = new Map(); // sourceId -> { length, count }
      
      solution.cuttingPlans?.forEach(plan => {
        if (plan.sourceType === 'module' && plan.sourceId) {
          const length = plan.sourceLength;
          if (!usedModules.has(plan.sourceId)) {
            usedModules.set(plan.sourceId, { length, count: 1 });
          }
        }
      });

      // 按长度分组统计
      for (const { length } of usedModules.values()) {
        if (!moduleStats[length]) {
          moduleStats[length] = { count: 0, totalLength: 0 };
        }
        moduleStats[length].count += 1;
        moduleStats[length].totalLength += length;
        
        groupTotalCount += 1;
        groupTotalLength += length;
      }

      usage[groupKey] = {
        specification: specification,
        crossSection: parseInt(crossSectionStr),
        moduleBreakdown: moduleStats,
        groupTotal: { count: groupTotalCount, totalLength: groupTotalLength }
      };

      grandTotalCount += groupTotalCount;
      grandTotalLength += groupTotalLength;
    });

    return {
      bySpecification: usage,
      grandTotal: { count: grandTotalCount, totalLength: grandTotalLength }
    };
  }

  // ==================== 损耗率计算方法 ====================
  // 🔧 统一架构：整合原LossRateCalculator的所有功能

  /**
   * 🎯 计算单规格损耗率
   * 公式：(真余料+废料)/该规格模数钢材总长度*100%
   * 注意：这里直接使用统计结果，不再依赖外部计算
   */
  calculateSpecificationLossRate(specificationStats) {
    const { waste, realRemainder, totalMaterial } = specificationStats;
    
    if (totalMaterial === 0) return 0;
    
    const totalWasteAndReal = waste + realRemainder;
    return parseFloat(((totalWasteAndReal / totalMaterial) * 100).toFixed(this.PRECISION));
  }

  /**
   * 🎯 计算总损耗率
   * 公式：各规格真余料废料总和/各规格模数钢材总长度总和*100%
   */
  calculateTotalLossRate(allSpecificationStats) {
    let totalWasteAndReal = 0;
    let totalModuleMaterial = 0;

    Object.values(allSpecificationStats).forEach(specStats => {
      totalWasteAndReal += specStats.waste + specStats.realRemainder;
      totalModuleMaterial += specStats.totalMaterial;
    });

    if (totalModuleMaterial === 0) return 0;

    return parseFloat(((totalWasteAndReal / totalModuleMaterial) * 100).toFixed(this.PRECISION));
  }

  /**
   * 🔍 验证损耗率计算正确性
   * 检查加权平均是否等于总损耗率
   */
  validateLossRateCalculation(allSpecificationStats) {
    const totalLossRate = this.calculateTotalLossRate(allSpecificationStats);
    
    // 计算加权平均
    let weightedSum = 0;
    let totalWeight = 0;

    const specResults = Object.values(allSpecificationStats).map(specStats => {
      const specLossRate = this.calculateSpecificationLossRate(specStats);
      const weight = specStats.totalMaterial;
      
      weightedSum += specLossRate * weight;
      totalWeight += weight;

      return {
        lossRate: specLossRate,
        weight: weight,
        contribution: specLossRate * weight
      };
    });

    const weightedAverage = totalWeight > 0 ? 
      parseFloat((weightedSum / totalWeight).toFixed(this.PRECISION)) : 0;

    const difference = Math.abs(totalLossRate - weightedAverage);
    const ERROR_THRESHOLD = 0.01; // 误差阈值
    const isValid = difference <= ERROR_THRESHOLD;

    return {
      isValid,
      totalLossRate,
      weightedAverage,
      difference,
      specResults,
      errorMessage: isValid ? null : `损耗率计算存在误差: ${difference.toFixed(4)}%`
    };
  }

  /**
   * 🎯 获取规格的模数钢材总长度
   * 统一架构：直接从统计结果获取，不再依赖solution对象
   */
  getSpecificationTotalMaterial(specificationStats) {
    return specificationStats.totalMaterial || 0;
  }

  /**
   * 🎯 增强的损耗率分析
   * 提供详细的损耗率分解和分析
   */
  analyzeLossRateBreakdown(allSpecificationStats) {
    const analysis = {
      totalLossRate: this.calculateTotalLossRate(allSpecificationStats),
      specificationBreakdown: {},
      summary: {
        highLossSpecs: [],
        lowLossSpecs: [],
        averageLossRate: 0,
        totalMaterial: 0,
        totalWaste: 0,
        totalRealRemainder: 0
      }
    };

    let totalMaterial = 0;
    let totalWaste = 0;
    let totalRealRemainder = 0;
    let totalLossRateSum = 0;
    let specCount = 0;

    Object.entries(allSpecificationStats).forEach(([groupKey, specStats]) => {
      const specLossRate = this.calculateSpecificationLossRate(specStats);
      const contribution = (specStats.totalMaterial / Object.values(allSpecificationStats).reduce((sum, s) => sum + s.totalMaterial, 0)) * 100;

      analysis.specificationBreakdown[groupKey] = {
        lossRate: specLossRate,
        waste: specStats.waste,
        realRemainder: specStats.realRemainder,
        totalMaterial: specStats.totalMaterial,
        contribution: parseFloat(contribution.toFixed(2)),
        efficiency: parseFloat((100 - specLossRate).toFixed(2))
      };

      // 分类高低损耗率规格
      if (specLossRate > 10) {
        analysis.summary.highLossSpecs.push({ groupKey, lossRate: specLossRate });
      } else if (specLossRate < 3) {
        analysis.summary.lowLossSpecs.push({ groupKey, lossRate: specLossRate });
      }

      totalMaterial += specStats.totalMaterial;
      totalWaste += specStats.waste;
      totalRealRemainder += specStats.realRemainder;
      totalLossRateSum += specLossRate;
      specCount++;
    });

    analysis.summary.averageLossRate = specCount > 0 ? parseFloat((totalLossRateSum / specCount).toFixed(this.PRECISION)) : 0;
    analysis.summary.totalMaterial = totalMaterial;
    analysis.summary.totalWaste = totalWaste;
    analysis.summary.totalRealRemainder = totalRealRemainder;

    return analysis;
  }
}

module.exports = StatisticsCalculator; 