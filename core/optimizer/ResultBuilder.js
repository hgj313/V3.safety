/**
 * 结果构建器 - V3.0 统一数据构建
 * 使用统一的StatisticsCalculator进行计算，只负责数据结构构建
 */

const StatisticsCalculator = require('../utils/StatisticsCalculator');

class ResultBuilder {
  constructor() {
    this.statisticsCalculator = new StatisticsCalculator();
  }

  /**
   * 构建完整的优化结果
   * 使用统一的StatisticsCalculator进行所有计算
   */
  buildCompleteResult(solutions, designSteels, moduleSteels, remainderManager, executionTime = 0) {
    console.log('🏗️ ResultBuilder: 开始构建完整结果...');
    
    // 🎯 核心修复：使用统一的StatisticsCalculator进行所有统计计算
    const statisticsResult = this.statisticsCalculator.calculateAllStatistics(solutions, remainderManager);
    const { globalStats, specificationStats, remainderStats, consistencyCheck } = statisticsResult;
    
    // 🔧 统一架构关键修复：确保每个solution对象都有totalMaterial属性
    // 这样就消除了"幽灵调用点"问题
    Object.entries(solutions).forEach(([groupKey, solution]) => {
      const specStats = specificationStats[groupKey];
      if (specStats) {
        solution.totalMaterial = specStats.totalMaterial;
        solution.totalRealRemainder = specStats.realRemainder;
        solution.totalWaste = specStats.waste;
        solution.lossRate = specStats.lossRate;
        console.log(`🔧 设置 ${groupKey} 的 totalMaterial: ${specStats.totalMaterial}mm`);
      }
    });
    
    // 2. 构建图表数据
    const chartData = this.statisticsCalculator.buildChartData(specificationStats);
    
    // 3. 构建需求验证数据
    const requirementValidation = this.buildRequirementValidation(solutions, designSteels);
    
    // 4. 构建模数钢材使用统计
    const moduleUsageStats = this.statisticsCalculator.buildModuleUsageStats(solutions);
    
    // 5. 构建规格级别的详细统计
    const specificationDetails = this.buildSpecificationDetails(specificationStats);
    
    const result = {
      // 原有字段保持兼容性
      solutions: solutions,
      totalModuleUsed: globalStats.totalModuleCount,
      totalMaterial: globalStats.totalModuleLength,
      totalWaste: globalStats.totalWaste,
      totalRealRemainder: globalStats.totalRealRemainder,
      totalPseudoRemainder: globalStats.totalPseudoRemainder,
      totalLossRate: globalStats.overallLossRate,
      executionTime: executionTime,
      
      // 新增的完整数据结构，前端直接使用
      completeStats: {
        global: globalStats,
        chartData: chartData,
        requirementValidation: requirementValidation,
        moduleUsageStats: moduleUsageStats,
        specificationDetails: specificationDetails,
        remainderStats: remainderStats,
        consistencyCheck: consistencyCheck
      }
    };
    
    console.log(`✅ ResultBuilder: 结果构建完成`);
    console.log(`📊 全局统计: 模数钢材${globalStats.totalModuleCount}根, 损耗率${globalStats.overallLossRate}%`);
    
    if (!consistencyCheck.isConsistent) {
      console.warn('⚠️ ResultBuilder: 数据一致性检查失败:', consistencyCheck.errors);
    }
    
    return result;
  }





  /**
   * 构建需求验证数据
   * 验证所有设计钢材需求是否得到满足
   */
  buildRequirementValidation(solutions, designSteels) {
    const validation = [];
    let allSatisfied = true;

    designSteels.forEach(steel => {
      let produced = 0;
      
      // 统计该设计钢材的实际生产数量
      Object.values(solutions).forEach(solution => {
        solution.cuttingPlans?.forEach(plan => {
          plan.cuts?.forEach(cut => {
            if (cut.designId === steel.id) {
              produced += cut.quantity;
            }
          });
        });
      });

      const satisfied = produced >= steel.quantity;
      if (!satisfied) allSatisfied = false;

      validation.push({
        key: steel.id,
        id: steel.displayId || steel.id,
        specification: steel.specification || '未知规格',
        crossSection: steel.crossSection,
        length: steel.length,
        quantity: steel.quantity,
        produced: produced,
        satisfied: satisfied,
        difference: produced - steel.quantity,
        groupKey: steel.groupKey,
        satisfactionRate: steel.quantity > 0 ? parseFloat((produced / steel.quantity * 100).toFixed(this.PRECISION)) : 0
      });
    });

    return {
      items: validation,
      summary: {
        total: validation.length,
        satisfied: validation.filter(item => item.satisfied).length,
        unsatisfied: validation.filter(item => !item.satisfied).length,
        allSatisfied: allSatisfied,
        overallSatisfactionRate: validation.length > 0 
          ? parseFloat((validation.filter(item => item.satisfied).length / validation.length * 100).toFixed(this.PRECISION))
          : 0
      }
    };
  }



  /**
   * 构建规格级别的详细统计
   * 使用StatisticsCalculator计算的数据
   */
  buildSpecificationDetails(specificationStats) {
    const details = {};

    Object.entries(specificationStats).forEach(([groupKey, specStats]) => {
      const [specification, crossSectionStr] = groupKey.split('_');
      
      details[groupKey] = {
        specification: specification,
        crossSection: parseInt(crossSectionStr),
        displayName: `${specification}(${crossSectionStr}mm²)`,
        stats: {
          moduleUsed: specStats.moduleUsed,
          totalMaterial: specStats.totalMaterial,
          designLength: specStats.designLength,
          waste: specStats.waste,
          realRemainder: specStats.realRemainder,
          pseudoRemainder: specStats.pseudoRemainder,
          lossRate: specStats.lossRate,
          utilization: specStats.utilization
        },
        cuttingPlansCount: specStats.cuttingPlansCount
      };
    });

    return details;
  }


}

module.exports = ResultBuilder; 