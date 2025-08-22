/**
 * 钢材优化器 V3.0 - 核心算法引擎
 * 集成新的余料系统、约束W和模块化架构
 */

const { 
  DesignSteel, 
  ModuleSteel, 
  RemainderV3, 
  OptimizationSolution,
  OptimizationResult,
  CuttingPlan,
  CuttingDetail,
  REMAINDER_TYPES,
  SOURCE_TYPES,
  OptimizationConstraints
} = require('../../api/types');

const ResultBuilder = require('./ResultBuilder');
const RemainderManager = require('../remainder/RemainderManager');
const ConstraintValidator = require('../constraints/ConstraintValidator');
const ParallelOptimizationMonitor = require('./ParallelOptimizationMonitor');
const constraintManager = require('../config/ConstraintManager');
const { v4: uuidv4 } = require('uuid');

class SteelOptimizerV3 {
  constructor(designSteels, moduleSteels, constraints) {
    this.designSteels = designSteels;
    this.moduleSteels = moduleSteels; // 保留原始模数钢材作为模板
    this.constraints = constraints;
    this.startTime = Date.now();
    // 🔧 统一架构：损耗率计算已整合到StatisticsCalculator中
    this.resultBuilder = new ResultBuilder(); // 统一的结果构建器
    this.moduleCounters = {}; // 模数钢材计数器
    
    // V3新增：规格化模数钢材池管理
    this.moduleSteelPools = new Map(); // 按规格+截面面积组合管理的池
    this.availableLengths = this.extractAvailableLengths(moduleSteels);
    
    console.log('🎯 V3规格化优化器初始化完成');
    console.log(`📊 可用模数钢材长度: ${this.availableLengths.join(', ')}mm`);
    
    // 核心组件
    this.remainderManager = new RemainderManager(constraints.wasteThreshold);
    this.constraintValidator = new ConstraintValidator();
    this.parallelMonitor = new ParallelOptimizationMonitor();
    
    // 计数器
    this.executionStats = {
      totalCuts: 0,
      remaindersGenerated: 0,
      remaindersReused: 0,
      weldingOperations: 0
    };
  }

  /**
   * 从原始模数钢材中提取可用长度
   */
  extractAvailableLengths(moduleSteels) {
    const lengths = moduleSteels.map(m => m.length);
    // 🔧 关键修复：必须按升序排列，这样 find() 才能正确选择最短的合格模数钢材
    return [...new Set(lengths)].sort((a, b) => a - b);
  }

  /**
   * 获取或创建规格化模数钢材池
   */
  getOrCreateModuleSteelPool(groupKey) {
    if (!this.moduleSteelPools.has(groupKey)) {
      const [specification, crossSection] = this.parseGroupKey(groupKey);
      const pool = new SpecificationModuleSteelPool(
        specification,
        parseFloat(crossSection),
        this.availableLengths
      );
      this.moduleSteelPools.set(groupKey, pool);
      console.log(`🔧 创建规格化模数钢材池: ${groupKey}`);
    }
    return this.moduleSteelPools.get(groupKey);
  }

  /**
   * 解析组合键
   */
  parseGroupKey(groupKey) {
    const parts = groupKey.split('_');
    return [parts[0], parts[1]]; // [specification, crossSection]
  }

  /**
   * 主优化入口
   */
  async optimize() {
    console.log('🚀 启动钢材优化算法 V3.0');
    
    try {
      // 1. 预验证约束条件
      const validation = this.constraintValidator.validateAllConstraints(
        this.designSteels, 
        this.moduleSteels, 
        this.constraints
      );
      
      if (!validation.isValid) {
        console.error('❌ 约束条件验证失败');
        return {
          success: false,
          error: '约束条件验证失败',
          validation: validation // 确保返回完整的验证对象
        };
      }

      // 2. 按规格分组优化
      const solutions = await this.optimizeByGroups();
      
      // 🔧 关键修复：在所有组合优化完成后，统一进行余料最终处理
      console.log('\n🏁 所有组合优化完成，开始余料最终处理...');
      const finalRemainderStats = this.remainderManager.finalizeRemainders();
      
      // 🔧 关键修复：更新所有切割计划中的余料状态
      this.updateCuttingPlansRemainderStatus(solutions);
      
      // 3. MW-CD交换已在各并行任务内部完成，无需全局处理
      console.log('\n✅ MW-CD交换优化已在各组合内部完成');
      
      // 4. 计算损耗率和验证
      const result = this.buildOptimizationResult(solutions, validation);
      
      // 5. 🔧 统一架构：使用StatisticsCalculator进行损耗率验证
      // 注意：验证需要使用统计结果而不是原始solutions
      const statisticsResult = this.resultBuilder.statisticsCalculator.calculateAllStatistics(solutions, this.remainderManager);
      const lossRateValidation = this.resultBuilder.statisticsCalculator.validateLossRateCalculation(statisticsResult.specificationStats);
      if (!lossRateValidation.isValid) {
        console.warn('⚠️ 损耗率计算验证失败:', lossRateValidation.errorMessage);
      }
      
      result.lossRateValidation = lossRateValidation;
      result.constraintValidation = validation;
      
      console.log('🎉 优化完成');
      return {
        success: true,
        result: result,
        stats: this.executionStats
      };
      
    } catch (error) {
      console.error('❌ 优化过程出错:', error);
      return {
        success: false,
        error: error.message,
        stack: error.stack
      };
    }
  }

  /**
   * V3并行计算框架：按规格+截面面积组合并行优化
   */
  async optimizeByGroups() {
    const groups = this.groupBySpecificationAndCrossSection();
    const groupKeys = Object.keys(groups);
    
    // 启动性能监控
    this.parallelMonitor.startMonitoring(groupKeys.length);
    
    console.log(`🚀 启动V3并行计算框架，共${groupKeys.length}个规格组合将并行优化`);
    
    // 创建并行优化任务（带监控）
    const parallelTasks = groupKeys.map((groupKey, index) => 
      this.createMonitoredParallelTask(groupKey, groups[groupKey], index)
    );
    
    // 并行执行所有优化任务
    const parallelResults = await Promise.allSettled(parallelTasks);
    
    // 合并并行结果
    const solutions = {};
    const parallelStats = {
      successful: 0,
      failed: 0,
      totalGroups: groupKeys.length,
      errors: []
    };
    
    parallelResults.forEach((result, index) => {
      const groupKey = groupKeys[index];
      
      if (result.status === 'fulfilled') {
        solutions[groupKey] = result.value.solution;
        parallelStats.successful++;
        this.parallelMonitor.recordTaskCompletion(index, result.value.stats);
        
        // 🔧 关键修复：合并独立余料管理器的数据到主余料管理器
        if (result.value.remainderManager) {
          this.mergeRemainderManager(result.value.remainderManager, groupKey);
        }
      } else {
        parallelStats.failed++;
        parallelStats.errors.push({
          groupKey: groupKey,
          error: result.reason.message || result.reason
        });
        this.parallelMonitor.recordTaskFailure(index, result.reason.message || result.reason);
        
        // 为失败的任务创建空解决方案
        solutions[groupKey] = new OptimizationSolution({});
      }
    });
    
    // 完成监控并生成报告
    const performanceReport = this.parallelMonitor.finishMonitoring();
    this.parallelMonitor.printPerformanceReport();
    
    console.log(`🏁 V3并行计算完成: ${parallelStats.successful}成功/${parallelStats.failed}失败，共${parallelStats.totalGroups}个组合`);
    
    // 如果有失败的任务，记录详细信息
    if (parallelStats.failed > 0) {
      console.warn('⚠️ 并行计算中的失败任务:', parallelStats.errors);
    }
    
    return solutions;
  }

  /**
   * 创建带监控的并行任务
   */
  async createMonitoredParallelTask(groupKey, steels, taskIndex) {
    // 记录任务开始
    this.parallelMonitor.recordTaskStart(taskIndex, groupKey, steels.length);
    
    try {
      return await this.createParallelOptimizationTask(groupKey, steels);
    } catch (error) {
      // 监控器会在上层记录失败
      throw error;
    }
  }

  /**
   * 创建独立的并行优化任务
   */
  async createParallelOptimizationTask(groupKey, steels) {
    const taskStartTime = Date.now();
    
    try {
      // 为每个并行任务创建独立的余料管理器
      const independentRemainderManager = this.createIndependentRemainderManager(groupKey);
      
      // 创建任务专用的统计信息
      const taskStats = {
        groupKey: groupKey,
        steelsCount: steels.length,
        cuts: 0,
        remaindersGenerated: 0,
        remaindersReused: 0,
        weldingOperations: 0
      };
      
      console.log(`🔧 并行任务启动: ${groupKey} (${steels.length}种设计钢材)`);
      
      // 执行独立优化
      const solution = await this.optimizeGroupIndependently(
        steels, 
        groupKey, 
        independentRemainderManager,
        taskStats
      );
      
      // 🔧 修复：此处不再调用finalizeRemainders，因为要在所有组合完成后统一处理
      // independentRemainderManager.finalizeRemainders();
      
      const executionTime = Date.now() - taskStartTime;
      taskStats.executionTime = executionTime;
      
      return {
        solution: solution,
        stats: taskStats,
        remainderManager: independentRemainderManager
      };
      
    } catch (error) {
      console.error(`💥 并行任务异常: ${groupKey} - ${error.message}`);
      throw new Error(`并行优化失败[${groupKey}]: ${error.message}`);
    }
  }

  /**
   * 创建独立的余料管理器实例
   */
  createIndependentRemainderManager(groupKey) {
    const independentManager = new RemainderManager(this.constraints.wasteThreshold);
    
    // 预初始化该组合的余料池
    independentManager.initializePool(groupKey);
    
    console.log(`🔧 为${groupKey}创建独立余料管理器`);
    return independentManager;
  }

  /**
   * 🔧 关键修复：合并独立余料管理器的数据到主余料管理器
   */
  mergeRemainderManager(independentManager, groupKey) {
    console.log(`🔄 合并独立余料管理器数据: ${groupKey}`);
    
    // 获取独立管理器中的余料池
    const independentPool = independentManager.remainderPools[groupKey] || [];
    
    if (independentPool.length > 0) {
      // 确保主余料管理器有对应的池
      this.remainderManager.initializePool(groupKey);
      
      // 合并余料到主管理器
      independentPool.forEach(remainder => {
        this.remainderManager.remainderPools[groupKey].push(remainder);
        console.log(`  ➕ 合并余料: ${remainder.id} (${remainder.length}mm, 类型: ${remainder.type})`);
      });
      
      console.log(`✅ 已合并 ${independentPool.length} 个余料到主管理器`);
    } else {
      console.log(`  - ${groupKey} 组合没有余料需要合并`);
    }
  }

  /**
   * 🔧 关键修复：更新所有切割计划中的余料状态
   * 在finalizeRemainders之后，将切割计划中pending状态的余料更新为正确状态
   */
  updateCuttingPlansRemainderStatus(solutions) {
    console.log('\n🔄 更新切割计划中的余料状态...');
    
    let updatedCount = 0;
    
    // 获取所有已最终化的余料
    const allFinalizedRemainders = this.remainderManager.getAllRemainders();
    const remainderMap = new Map();
    
    allFinalizedRemainders.forEach(remainder => {
      remainderMap.set(remainder.id, remainder);
    });
    
    console.log(`📋 主余料管理器中共有 ${allFinalizedRemainders.length} 个已最终化的余料`);
    
    // 遍历所有解决方案和切割计划
    Object.entries(solutions).forEach(([groupKey, solution]) => {
      solution.cuttingPlans?.forEach((plan, planIndex) => {
        // 更新newRemainders中的状态
        if (plan.newRemainders && plan.newRemainders.length > 0) {
          plan.newRemainders.forEach((remainder, remainderIndex) => {
            const finalizedRemainder = remainderMap.get(remainder.id);
            if (finalizedRemainder && remainder.type !== finalizedRemainder.type) {
              console.log(`  🔄 更新余料状态: ${remainder.id} (${remainder.type} → ${finalizedRemainder.type})`);
              remainder.type = finalizedRemainder.type;
              updatedCount++;
            } else if (!finalizedRemainder) {
              // 如果在主余料管理器中找不到，可能是因为余料被使用了，需要特殊处理
              console.log(`  ⚠️ 余料 ${remainder.id} 在主管理器中未找到，可能已被使用`);
            }
          });
        }
        
        // 更新realRemainders中的状态
        if (plan.realRemainders && plan.realRemainders.length > 0) {
          plan.realRemainders.forEach(remainder => {
            const finalizedRemainder = remainderMap.get(remainder.id);
            if (finalizedRemainder && remainder.type !== finalizedRemainder.type) {
              console.log(`  🔄 更新真余料状态: ${remainder.id} (${remainder.type} → ${finalizedRemainder.type})`);
              remainder.type = finalizedRemainder.type;
              updatedCount++;
            }
          });
        }
      });
      
      // 🔧 修复：在更新余料状态后，不再需要重新计算统计数据
      // 所有统计将在优化流程的最后，由ResultBuilder统一计算
    });
    
    console.log(`✅ 切割计划余料状态更新完成，共更新 ${updatedCount} 个余料`);
  }

  /**
   * 独立优化单个规格+截面面积组合（并行安全）- V3.1 装箱算法
   * @description 使用首次适应递减(FFD)启发式算法，取代原有的贪婪策略，以获得更好的全局优化效果。
   */
  async optimizeGroupIndependently(steels, groupKey, remainderManager, taskStats) {
    const solution = new OptimizationSolution({});
    let unfulfilledDemands = this.createFlatDemandList(steels);

    // 按长度降序排序 (FFD 中的 "Decreasing")
    unfulfilledDemands.sort((a, b) => b.length - a.length);

    let binCount = 0;
    while (unfulfilledDemands.length > 0 && !this.isTimeExceeded()) {
      binCount++;
      const longestDemand = unfulfilledDemands[0];
      let sourceMaterial = null;
      let sourceType = '';
      let usedRemaindersList = [];

      // 步骤 1: 寻找一个合适的"箱子" - 优先使用余料
      // 寻找能容纳最长需求的、最省料的单个余料
      const bestRemainder = remainderManager.findBestSingleRemainder(longestDemand.length, groupKey);

      if (bestRemainder) {
        sourceMaterial = bestRemainder;
        sourceType = 'remainder';
        usedRemaindersList = [bestRemainder];
        remainderManager.useSingleRemainder(bestRemainder.id, groupKey); // 标记余料为已用
        taskStats.remaindersReused = (taskStats.remaindersReused || 0) + 1;
      } else {
        // 1b: 找不到合适的单个余料，尝试焊接组合
        const bestCombination = remainderManager.findBestRemainderCombination(
          longestDemand.length,
          groupKey,
          this.constraints.maxWeldingSegments // 关键：传入焊接段数约束
        );

        if (bestCombination) {
          // 找到了焊接组合
          sourceType = 'remainder';
          
          // 标记并移除
          bestCombination.remainders.forEach(r => r.markAsPseudo());
          remainderManager.removeRemaindersFromPool(bestCombination.indices, groupKey);

          // 创建一个临时的"焊接后材料"对象作为箱子
          sourceMaterial = { 
              id: bestCombination.remainders.map(r => r.id).join('+'), 
              length: bestCombination.totalLength, 
              crossSection: this.parseGroupKey(groupKey)[1] 
          };
          usedRemaindersList = bestCombination.remainders;

          taskStats.remaindersReused = (taskStats.remaindersReused || 0) + bestCombination.remainders.length;
          if (bestCombination.remainders.length > 1) {
              taskStats.weldingOperations = (taskStats.weldingOperations || 0) + 1;
          }
        } else {
          // 步骤 2: 如果所有余料方案都失败，则启用一根新的模数钢材
          const newModule = this.selectBestModule(longestDemand, groupKey, unfulfilledDemands, true); // 强制选择
          if (!newModule) {
            console.error(`❌ 严重错误: 无法为需求 ${longestDemand.length}mm (组: ${groupKey}) 选择模数钢材。`);
            break; 
          }
          sourceMaterial = newModule;
          sourceType = 'module';
          taskStats.moduleSteelsUsed = (taskStats.moduleSteelsUsed || 0) + 1;
          taskStats.totalModuleLength = (taskStats.totalModuleLength || 0) + newModule.length;
        }
      }
      
      // 步骤 3: "装箱" - 使用首次适应策略 (FFD 中的 "First Fit")
      const cuts = [];
      let remainingLength = sourceMaterial.length;
      const packedDemandUniqueIds = new Set();

      for (const demand of unfulfilledDemands) {
        if (demand.length <= remainingLength) {
          cuts.push({ designId: demand.id, length: demand.length, quantity: 1 });
          remainingLength -= demand.length;
          packedDemandUniqueIds.add(demand.uniqueId);

          solution.details.push(new CuttingDetail({
            sourceType: sourceType, sourceId: sourceMaterial.id, sourceLength: sourceMaterial.length,
            designId: demand.id, length: demand.length, quantity: 1, weldingCount: 1
          }));
        }
      }

      // 步骤 4: 从待办列表中移除已满足的需求
      unfulfilledDemands = unfulfilledDemands.filter(d => !packedDemandUniqueIds.has(d.uniqueId));

      // 步骤 5: 处理切割后产生的余料或废料
      let waste = 0;
      const newRemainders = [];
      if (remainingLength > 0) {
        const remainderToEvaluate = new RemainderV3({
          id: `${sourceMaterial.id}_rem`, length: remainingLength, type: REMAINDER_TYPES.PENDING,
          sourceChain: [sourceMaterial.id], crossSection: sourceMaterial.crossSection || this.parseGroupKey(groupKey)[1],
          specification: groupKey, createdAt: new Date().toISOString(),
          originalLength: sourceMaterial.length, parentId: sourceMaterial.id
        });
        
        const evalResult = remainderManager.evaluateAndProcessRemainder(remainderToEvaluate, groupKey, { source: `${sourceType}切割后` });

        if (evalResult.isWaste) {
          waste = evalResult.wasteLength;
          taskStats.wasteGenerated = (taskStats.wasteGenerated || 0) + waste;
        } else if (evalResult.isPendingRemainder) {
          newRemainders.push(remainderToEvaluate);
        }
      }

      // 步骤 6: 创建该"箱子"的切割计划
      const cuttingPlan = new CuttingPlan({
        sourceType: sourceType, sourceId: sourceMaterial.id,
        sourceDescription: `${groupKey}组合${sourceType} ${sourceMaterial.id}`,
        sourceLength: sourceMaterial.length,
        cuts: this.groupCuts(cuts), // 将单个切割合并为带数量的
        newRemainders: newRemainders, realRemainders: newRemainders,
        waste: waste, usedRemainders: usedRemaindersList,
        moduleLength: sourceType === 'module' ? sourceMaterial.length : undefined,
        moduleType: sourceType === 'module' ? sourceMaterial.name : undefined,
      });

      solution.cuttingPlans.push(cuttingPlan);
      taskStats.cuts += 1; // 这里的cut代表一次完整的切割计划
    }

    if (unfulfilledDemands.length > 0) {
        console.error(`❌ 需求未完全满足，组: ${groupKey}。剩余数量: ${unfulfilledDemands.length}`);
    }

    // 步骤 7: (可选) 在该组合内部执行后续优化
    console.log(`\n🔄 ${groupKey}组合内部MW-CD交换优化...`);
    const mwcdStats = await this.performInternalMWCDOptimization(solution, groupKey, remainderManager);
    if (mwcdStats.exchangesPerformed > 0) {
      console.log(`✅ ${groupKey}组合完成${mwcdStats.exchangesPerformed}次内部交换，效益提升${mwcdStats.totalBenefitGained.toFixed(2)}mm`);
    }
    
    this.mergeTaskStatsToSolution(solution, taskStats);
    console.log(`✅ 并行任务${groupKey}优化完成: 使用了 ${binCount} 个原材料, 完成 ${mwcdStats.exchangesPerformed} 次内部交换`);
    
    return solution;
  }

  /**
   * 将taskStats数据汇总到solution对象，解决"数据孤岛"问题
   */
  mergeTaskStatsToSolution(solution, taskStats) {
    console.log(`🔄 合并taskStats数据到solution对象:`);
    console.log(`  - taskStats.cuts: ${taskStats.cuts}`);
    console.log(`  - taskStats.moduleSteelsUsed: ${taskStats.moduleSteelsUsed}`);
    console.log(`  - taskStats.totalModuleLength: ${taskStats.totalModuleLength}`);
    console.log(`  - taskStats.wasteGenerated: ${taskStats.wasteGenerated}`);
    console.log(`  - taskStats.remaindersReused: ${taskStats.remaindersReused}`);
    console.log(`  - taskStats.weldingOperations: ${taskStats.weldingOperations}`);
    
    // 初始化solution的统计字段（如果不存在）
    solution.taskStats = solution.taskStats || {
      totalCuts: 0,
      totalModuleSteelsUsed: 0,
      totalModuleLength: 0,
      totalWasteGenerated: 0,
      totalRemaindersReused: 0,
      totalWeldingOperations: 0
    };
    
    // 累加taskStats数据到solution
    solution.taskStats.totalCuts += taskStats.cuts || 0;
    solution.taskStats.totalModuleSteelsUsed += taskStats.moduleSteelsUsed || 0;
    solution.taskStats.totalModuleLength += taskStats.totalModuleLength || 0;
    // 🎯 统一计算架构：废料统计已在使用时精确计算，此处不再重复累加
    // solution.taskStats.totalWasteGenerated += taskStats.wasteGenerated || 0;
    solution.taskStats.totalRemaindersReused += taskStats.remaindersReused || 0;
    solution.taskStats.totalWeldingOperations += taskStats.weldingOperations || 0;
    
    console.log(`✅ taskStats数据已合并到solution.taskStats`);
  }

  /**
   * [V3.2] 智能选择最佳模数钢材 (使用"向前看"决策)
   * @description 不再只选最短够用的，而是对每个候选模数钢材进行虚拟装箱，选出潜在利用率最高的。
   */
  selectBestModule(longestDemand, groupKey, unfulfilledDemands, force = false) {
    const pool = this.getOrCreateModuleSteelPool(groupKey);
    const availableModuleLengths = pool.availableLengths;

    // 1. 找出所有可行的候选模数钢材
    const candidates = availableModuleLengths.filter(len => len >= longestDemand.length);

    if (candidates.length === 0) {
      if (force) {
        const maxLength = Math.max(...availableModuleLengths);
        console.warn(`⚠️ (强制模式) 需求 ${longestDemand.length}mm 过长，所有模数钢材均不满足。将尝试使用最长的 ${maxLength}mm`);
        return pool.createSteel(maxLength);
      }
      return null;
    }

    if (candidates.length === 1) {
        const bestLength = candidates[0];
        console.log(`🎯 只有一个候选模数钢材 (${bestLength}mm), 直接选择。`);
        return pool.createSteel(bestLength);
    }
    
    let bestChoice = {
      length: -1,
      utilization: -1,
    };
    
    console.log(`🧐 [向前看] 开始评估 ${candidates.length} 个候选模数钢材...`);

    for (const candidateLength of candidates) {
      const { packedLength } = this.calculatePotentialUtilization(candidateLength, unfulfilledDemands);
      const utilization = packedLength / candidateLength;

      console.log(`  - 候选 ${candidateLength}mm: 潜力装载 ${packedLength}mm, 预期利用率 ${(utilization * 100).toFixed(2)}%`);

      if (utilization > bestChoice.utilization) {
        bestChoice.length = candidateLength;
        bestChoice.utilization = utilization;
      }
    }

    console.log(`✅ [向前看] 决策完成: 选择 ${bestChoice.length}mm (预期利用率 ${(bestChoice.utilization * 100).toFixed(2)}%)`);
    
    return pool.createSteel(bestChoice.length);
  }

  /**
   * [新增] 辅助方法 - 为"向前看"机制计算给定箱子的潜在利用率
   */
  calculatePotentialUtilization(binSize, demands) {
    let remainingLength = binSize;
    let packedLength = 0;

    for (const demand of demands) {
      if (demand.length <= remainingLength) {
        remainingLength -= demand.length;
        packedLength += demand.length;
      }
    }
    return { packedLength, remainingLength };
  }

  /**
   * 按规格+截面面积组合分组（真正的规格化设计）
   */
  groupBySpecificationAndCrossSection() {
    const groups = {};
    
    this.designSteels.forEach(steel => {
      const specification = steel.specification || '未知规格';
      const crossSection = Math.round(steel.crossSection);
      const groupKey = `${specification}_${crossSection}`;
      
      if (!groups[groupKey]) {
        groups[groupKey] = [];
      }
      groups[groupKey].push(steel);
    });
    
    console.log('🎯 按规格+截面面积组合分组结果:', Object.keys(groups));
    return groups;
  }

  /**
   * V3规格化改进：生成模数钢材ID
   */
  generateModuleId(groupKey) {
    if (!this.moduleCounters[groupKey]) {
      this.moduleCounters[groupKey] = 0;
    }
    this.moduleCounters[groupKey]++;
    return `${groupKey}_M${this.moduleCounters[groupKey]}`;
  }

  /**
   * 🔧 修复版：MW-CD交换优化算法
   * 基于规格匹配和效益分析的智能交换
   */
  async performMWCDOptimization(solutions) {
    console.log('🔄 执行改进版MW-CD交换优化算法');
    
    const stats = {
      totalMWFound: 0,
      totalCDFound: 0,
      validExchanges: 0,
      exchangesPerformed: 0,
      totalBenefitGained: 0
    };
    
    for (const [groupKey, solution] of Object.entries(solutions)) {
      if (this.isTimeExceeded()) break;
      
      console.log(`\n📋 分析 ${groupKey} 组合的MW-CD交换机会...`);
      
      const mwRemainders = this.findMWRemainders(solution);
      const cdPlans = this.findCDPlans(solution);
      
      stats.totalMWFound += mwRemainders.length;
      stats.totalCDFound += cdPlans.length;
      
      console.log(`  - 发现MW余料: ${mwRemainders.length}个`);
      console.log(`  - 发现CD计划: ${cdPlans.length}个`);
      
      if (mwRemainders.length === 0 || cdPlans.length === 0) {
        console.log(`  - 跳过${groupKey}组合（缺少交换对象）`);
        continue;
      }
      
      // 🔧 改进：按效益排序，优先执行最有价值的交换
      const exchangeOpportunities = this.analyzeExchangeOpportunities(mwRemainders, cdPlans, groupKey);
      
      for (const opportunity of exchangeOpportunities) {
        if (this.isTimeExceeded()) break;
        
        stats.validExchanges++;
        
        // 🔧 改进：执行完整的交换操作
        const exchangeResult = this.executeImprovedInterchange(solution, opportunity, groupKey);
        
        if (exchangeResult.success) {
          stats.exchangesPerformed++;
          stats.totalBenefitGained += exchangeResult.benefit;
          console.log(`  ✅ 交换成功: 效益提升 ${exchangeResult.benefit.toFixed(2)}mm`);
        } else {
          console.log(`  ❌ 交换失败: ${exchangeResult.reason}`);
        }
      }
    }
    
    console.log('\n📊 MW-CD交换优化统计:');
    console.log(`  - MW余料总数: ${stats.totalMWFound}`);
    console.log(`  - CD计划总数: ${stats.totalCDFound}`);
    console.log(`  - 有效交换机会: ${stats.validExchanges}`);
    console.log(`  - 实际执行交换: ${stats.exchangesPerformed}`);
    console.log(`  - 总效益提升: ${stats.totalBenefitGained.toFixed(2)}mm`);
    
    return stats;
  }

  /**
   * 🔧 改进：分析交换机会并按效益排序
   */
  analyzeExchangeOpportunities(mwRemainders, cdPlans, groupKey) {
    const opportunities = [];
      
      for (const mw of mwRemainders) {
        for (const cd of cdPlans) {
        // 🔧 改进：全面的交换可行性检查
        const feasibility = this.checkExchangeFeasibility(mw, cd, groupKey);
        
        if (feasibility.isFeasible) {
          opportunities.push({
            mw: mw,
            cd: cd,
            benefit: feasibility.benefit,
            totalBenefit: feasibility.benefit, // 🔧 修复：添加totalBenefit字段
            confidence: feasibility.confidence,
            reason: feasibility.reason
          });
        }
      }
    }
    
    // 按效益降序排列，优先执行最有价值的交换
    opportunities.sort((a, b) => b.benefit - a.benefit);
    
    console.log(`  - 分析出 ${opportunities.length} 个可行的交换机会`);
    
    return opportunities.slice(0, 5); // 限制最多5个交换，避免过度优化
  }

  /**
   * 🔧 改进：全面的交换可行性检查
   */
  checkExchangeFeasibility(mw, cd, groupKey) {
    const mwLength = mw.remainder.length;
    const cdTotalLength = cd.usedRemainders.reduce((sum, r) => sum + r.length, 0);
    const cdSegments = cd.usedRemainders.length;
    
    // 检查1：规格匹配（同一组合内才能交换）
    const mwGroupKey = mw.remainder.groupKey || groupKey;
    const cdGroupKey = cd.usedRemainders[0]?.groupKey || groupKey;
    
    if (mwGroupKey !== cdGroupKey) {
      return {
        isFeasible: false,
        reason: '规格不匹配'
      };
    }
    
    // 检查2：长度合理性（MW余料必须能满足CD的切割需求）
    const targetLength = cd.cuts[0]?.length || 0;
    if (mwLength < targetLength) {
      return {
        isFeasible: false,
        reason: 'MW余料长度不足'
      };
    }
    
    // 检查3：焊接约束（用MW替代CD可以减少焊接段数）
    if (cdSegments <= 1) {
      return {
        isFeasible: false,
        reason: 'CD计划已是单段，无需交换'
      };
    }
    
    // 检查4：效益计算（交换必须有正效益）
    const weldingCostSaved = (cdSegments - 1) * 50; // 假设每个焊接点成本50mm等效
    const materialWasteDiff = mwLength - cdTotalLength;
    const totalBenefit = weldingCostSaved - Math.abs(materialWasteDiff);
    
    if (totalBenefit <= 0) {
      return {
        isFeasible: false,
        reason: '交换效益为负',
        benefit: totalBenefit
      };
    }
    
    // 检查5：废料阈值约束
    const newWaste = mwLength - targetLength;
    if (newWaste >= this.constraints.wasteThreshold) {
      return {
        isFeasible: false,
        reason: '交换后废料过多'
      };
    }
    
    return {
      isFeasible: true,
      benefit: totalBenefit,
      confidence: Math.min(1.0, totalBenefit / 100), // 信心度基于效益大小
      reason: `可节省${cdSegments-1}个焊接点，效益${totalBenefit.toFixed(2)}mm`
    };
  }
  /**
   * 🔧 改进：执行完整的交换操作
   */
  executeImprovedInterchange(solution, opportunity, groupKey) {
    const { mw, cd, benefit } = opportunity;
    
    try {
      console.log(`    🔄 执行交换: MW(${mw.remainder.id}, ${mw.remainder.length}mm) ↔ CD(${cd.usedRemainders.map(r => r.id).join('+')}, ${cd.usedRemainders.reduce((sum, r) => sum + r.length, 0)}mm)`);
      
      // 1. 创建新的切割计划（用MW替代CD）
      const targetLength = cd.cuts[0]?.length || 0;
      const newWaste = mw.remainder.length - targetLength;
      
      const newCuttingPlan = new CuttingPlan({
        sourceType: 'remainder',
        sourceId: mw.remainder.id,
        sourceDescription: `${groupKey}组合MW余料 ${mw.remainder.id} (交换优化)`,
        sourceLength: mw.remainder.length,
        cuts: cd.cuts, // 复用原CD的切割需求
        usedRemainders: [mw.remainder],
        newRemainders: [],
        pseudoRemainders: [mw.remainder], // MW余料变为伪余料
        realRemainders: [],
        waste: newWaste
      });
      
      // 2. 移除原CD计划
      const cdIndex = solution.cuttingPlans.findIndex(plan => plan === cd);
      if (cdIndex !== -1) {
        solution.cuttingPlans.splice(cdIndex, 1);
      }
      
      // 3. 添加新计划
      solution.cuttingPlans.push(newCuttingPlan);
      
      // 4. 更新统计数据
      solution.totalWaste += newWaste;
      solution.totalRealRemainder -= mw.remainder.length;
      solution.totalPseudoRemainder += mw.remainder.length;
      
      // 5. 标记MW余料状态
      mw.remainder.markAsPseudo();
      
      // 6. 恢复CD使用的余料到池中（它们现在可以重新使用）
      cd.usedRemainders.forEach(remainder => {
        if (remainder.type === REMAINDER_TYPES.PSEUDO) {
          remainder.type = REMAINDER_TYPES.REAL; // 恢复为真余料
          this.remainderManager.remainderPools[groupKey].push(remainder);
        }
      });
      
      return {
        success: true,
        benefit: benefit,
        newPlan: newCuttingPlan
      };
      
    } catch (error) {
      console.error(`    ❌ 交换执行失败: ${error.message}`);
      return {
        success: false,
        reason: error.message
      };
    }
  }

  /**
   * 🔧 关键修复：在并行任务内部执行MW-CD交换优化
   * 这是正确的架构设计 - MW-CD应该在每个规格组合内部进行
   */
  async performInternalMWCDOptimization(solution, groupKey, remainderManager) {
    const stats = {
      totalMWFound: 0,
      totalCDFound: 0,
      validExchanges: 0,
      exchangesPerformed: 0,
      totalBenefitGained: 0,
      iterations: 0
    };
    
    console.log(`\n🔄 开始${groupKey}组合内部MW-CD交换优化...`);
    
    // 🔧 基于效益的收敛条件设置
    const minBenefitThreshold = 50; // 最小效益阈值(mm)
    const maxIterations = 10; // 最大迭代次数，防止无限循环
    let iteration = 0;
    
    while (iteration < maxIterations && !this.isTimeExceeded()) {
      iteration++;
      stats.iterations = iteration;
      
      console.log(`  📊 第${iteration}轮MW-CD分析...`);
      
      // 在该组合内查找MW余料和CD计划
      const mwRemainders = this.findMWRemainders(solution);
      const cdPlans = this.findCDPlans(solution);
      
      stats.totalMWFound = mwRemainders.length;
      stats.totalCDFound = cdPlans.length;
      
      console.log(`    - 发现MW余料: ${mwRemainders.length}个`);
      console.log(`    - 发现CD计划: ${cdPlans.length}个`);
      
      if (mwRemainders.length === 0 || cdPlans.length === 0) {
        console.log(`    - 第${iteration}轮无交换对象，收敛退出`);
            break;
          }
      
      // 分析交换机会并按效益排序
      const exchangeOpportunities = this.analyzeExchangeOpportunities(mwRemainders, cdPlans, groupKey);
      
      // 🔧 关键改进：基于效益的收敛条件
      let hasPositiveBenefitExchange = false;
      const validOpportunities = exchangeOpportunities.filter(opportunity => {
        if (opportunity.totalBenefit > minBenefitThreshold) {
          hasPositiveBenefitExchange = true;
          return true;
        }
        return false;
      });
      
      if (!hasPositiveBenefitExchange) {
        console.log(`    - 第${iteration}轮无正效益交换(阈值${minBenefitThreshold}mm)，收敛退出`);
        break;
      }
      
      console.log(`    - 发现${validOpportunities.length}个有效交换机会`);
      
      // 执行本轮最优交换
      let roundExchangeCount = 0;
      for (const opportunity of validOpportunities) {
        if (this.isTimeExceeded()) break;
        
        stats.validExchanges++;
        
        // 执行内部交换
        const exchangeResult = this.executeInternalInterchange(solution, opportunity, groupKey, remainderManager);
        
        if (exchangeResult.success) {
          stats.exchangesPerformed++;
          stats.totalBenefitGained += exchangeResult.benefit;
          roundExchangeCount++;
          console.log(`      ✅ 第${iteration}轮交换${roundExchangeCount}: 效益提升 ${exchangeResult.benefit.toFixed(2)}mm`);
          
          // 🔧 重要：每次交换后重新分析，因为状态已改变
          break;
        } else {
          console.log(`      ❌ 交换失败: ${exchangeResult.reason}`);
        }
      }
      
      if (roundExchangeCount === 0) {
        console.log(`    - 第${iteration}轮无成功交换，收敛退出`);
        break;
      }
    }
    
    if (iteration >= maxIterations) {
      console.log(`  ⚠️ 达到最大迭代次数${maxIterations}，强制收敛`);
    }
    
    console.log(`✅ ${groupKey}组合MW-CD优化完成: ${stats.iterations}轮迭代，${stats.exchangesPerformed}次交换，总效益${stats.totalBenefitGained.toFixed(2)}mm`);
    
    return stats;
  }

  /**
   * 🔧 内部交换执行方法（针对并行任务优化）
   */
  executeInternalInterchange(solution, opportunity, groupKey, remainderManager) {
    const { mw, cd, benefit } = opportunity;
    
    try {
      console.log(`      🔄 执行内部交换: MW(${mw.remainder.id}, ${mw.remainder.length}mm) ↔ CD(${cd.usedRemainders.map(r => r.id).join('+')}, ${cd.usedRemainders.reduce((sum, r) => sum + r.length, 0)}mm)`);
      
      // 1. 创建新的切割计划（用MW替代CD）
      const targetLength = cd.cuts[0]?.length || 0;
      const newWaste = mw.remainder.length - targetLength;
      
      const newCuttingPlan = new CuttingPlan({
        sourceType: 'remainder',
        sourceId: mw.remainder.id,
        sourceDescription: `${groupKey}组合MW余料 ${mw.remainder.id} (内部交换优化)`,
        sourceLength: mw.remainder.length,
        cuts: cd.cuts, // 复用原CD的切割需求
        usedRemainders: [mw.remainder],
        newRemainders: [],
        pseudoRemainders: [mw.remainder], // MW余料变为伪余料
        realRemainders: [],
        waste: newWaste
      });
      
      // 2. 移除原CD计划
      const cdIndex = solution.cuttingPlans.findIndex(plan => plan === cd);
      if (cdIndex !== -1) {
        solution.cuttingPlans.splice(cdIndex, 1);
      }
      
      // 3. 添加新计划
      solution.cuttingPlans.push(newCuttingPlan);
      
      // 4. 更新余料管理器状态
      // 标记MW余料为伪余料（已被使用）
      mw.remainder.markAsPseudo();
      
      // 从余料池中移除MW余料
      const mwIndex = remainderManager.remainderPools[groupKey].findIndex(r => r.id === mw.remainder.id);
      if (mwIndex !== -1) {
        remainderManager.remainderPools[groupKey].splice(mwIndex, 1);
      }
      
      // 恢复CD使用的余料到池中（它们现在可以重新使用）
      cd.usedRemainders.forEach(remainder => {
        if (remainder.type === REMAINDER_TYPES.PSEUDO) {
          remainder.type = REMAINDER_TYPES.PENDING; // 恢复为待定状态
          remainderManager.remainderPools[groupKey].push(remainder);
        }
      });
      
      // 重新排序余料池
      remainderManager.remainderPools[groupKey].sort((a, b) => a.length - b.length);
      
      return {
        success: true,
        benefit: benefit,
        newPlan: newCuttingPlan
      };
      
    } catch (error) {
      console.error(`      ❌ 内部交换执行失败: ${error.message}`);
      return {
        success: false,
        reason: error.message
      };
    }
  }

  /**
   * 查找MW余料（在内部优化时查找PENDING状态的余料，因为REAL状态在最后才确定）
   */
  findMWRemainders(solution) {
    const mwRemainders = [];
    
    solution.cuttingPlans.forEach(plan => {
      // 🔧 修复：在内部MW-CD优化时，余料还是PENDING状态
      if (plan.newRemainders) {
        plan.newRemainders.forEach(remainder => {
          if (remainder.type === REMAINDER_TYPES.PENDING && remainder.length >= this.constraints.wasteThreshold) {
            mwRemainders.push({
              remainder: remainder,
              plan: plan
            });
          }
        });
      }
      
      // 同时检查realRemainders（如果已经有的话）
      if (plan.realRemainders) {
        plan.realRemainders.forEach(remainder => {
          if (remainder.type === REMAINDER_TYPES.REAL) {
            mwRemainders.push({
              remainder: remainder,
              plan: plan
            });
          }
        });
      }
    });
    
    return mwRemainders;
  }

  /**
   * 查找CD计划（使用余料组合的）
   */
  findCDPlans(solution) {
    return solution.cuttingPlans.filter(plan => 
      plan.sourceType === 'remainder' && 
      plan.usedRemainders && 
      plan.usedRemainders.length >= 2
    );
  }

  /**
   * 检查是否可以执行交换
   */
  canPerformInterchange(mw, cd) {
    // 简化的交换条件检查
    const mwLength = mw.remainder.length;
    const cdTotalLength = cd.usedRemainders.reduce((sum, r) => sum + r.length, 0);
    
    return Math.abs(mwLength - cdTotalLength) < this.constraints.wasteThreshold;
  }

  /**
   * 执行交换操作
   */
  executeInterchange(solution, mw, cd, groupKey) {
    // 实现MW-CD交换逻辑
    console.log(`🔄 执行MW-CD交换: ${mw.remainder.id} ↔ ${cd.usedRemainders.map(r => r.id).join('+')}`);
    
    // 更新损耗计算
    solution.totalWaste += mw.remainder.length;
    solution.totalRealRemainder -= mw.remainder.length;
    
    // 标记交换记录
    mw.remainder.markAsWaste();
  }

  /**
   * 构建最终优化结果
   * 使用统一的StatisticsCalculator确保数据一致性
   */
  buildOptimizationResult(solutions, validation) {
    const endTime = Date.now();
    const executionTime = endTime - this.startTime;
    
    console.log('🏗️ 使用统一StatisticsCalculator构建最终优化结果...');
    
    // 🎯 关键修复：使用统一的ResultBuilder构建完整结果，传入remainderManager
    const completeResult = this.resultBuilder.buildCompleteResult(
      solutions, 
      this.designSteels, 
      this.moduleSteels,
      this.remainderManager, // 新增：传入余料管理器
      executionTime
    );
    
    // 🔧 保持向后兼容：创建原有的OptimizationResult结构
    const optimizationResult = new OptimizationResult({
      solutions: solutions,
      totalModuleUsed: completeResult.totalModuleUsed,
      totalMaterial: completeResult.totalMaterial,
      totalWaste: completeResult.totalWaste,
      totalRealRemainder: completeResult.totalRealRemainder,
      totalPseudoRemainder: completeResult.totalPseudoRemainder,
      totalLossRate: completeResult.totalLossRate,
      executionTime: executionTime,
      timestamp: new Date().toISOString(),
      version: '3.0',
      constraintValidation: validation,
      
      // 🔧 新增：完整的统计数据，前端直接使用
      completeStats: completeResult.completeStats,
      
      // 🔧 添加优化完成状态标记
      processingStatus: {
        isCompleted: true,
        remaindersFinalized: true,
        readyForRendering: true,
        completedAt: new Date().toISOString(),
        dataConsistencyChecked: completeResult.completeStats.consistencyCheck.isConsistent
      }
    });

    // 🔧 保持兼容：添加原有的附加数据
    const moduleSteelStats = this.collectModuleSteelUsageStats();
    optimizationResult.moduleSteelUsage = moduleSteelStats;
    
    const databaseRecords = this.collectDatabaseRecords();
    optimizationResult.databaseRecords = databaseRecords;
    
    console.log('✅ 统一StatisticsCalculator构建完成，数据一致性已验证');
    console.log(`📊 全局统计: 模数钢材${completeResult.totalModuleUsed}根, 损耗率${completeResult.totalLossRate}%, 执行时间${executionTime}ms`);
    
    // 🔧 数据一致性警告
    if (!completeResult.completeStats.consistencyCheck.isConsistent) {
      console.warn('⚠️ 数据一致性检查失败，请检查算法逻辑');
      completeResult.completeStats.consistencyCheck.errors.forEach(error => {
        console.warn(`   - ${error}`);
      });
    }
    
    return optimizationResult;
  }

  /**
   * V3新增：收集模数钢材使用统计（按根数）
   */
  collectModuleSteelUsageStats() {
    const stats = {};
    
    this.moduleSteelPools.forEach((pool, groupKey) => {
      const [specification, crossSection] = this.parseGroupKey(groupKey);
      
      if (!stats[specification]) {
        stats[specification] = {};
      }
      
      const poolStats = pool.getUsageStats();
      Object.entries(poolStats).forEach(([length, count]) => {
        const key = `${length}`;
        stats[specification][key] = (stats[specification][key] || 0) + count;
      });
    });
    
    return stats;
  }

  /**
   * V3新增：收集数据库存储记录
   */
  collectDatabaseRecords() {
    const records = [];
    
    this.moduleSteelPools.forEach((pool, groupKey) => {
      records.push(...pool.getDatabaseRecords());
    });
    
    return records;
  }

  /**
   * V3新增：重置所有模数钢材池
   */
  resetAllPools() {
    this.moduleSteelPools.forEach((pool) => {
      pool.reset();
    });
    this.moduleSteelPools.clear();
    console.log('🔄 所有模数钢材池已重置');
  }

  /**
   * V3新增：批量保存模数钢材使用记录到数据库（预留接口）
   */
  async saveToDatabaseAsync(records) {
    // TODO: 实现数据库集成
    console.log(`💾 准备保存${records.length}条模数钢材使用记录到数据库`);
    
    // 示例：分批保存避免数据库压力
    const batchSize = 100;
    for (let i = 0; i < records.length; i += batchSize) {
      const batch = records.slice(i, i + batchSize);
      console.log(`💾 保存批次${Math.floor(i/batchSize) + 1}: ${batch.length}条记录`);
      
      // await database.batchInsert('module_steel_usage', batch);
    }
    
    return true;
  }

  /**
   * 辅助方法 - [原版]
   */
  createDemandList(steels) {
    return steels.map(steel => ({
      id: steel.id,
      length: steel.length,
      quantity: steel.quantity,
      remaining: steel.quantity,
      crossSection: steel.crossSection
    }));
  }

  /**
   * [新增] 辅助方法 - 创建扁平化的需求列表，每个元素代表一根钢材
   */
  createFlatDemandList(steels) {
    const flatList = [];
    steels.forEach(steel => {
      for (let i = 0; i < steel.quantity; i++) {
        // 给予每个独立的部件一个唯一ID，便于追踪
        flatList.push({ ...steel, uniqueId: `${steel.id}_${i}` });
      }
    });
    return flatList;
  }

  /**
   * [新增] 辅助方法 - 将单个切割项按 designId 和 length 分组合并
   */
  groupCuts(flatCuts) {
    const grouped = new Map();
    flatCuts.forEach(cut => {
      const key = `${cut.designId}_${cut.length}`;
      if (!grouped.has(key)) {
        grouped.set(key, { ...cut, quantity: 0 });
      }
      grouped.get(key).quantity += 1;
    });
    return Array.from(grouped.values());
  }

  isTimeExceeded() {
    return (Date.now() - this.startTime) > this.constraints.timeLimit;
  }
}

/**
 * V3规格化模数钢材池 - 支持动态生成和数据库集成
 */
class SpecificationModuleSteelPool {
  constructor(specification, crossSection, availableLengths = null) {
    // 🔧 修复：使用约束配置中心的默认模数钢材长度，消除硬编码
    this.availableLengths = (availableLengths || constraintManager.getDefaultModuleLengths()).sort((a, b) => a - b); // 🔧 修复：改为升序排列
    this.specification = specification;
    this.crossSection = crossSection;
    this.usedSteels = []; // 记录实际使用的钢材（用于数据库存储）
    this.counter = 0;     // 钢材编号计数器
  }

  /**
   * 获取指定长度的模数钢材（动态生成）
   */
  getSteel(requiredLength) {
    // 🔧 优化：升序排列后，find()直接找到第一个≥需求长度的（即最短合适的）
    const bestLength = this.availableLengths.find(length => length >= requiredLength);
    
    if (!bestLength) {
      // 如果没有合适的标准长度，使用最长的
      const maxLength = Math.max(...this.availableLengths);
      console.warn(`⚠️ 需求长度${requiredLength}mm超过最大模数钢材${maxLength}mm，使用最长规格`);
      return this.createSteel(maxLength);
    }
    
    const waste = bestLength - requiredLength;
    console.log(`🎯 需求${requiredLength}mm，选择${bestLength}mm（浪费${waste}mm）`);
    return this.createSteel(bestLength);
  }

  /**
   * 创建新的模数钢材实例
   */
  createSteel(length) {
    this.counter++;
    const steel = new ModuleSteel({
      id: `${this.specification}_${this.crossSection}_M${this.counter}`,
      name: `${this.specification}-${length}mm模数钢材`,
      length: length,
      specification: this.specification,
      crossSection: this.crossSection,
      createdAt: new Date().toISOString()
    });
    
    // 记录用于数据库存储
    this.usedSteels.push({
      id: steel.id,
      specification: this.specification,
      crossSection: this.crossSection,
      length: length,
      usedAt: new Date().toISOString()
    });
    
    console.log(`🔧 动态生成模数钢材: ${steel.id} (${this.specification}, ${length}mm)`);
    
    return steel;
  }

  /**
   * 获取使用统计
   */
  getUsageStats() {
    const stats = {};
    this.usedSteels.forEach(steel => {
      const key = steel.length;
      stats[key] = (stats[key] || 0) + 1;
    });
    return stats;
  }

  /**
   * 获取数据库记录
   */
  getDatabaseRecords() {
    return [...this.usedSteels];
  }

  /**
   * 重置池状态
   */
  reset() {
    this.usedSteels = [];
    this.counter = 0;
  }
}

module.exports = SteelOptimizerV3;