import { useMemo, useCallback, useEffect, useState, useRef } from 'react';
import { DesignSteel, OptimizationResult } from '../types';
import { generateDisplayIds, regroupOptimizationResultsBySpecification } from '../utils/steelUtils';

// 统计数据接口
export interface TotalStats {
  totalModuleCount: number;
  totalModuleLength: number;
  totalWaste: number;
  totalRealRemainder: number;
  totalPseudoRemainder: number;
  overallLossRate: number;
}

// 图表数据接口
export interface ChartDataPoint {
  specification: string;
  groupKey: string;
  lossRate: number;
  moduleUsed: number;
  waste: number;
  realRemainder: number;
  pseudoRemainder: number;
}

export interface PieDataPoint {
  name: string;
  value: number;
  fill: string;
}

export interface ChartData {
  lossRateData: ChartDataPoint[];
  pieData: PieDataPoint[];
}

// 需求验证数据接口
export interface RequirementValidationItem {
  key: string;
  id: string;
  specification: string;
  crossSection: number;
  length: number;
  quantity: number;
  produced: number;
  satisfied: boolean;
  difference: number;
  groupKey?: string;
}

// 模数钢材统计接口
export interface ModuleUsageItem {
  key: string;
  specification: string;
  length: number | string;
  count: number;
  totalLength: number;
  isSubtotal?: boolean;
  isTotal?: boolean;
}

// 处理后的优化结果接口
export interface ProcessedOptimizationResults {
  totalStats: TotalStats;
  chartData: ChartData;
  requirementValidation: RequirementValidationItem[];
  moduleUsageStats: {
    sortedStats: ModuleUsageItem[];
    groupKeyTotals: Record<string, { count: number; totalLength: number }>;
    grandTotal: { count: number; totalLength: number };
  };
  regroupedResults: Record<string, any>;
  designIdToDisplayIdMap: Map<string, string>;
  isAllRequirementsSatisfied: boolean;
  hasDataError: boolean;
  errorMessage?: string;
}

/**
 * 统一的优化结果数据处理Hook
 * 解决错误引用值问题，确保所有数据都来自同一个权威源
 */
export const useOptimizationResults = (
  results: OptimizationResult | null,
  designSteels: DesignSteel[],
  moduleSteels: any[]
): ProcessedOptimizationResults => {
  // 生成显示ID映射
  const designSteelsWithIds = useMemo(() => {
    return generateDisplayIds(designSteels || []);
  }, [designSteels]);

  const designIdToDisplayIdMap = useMemo(() => {
    const map = new Map<string, string>();
    designSteelsWithIds.forEach(steel => {
      map.set(steel.id, steel.displayId || steel.id);
    });
    return map;
  }, [designSteelsWithIds]);

  // 🔧 修复：直接使用后端ResultBuilder计算的完整统计数据
  const totalStats = useMemo((): TotalStats => {
    if (!results) {
      return {
        totalModuleCount: 0,
        totalModuleLength: 0,
        totalWaste: 0,
        totalRealRemainder: 0,
        totalPseudoRemainder: 0,
        overallLossRate: 0,
      };
    }

    // ✅ 关键修复：优先使用completeStats中的预计算数据
    if (results.completeStats?.global) {
      const globalStats = results.completeStats.global;
      return {
        totalModuleCount: globalStats.totalModuleCount,
        totalModuleLength: globalStats.totalModuleLength,
        totalWaste: globalStats.totalWaste,
        totalRealRemainder: globalStats.totalRealRemainder,
        totalPseudoRemainder: globalStats.totalPseudoRemainder,
        overallLossRate: globalStats.overallLossRate,
      };
    }

    // 🔄 兼容性保证：如果新数据结构不存在，使用原有字段
    console.warn('⚠️ 使用兼容模式：completeStats不存在，使用原有字段');
    return {
      totalModuleCount: results.totalModuleUsed || 0,
      totalModuleLength: results.totalMaterial || 0,
      totalWaste: results.totalWaste || 0,
      totalRealRemainder: results.totalRealRemainder || 0,
      totalPseudoRemainder: results.totalPseudoRemainder || 0,
      overallLossRate: results.totalLossRate || 0,
    };
  }, [results]);

  // 🔧 修复：图表数据直接使用后端预计算结果
  const chartData = useMemo((): ChartData => {
    if (!results) {
      return { lossRateData: [], pieData: [] };
    }

    // ✅ 优先使用completeStats中的预计算图表数据
    if (results.completeStats?.chartData) {
      return results.completeStats.chartData;
    }

    // 🔄 兼容性保证：如果新数据结构不存在，保持原有逻辑
    console.warn('⚠️ 图表数据兼容模式：completeStats.chartData不存在，使用原有计算逻辑');
    
    if (!results.solutions) {
      return { lossRateData: [], pieData: [] };
    }

    const groupKeys = Object.keys(results.solutions);
    const colors = ['#0088FE', '#00C49F', '#FFBB28', '#FF8042', '#8884D8'];

    const lossRateData = groupKeys.map(groupKey => {
      const solution = results.solutions[groupKey];
      const [specification, crossSectionStr] = groupKey.split('_');
      const displayName = `${specification}(${crossSectionStr}mm²)`;
      
      const lossRate = (solution.totalMaterial || 0) > 0 
        ? ((solution.totalWaste + solution.totalRealRemainder) / (solution.totalMaterial || 1)) * 100 
        : 0;
      
      return {
        specification: displayName,
        groupKey: groupKey,
        lossRate: parseFloat(lossRate.toFixed(2)),
        moduleUsed: solution.totalModuleUsed || 0,
        waste: solution.totalWaste || 0,
        realRemainder: solution.totalRealRemainder || 0,
        pseudoRemainder: solution.totalPseudoRemainder || 0
      };
    });

    const pieData = groupKeys.map((groupKey, index) => {
      const [specification, crossSectionStr] = groupKey.split('_');
      const displayName = `${specification}(${crossSectionStr}mm²)`;
      
      return {
        name: displayName,
        value: results.solutions[groupKey].totalModuleUsed || 0,
        fill: colors[index % colors.length]
      };
    });

    return { lossRateData, pieData };
  }, [results]);

  // 🔧 修复：需求验证数据直接使用后端预计算结果
  const requirementValidation = useMemo((): RequirementValidationItem[] => {
    if (!results) return [];

    // ✅ 优先使用completeStats中的预计算需求验证数据
    if (results.completeStats?.requirementValidation?.items) {
      return results.completeStats.requirementValidation.items;
    }

    // 🔄 兼容性保证：如果新数据结构不存在，保持原有逻辑
    console.warn('⚠️ 需求验证兼容模式：completeStats.requirementValidation不存在，使用原有计算逻辑');
    
    if (!results.solutions) return [];

    return designSteelsWithIds.map(steel => {
      // 计算该设计钢材的实际生产数量
      let produced = 0;
      Object.values(results.solutions).forEach((solution: any) => {
        solution.cuttingPlans?.forEach((plan: any) => {
          plan.cuts?.forEach((cut: any) => {
            if (cut.designId === steel.id) {
              produced += cut.quantity;
            }
          });
        });
      });

      return {
        key: steel.id,
        id: steel.displayId || steel.id,
        specification: steel.specification || '未知规格',
        crossSection: steel.crossSection,
        length: steel.length,
        quantity: steel.quantity,
        produced: produced,
        satisfied: produced >= steel.quantity,
        difference: produced - steel.quantity,
        groupKey: steel.groupKey
      };
    });
  }, [results, designSteelsWithIds]);

  // 🔧 修复：模数钢材使用统计直接使用后端预计算结果
  const moduleUsageStats = useMemo(() => {
    if (!results) {
      return { 
        sortedStats: [], 
        groupKeyTotals: {}, 
        grandTotal: { count: 0, totalLength: 0 } 
      };
    }

    // ✅ 优先使用completeStats中的预计算模数钢材统计
    if (results.completeStats?.moduleUsageStats) {
      const statsData = results.completeStats.moduleUsageStats;
      
      // 转换后端数据格式为前端需要的格式
      const sortedStats: ModuleUsageItem[] = [];
      const groupKeyTotals: Record<string, { count: number; totalLength: number }> = {};
      
      Object.entries(statsData.bySpecification).forEach(([groupKey, specData]: [string, any]) => {
        const [specification] = groupKey.split('_');
        
        Object.entries(specData.moduleBreakdown).forEach(([lengthStr, moduleData]: [string, any]) => {
          const length = parseInt(lengthStr);
          sortedStats.push({
            key: `${groupKey}_${length}`,
            specification: specification,
            length: length,
            count: moduleData.count,
            totalLength: moduleData.totalLength
          });
        });
        
        groupKeyTotals[groupKey] = specData.groupTotal;
      });
      
      // 排序
      sortedStats.sort((a, b) => {
        if (a.specification !== b.specification) {
          return a.specification.localeCompare(b.specification);
        }
        const lengthA = typeof a.length === 'number' ? a.length : parseInt(String(a.length), 10) || 0;
        const lengthB = typeof b.length === 'number' ? b.length : parseInt(String(b.length), 10) || 0;
        return lengthA - lengthB;
      });
      
      return { 
        sortedStats, 
        groupKeyTotals, 
        grandTotal: statsData.grandTotal 
      };
    }

    // 🔄 兼容性保证：如果新数据结构不存在，保持原有逻辑
    console.warn('⚠️ 模数钢材统计兼容模式：completeStats.moduleUsageStats不存在，使用原有计算逻辑');
    
    if (!results.solutions) {
      return { 
        sortedStats: [], 
        groupKeyTotals: {}, 
        grandTotal: { count: 0, totalLength: 0 } 
      };
    }

    // 原有的复杂计算逻辑保持不变作为兼容性保证
    const moduleUsageStatsMap: Record<string, {
      groupKey: string;
      specification: string;
      crossSection: number;
      length: number;
      count: number;
      totalLength: number;
    }> = {};

    Object.entries(results.solutions).forEach(([groupKey, solution]) => {
      const [specification, crossSectionStr] = groupKey.split('_');
      const crossSection = parseInt(crossSectionStr);
      
      const uniqueModuleBars: Record<string, { length: number; sourceId: string }> = {};
      
      (solution as any).details?.forEach((detail: any) => {
        if (detail.sourceType === 'module' && detail.sourceId) {
          const length = detail.moduleLength || detail.sourceLength;
          const sourceId = detail.sourceId;
          
          if (!uniqueModuleBars[sourceId]) {
            uniqueModuleBars[sourceId] = {
              length: length,
              sourceId: sourceId
            };
          }
        }
      });
      
      const moduleBarCounts: Record<number, number> = {};
      Object.values(uniqueModuleBars).forEach(bar => {
        if (!moduleBarCounts[bar.length]) {
          moduleBarCounts[bar.length] = 0;
        }
        moduleBarCounts[bar.length] += 1;
      });
      
      Object.entries(moduleBarCounts).forEach(([lengthStr, count]) => {
        const length = parseInt(lengthStr);
        const key = `${groupKey}_${length}`;
        if (!moduleUsageStatsMap[key]) {
          moduleUsageStatsMap[key] = {
            groupKey: groupKey,
            specification: specification,
            crossSection: crossSection,
            length: length,
            count: 0,
            totalLength: 0
          };
        }
        moduleUsageStatsMap[key].count += count;
        moduleUsageStatsMap[key].totalLength += length * count;
      });
    });

    const sortedStats = Object.values(moduleUsageStatsMap).map(stat => ({
      ...stat,
      key: `${stat.groupKey}_${stat.length}`
    })).sort((a, b) => {
      if (a.specification !== b.specification) {
        return a.specification.localeCompare(b.specification);
      }
      if (a.crossSection !== b.crossSection) {
        return a.crossSection - b.crossSection;
      }
      return a.length - b.length;
    });

    const groupKeyTotals: Record<string, { count: number; totalLength: number }> = {};
    sortedStats.forEach(stat => {
      if (!groupKeyTotals[stat.groupKey]) {
        groupKeyTotals[stat.groupKey] = { count: 0, totalLength: 0 };
      }
      groupKeyTotals[stat.groupKey].count += stat.count;
      groupKeyTotals[stat.groupKey].totalLength += stat.totalLength;
    });

    const grandTotal = {
      count: sortedStats.reduce((sum, stat) => sum + stat.count, 0),
      totalLength: sortedStats.reduce((sum, stat) => sum + stat.totalLength, 0)
    };

    return { sortedStats, groupKeyTotals, grandTotal };
  }, [results]);

  // 重新分组的结果数据
  const regroupedResults = useMemo(() => {
    if (!results?.solutions) return {};
    const regrouped = regroupOptimizationResultsBySpecification(results.solutions);
    
    // 注入displayId到切割计划中
    Object.values(regrouped).forEach((solution: any) => {
      solution.cuttingPlans?.forEach((plan: any) => {
        plan.cuts?.forEach((cut: any) => {
          cut.displayId = designIdToDisplayIdMap.get(cut.designId) || cut.designId;
        });
      });
    });
    return regrouped;
  }, [results, designIdToDisplayIdMap]);

  // 🔧 修复：检查所有需求是否满足，优先使用后端预计算结果
  const isAllRequirementsSatisfied = useMemo(() => {
    // ✅ 优先使用completeStats中的预计算结果
    if (results?.completeStats?.requirementValidation?.summary?.allSatisfied !== undefined) {
      return results.completeStats.requirementValidation.summary.allSatisfied;
    }

    // 🔄 兼容性保证：使用前端计算结果
    return requirementValidation.every(item => item.satisfied);
  }, [results, requirementValidation]);

  // 🔧 修复：错误检查，优先使用后端数据一致性检查结果
  const { hasDataError, errorMessage } = useMemo(() => {
    if (!results) {
      return { hasDataError: true, errorMessage: '暂无优化结果数据' };
    }

    // ✅ 优先检查后端数据一致性验证结果
    if (results.completeStats?.consistencyCheck) {
      const consistencyCheck = results.completeStats.consistencyCheck;
      if (!consistencyCheck.isConsistent) {
        return { 
          hasDataError: true, 
          errorMessage: `数据一致性检查失败: ${consistencyCheck.errors.join('; ')}` 
        };
      }
    }

    if (!results.solutions || Object.keys(results.solutions).length === 0) {
      return { hasDataError: true, errorMessage: '优化结果为空，请检查输入数据' };
    }

    // 检查数据合理性
    if (totalStats.totalModuleCount === 0 && totalStats.totalModuleLength === 0) {
      return { hasDataError: true, errorMessage: '模数钢材用量为0，可能存在数据异常' };
    }

    return { hasDataError: false };
  }, [results, totalStats]);

  return {
    totalStats,
    chartData,
    requirementValidation,
    moduleUsageStats,
    regroupedResults,
    designIdToDisplayIdMap,
    isAllRequirementsSatisfied,
    hasDataError,
    errorMessage
  };
}; 

/**
 * 异步优化任务处理Hook
 * 实现提交任务、轮询状态、获取结果的完整流程
 */
export const useAsyncOptimization = () => {
  const [currentTask, setCurrentTask] = useState<{
    taskId: string | null;
    status: 'idle' | 'pending' | 'running' | 'completed' | 'failed' | 'cancelled';
    progress: number;
    message: string;
    results: OptimizationResult | null;
    error: string | null;
    executionTime: number;
  }>({
    taskId: null,
    status: 'idle',
    progress: 0,
    message: '',
    results: null,
    error: null,
    executionTime: 0
  });

  const [isPolling, setIsPolling] = useState(false);
  const pollingRef = useRef<NodeJS.Timeout | null>(null);

  // 停止轮询
  const stopPolling = useCallback(() => {
    if (pollingRef.current) {
      clearInterval(pollingRef.current);
      pollingRef.current = null;
    }
    setIsPolling(false);
  }, []);

  // 开始轮询任务状态
  const startPolling = useCallback((taskId: string) => {
    if (!taskId || taskId === 'undefined') {
      console.error('无效的taskId，轮询已中止:', taskId);
      return;
    }

    if (pollingRef.current) {
      clearInterval(pollingRef.current);
    }

    setIsPolling(true);
    
    const pollTaskStatus = async () => {
      try {
        const response = await fetch(`/api/task/${taskId}`);
        const result = await response.json();

        if (result.success) {
          setCurrentTask(prev => ({
            ...prev,
            status: result.status,
            progress: result.progress || 0,
            message: result.message || '',
            executionTime: result.executionTime || 0,
            results: result.results || null,
            error: result.error || null
          }));

          // 如果任务完成、失败或取消，停止轮询
          if (['completed', 'failed', 'cancelled'].includes(result.status)) {
            stopPolling();
            console.log(`📋 任务${result.status}:`, taskId);
          }
        } else {
          console.error('❌ 轮询任务状态失败:', result.error);
          // 继续轮询，可能是临时网络问题
        }
      } catch (error) {
        console.error('❌ 轮询请求异常:', error);
        // 继续轮询，可能是临时网络问题
      }
    };

    // 立即执行一次
    pollTaskStatus();
    
    // 每2秒轮询一次
    pollingRef.current = setInterval(pollTaskStatus, 2000);
  }, [stopPolling]);

  // 提交优化任务
  const submitOptimization = useCallback(async (optimizationData: any) => {
    try {
      console.log('🚀 提交优化任务...');
      
      const response = await fetch('/api/optimize', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(optimizationData)
      });

      const result = await response.json();

      if (result.success && result.taskId) {
        console.log('✅ 任务创建成功:', result.taskId);
        
        setCurrentTask({
          taskId: result.taskId,
          status: 'pending',
          progress: 0,
          message: result.message || '任务已创建',
          results: null,
          error: null,
          executionTime: 0
        });

        // 开始轮询任务状态
        startPolling(result.taskId);
        
        return { success: true, taskId: result.taskId };
      } else {
        const errorMsg = result.error || '任务创建失败，但未返回任务ID';
        console.error('❌ 任务创建失败:', errorMsg);
        setCurrentTask(prev => ({
          ...prev,
          status: 'failed',
          error: errorMsg
        }));
        return { success: false, error: errorMsg };
      }
    } catch (error) {
      console.error('❌ 提交优化任务异常:', error);
      const errorMessage = error instanceof Error ? error.message : '网络错误';
      setCurrentTask(prev => ({
        ...prev,
        status: 'failed',
        error: errorMessage
      }));
      return { success: false, error: errorMessage };
    }
  }, [startPolling]);

  // 取消任务
  const cancelTask = useCallback(async () => {
    if (!currentTask.taskId) {
      return { success: false, error: '没有活跃的任务' };
    }

    try {
      const response = await fetch(`/api/task/${currentTask.taskId}`, {
        method: 'DELETE'
      });

      const result = await response.json();

      if (result.success) {
        console.log('✅ 任务已取消');
        stopPolling();
        setCurrentTask(prev => ({
          ...prev,
          status: 'cancelled',
          message: '任务已取消'
        }));
        return { success: true };
      } else {
        console.error('❌ 取消任务失败:', result.error);
        return { success: false, error: result.error };
      }
    } catch (error) {
      console.error('❌ 取消任务异常:', error);
      const errorMessage = error instanceof Error ? error.message : '网络错误';
      return { success: false, error: errorMessage };
    }
  }, [currentTask.taskId, stopPolling]);

  // 重置任务状态
  const resetTask = useCallback(() => {
    stopPolling();
    setCurrentTask({
      taskId: null,
      status: 'idle',
      progress: 0,
      message: '',
      results: null,
      error: null,
      executionTime: 0
    });
  }, [stopPolling]);

  // 获取任务历史
  const getTaskHistory = useCallback(async (limit = 20) => {
    try {
      const response = await fetch(`/api/tasks?limit=${limit}`);
      const result = await response.json();

      if (result.success) {
        return { success: true, tasks: result.tasks, total: result.total };
      } else {
        console.error('❌ 获取任务历史失败:', result.error);
        return { success: false, error: result.error };
      }
    } catch (error) {
      console.error('❌ 获取任务历史异常:', error);
      const errorMessage = error instanceof Error ? error.message : '网络错误';
      return { success: false, error: errorMessage };
    }
  }, []);

  // 组件卸载时清理轮询
  useEffect(() => {
    return () => {
      stopPolling();
    };
  }, [stopPolling]);

  // 确保返回的对象引用是稳定的，只有在内部状态变化时才创建新对象
  return useMemo(() => ({
    // 任务状态
    currentTask,
    isPolling,
    
    // 任务操作
    submitOptimization,
    cancelTask,
    resetTask,
    
    // 工具方法
    getTaskHistory,
    
    // 便捷状态判断
    isIdle: currentTask.status === 'idle',
    isPending: currentTask.status === 'pending',
    isRunning: currentTask.status === 'running',
    isCompleted: currentTask.status === 'completed',
    isFailed: currentTask.status === 'failed',
    isCancelled: currentTask.status === 'cancelled',
    isActive: ['pending', 'running'].includes(currentTask.status),
    hasResults: currentTask.status === 'completed' && !!currentTask.results,
    hasError: currentTask.status === 'failed' && !!currentTask.error
  }), [currentTask, isPolling, submitOptimization, cancelTask, resetTask, getTaskHistory]);
}; 