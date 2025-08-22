import React, { createContext, useContext, ReactNode, useState, useEffect, useCallback } from 'react';
import { useAsyncOptimization } from '../hooks/useOptimizationResults';
import { DEFAULT_CONSTRAINTS } from '../types';

// 设计钢材接口
export interface DesignSteel {
  id: string;
  length: number;
  quantity: number;
  crossSection: number;
  displayId?: string;
  componentNumber?: string;
  specification?: string;
  partNumber?: string;
}

// 模数钢材接口
export interface ModuleSteel {
  id: string;
  name: string;
  length: number;
}

// 约束条件接口
export interface OptimizationConstraints {
  wasteThreshold: number;
  targetLossRate: number;
  timeLimit: number;
  maxWeldingSegments: number;
}

// 优化任务状态
export type OptimizationStatus = 'idle' | 'running' | 'completed' | 'error';

// 优化结果接口
export interface OptimizationResult {
  id: string;
  timestamp: number;
  designSteels?: DesignSteel[];
  moduleSteels?: ModuleSteel[];
  constraints?: OptimizationConstraints;
  results?: any;
  status: OptimizationStatus;
  progress?: number;
  error?: string;
  summary?: {
    totalLossRate?: number;
    totalModuleUsed?: number;
  }
}

// 上下文状态接口
export interface OptimizationContextType {
  // 钢材数据
  designSteels: DesignSteel[];
  moduleSteels: ModuleSteel[];
  constraints: OptimizationConstraints;
  
  // 优化结果
  currentOptimization: OptimizationResult | null;
  optimizationHistory: OptimizationResult[];
  
  // 状态 - 扩展支持异步任务
  isOptimizing: boolean;
  progress: number;
  error: string | null;
  isDataLoaded: boolean;
  
  // 新增：异步任务状态
  currentTaskId: string | null;
  taskStatus: 'idle' | 'pending' | 'running' | 'completed' | 'failed' | 'cancelled';
  taskMessage: string;
  isPolling: boolean;
  
  // 钢材管理方法
  setDesignSteels: (steels: DesignSteel[]) => void;
  addDesignSteel: (steel: DesignSteel) => void;
  updateDesignSteel: (id: string, steel: Partial<DesignSteel>) => void;
  removeDesignSteel: (id: string) => void;
  
  setModuleSteels: (steels: ModuleSteel[]) => void;
  addModuleSteel: (steel: ModuleSteel) => void;
  updateModuleSteel: (id: string, steel: Partial<ModuleSteel>) => void;
  removeModuleSteel: (id: string) => void;
  
  setConstraints: (constraints: OptimizationConstraints) => void;
  
  // 优化方法 - 更新为异步任务模式
  startOptimization: () => Promise<void>;
  cancelOptimization: () => void;
  resetTask: () => void;
  clearOptimizationData: () => void;
  clearHistory: () => void;
  getOptimizationById: (id: string) => OptimizationResult | undefined;
  
  // 新增：异步任务方法
  getTaskHistory: () => Promise<any>;
  
  // 调试方法
  debugContext: () => void;
  forceReloadData: () => void;
  }

// 默认约束条件
const defaultConstraints: OptimizationConstraints = { ...DEFAULT_CONSTRAINTS };

// LocalStorage 键名
const STORAGE_KEYS = {
  DESIGN_STEELS: 'hgj_design_steels',
  MODULE_STEELS: 'hgj_module_steels',
  CONSTRAINTS: 'hgj_constraints',
  CURRENT_OPTIMIZATION: 'hgj_current_optimization',
  OPTIMIZATION_HISTORY: 'hgj_optimization_history'
};

// 创建上下文
export const OptimizationContext = createContext<OptimizationContextType | undefined>(undefined);

// 上下文提供者组件
export const OptimizationProvider: React.FC<{ children: ReactNode }> = ({ children }) => {
  // 钢材数据状态
  const [designSteels, setDesignSteelsState] = useState<DesignSteel[]>([]);
  const [moduleSteels, setModuleSteelsState] = useState<ModuleSteel[]>([]);
  const [constraints, setConstraintsState] = useState<OptimizationConstraints>(defaultConstraints);
  
  // 优化结果状态
  const [currentOptimization, setCurrentOptimization] = useState<OptimizationResult | null>(null);
  const [optimizationHistory, setOptimizationHistory] = useState<OptimizationResult[]>([]);
  
  // 优化过程状态
  const [isOptimizing, setIsOptimizing] = useState(false);
  const [progress, setProgress] = useState(0);
  const [error, setError] = useState<string | null>(null);
  const [isDataLoaded, setIsDataLoaded] = useState(false);

  // 集成异步优化任务Hook
  const asyncOptimization = useAsyncOptimization();

  // 保存数据到localStorage（增强版，带日志、异常捕获、体积限制）
  const saveToStorage = useCallback((key: string, data: any) => {
    try {
      const dataStr = JSON.stringify(data);
      console.log(`[localStorage] 即将写入 key=${key}，长度=${dataStr.length}，内容预览=`, dataStr.slice(0, 500));
      localStorage.setItem(key, dataStr);
    } catch (error: any) {
      console.error(`[localStorage] 保存数据失败:`, error);
      if (error && error.name === 'QuotaExceededError') {
        alert('本地存储空间不足，部分功能将无法使用，请清理历史或联系开发优化！');
      }
    }
  }, []);

  // 监听异步任务状态变化，同步到Context状态
  useEffect(() => {
    const { currentTask, isActive } = asyncOptimization;
    
    console.log('🔄 异步任务状态变化:', {
      taskId: currentTask.taskId,
      status: currentTask.status,
      progress: currentTask.progress,
      hasResults: !!currentTask.results
    });
    
    // 同步优化状态
    setIsOptimizing(isActive);
    setProgress(currentTask.progress);
    
    // 同步错误状态
    if (currentTask.status === 'failed' && currentTask.error) {
      setError(currentTask.error);
    } else if (currentTask.status !== 'failed') {
      setError(null);
    }
    
    // 当任务完成时，处理结果
    if (currentTask.status === 'completed' && currentTask.results) {
      console.log('✅ 异步优化任务完成，处理结果');
      console.log('📊 结果数据:', currentTask.results);
      
      const completedOptimization: OptimizationResult = {
        id: currentTask.taskId || `opt_${Date.now()}`,
        timestamp: Date.now(),
        designSteels: [...designSteels], 
        moduleSteels: [...moduleSteels],
        constraints: { ...constraints },
        results: currentTask.results,
        status: 'completed' as OptimizationStatus,
        progress: 100,
      };
      
      console.log('💾 设置currentOptimization:', completedOptimization);
      
      setCurrentOptimization(completedOptimization);
      // 关键修复：将完整的、最新的结果也存入localStorage，确保刷新后数据一致
      saveToStorage(STORAGE_KEYS.CURRENT_OPTIMIZATION, completedOptimization);

      const historyRecord: OptimizationResult = {
        id: completedOptimization.id,
        timestamp: completedOptimization.timestamp,
        status: completedOptimization.status,
        summary: {
          totalLossRate: currentTask.results?.totalLossRate,
          totalModuleUsed: currentTask.results?.totalModuleUsed,
        }
      };
      
      setOptimizationHistory(prev => {
        const newHistory = [historyRecord, ...prev].slice(0, 10);
        saveToStorage(STORAGE_KEYS.OPTIMIZATION_HISTORY, newHistory);
        return newHistory;
      });
    }
  }, [asyncOptimization, designSteels, moduleSteels, constraints, saveToStorage]);

  // 从localStorage加载数据（增强版，带异常捕获）
  useEffect(() => {
    console.log('=== OptimizationContext 初始化加载数据 ===');
    try {
      const savedDesignSteels = localStorage.getItem(STORAGE_KEYS.DESIGN_STEELS);
      if (savedDesignSteels) setDesignSteelsState(JSON.parse(savedDesignSteels));
      
      const savedModuleSteels = localStorage.getItem(STORAGE_KEYS.MODULE_STEELS);
      if (savedModuleSteels) setModuleSteelsState(JSON.parse(savedModuleSteels));
      
      const savedConstraints = localStorage.getItem(STORAGE_KEYS.CONSTRAINTS);
      if (savedConstraints) setConstraintsState(JSON.parse(savedConstraints));
      
      const savedHistory = localStorage.getItem(STORAGE_KEYS.OPTIMIZATION_HISTORY);
      if (savedHistory) setOptimizationHistory(JSON.parse(savedHistory).slice(0, 10));

    } catch (error) {
      console.error('加载基础数据失败:', error);
    }
    setIsDataLoaded(true);
    console.log('=== 数据加载完成 ===');
  }, []); 

  // 这个useEffect独立出来，专门处理currentOptimization的加载，并响应isDataLoaded
  useEffect(() => {
    if (isDataLoaded) {
      if (!asyncOptimization.isActive && asyncOptimization.currentTask.status === 'idle') {
        const savedCurrentOptimization = localStorage.getItem(STORAGE_KEYS.CURRENT_OPTIMIZATION);
        if (savedCurrentOptimization) {
          try {
            const parsedOptimization = JSON.parse(savedCurrentOptimization);

            // 关键修复：只恢复未完成的任务，防止在新会话开始时直接跳转到旧的结果页
            if (parsedOptimization && parsedOptimization.status !== 'completed') {
              setCurrentOptimization(parsedOptimization);
              console.log('恢复了未完成的优化任务:', parsedOptimization.id, parsedOptimization.status);
            } else {
              console.log('忽略了已完成的旧优化结果，确保应用从主页开始。');
              // 可选：在这里清除localStorage中的旧结果，以保持清洁
              // localStorage.removeItem(STORAGE_KEYS.CURRENT_OPTIMIZATION);
            }
          } catch(e) {
            console.error("解析currentOptimization失败", e)
          }
        }
      } else {
        console.log('🚫 跳过加载localStorage中的currentOptimization，因为有活跃的异步任务');
      }
    }
  }, [isDataLoaded, asyncOptimization.isActive, asyncOptimization.currentTask.status]);

  // 强制重新加载数据
  const forceReloadData = useCallback(() => {
    console.log('=== 强制重新加载数据 ===');
    setIsDataLoaded(false);
    
    try {
      const savedDesignSteels = localStorage.getItem(STORAGE_KEYS.DESIGN_STEELS);
      if (savedDesignSteels) {
        const parsedDesignSteels = JSON.parse(savedDesignSteels);
        setDesignSteelsState(parsedDesignSteels);
        console.log('重新加载设计钢材:', parsedDesignSteels.length, '条');
      }

      const savedModuleSteels = localStorage.getItem(STORAGE_KEYS.MODULE_STEELS);
      if (savedModuleSteels) {
        const parsedModuleSteels = JSON.parse(savedModuleSteels);
        setModuleSteelsState(parsedModuleSteels);
        console.log('重新加载模数钢材:', parsedModuleSteels.length, '条');
      }

      const savedConstraints = localStorage.getItem(STORAGE_KEYS.CONSTRAINTS);
      if (savedConstraints) {
        const parsedConstraints = JSON.parse(savedConstraints);
        setConstraintsState(parsedConstraints);
        console.log('重新加载约束条件:', parsedConstraints);
      }

      // 只在没有活跃的异步任务时才重新加载currentOptimization
      if (!asyncOptimization.isActive && asyncOptimization.currentTask.status === 'idle') {
        const savedCurrentOptimization = localStorage.getItem(STORAGE_KEYS.CURRENT_OPTIMIZATION);
        if (savedCurrentOptimization) {
          const parsedOptimization = JSON.parse(savedCurrentOptimization);
          setCurrentOptimization(parsedOptimization);
          console.log('重新加载当前优化结果:', parsedOptimization.id, parsedOptimization.status);
        }
      } else {
        console.log('🚫 跳过重新加载currentOptimization，因为有活跃的异步任务');
      }

      const savedHistory = localStorage.getItem(STORAGE_KEYS.OPTIMIZATION_HISTORY);
      if (savedHistory) {
        const parsedHistory = JSON.parse(savedHistory);
        setOptimizationHistory(parsedHistory);
        console.log('重新加载优化历史:', parsedHistory.length, '条');
      }
      
      setIsDataLoaded(true);
      console.log('=== 数据重新加载完成 ===');
    } catch (error) {
      console.error('重新加载数据失败:', error);
      setIsDataLoaded(true);
    }
  }, [asyncOptimization.isActive, asyncOptimization.currentTask.status]);

  // 调试Context状态
  const debugContext = useCallback(() => {
    console.log('=== OptimizationContext 状态调试 ===');
    console.log('isDataLoaded:', isDataLoaded);
    console.log('designSteels:', designSteels);
    console.log('moduleSteels:', moduleSteels);
    console.log('constraints:', constraints);
    console.log('currentOptimization:', currentOptimization);
    console.log('optimizationHistory:', optimizationHistory);
    console.log('isOptimizing:', isOptimizing);
    console.log('progress:', progress);
    console.log('error:', error);
    
    // 检查localStorage
    console.log('=== LocalStorage 状态 ===');
    Object.entries(STORAGE_KEYS).forEach(([key, storageKey]) => {
      const data = localStorage.getItem(storageKey);
      console.log(`${key}:`, data ? JSON.parse(data) : '无数据');
    });
    console.log('=== 调试完成 ===');
  }, [isDataLoaded, designSteels, moduleSteels, constraints, currentOptimization, optimizationHistory, isOptimizing, progress, error]);

  // 钢材管理方法
  const setDesignSteels = useCallback((steels: DesignSteel[]) => {
    setDesignSteelsState(steels);
    saveToStorage(STORAGE_KEYS.DESIGN_STEELS, steels);
  }, [saveToStorage]);

  const addDesignSteel = useCallback((steel: DesignSteel) => {
    const newSteels = [...designSteels, steel];
    setDesignSteels(newSteels);
  }, [designSteels, setDesignSteels]);

  const updateDesignSteel = useCallback((id: string, updates: Partial<DesignSteel>) => {
    const newSteels = designSteels.map(steel => 
      steel.id === id ? { ...steel, ...updates } : steel
    );
    setDesignSteels(newSteels);
  }, [designSteels, setDesignSteels]);

  const removeDesignSteel = useCallback((id: string) => {
    const newSteels = designSteels.filter(steel => steel.id !== id);
    setDesignSteels(newSteels);
  }, [designSteels, setDesignSteels]);

  const setModuleSteels = useCallback((steels: ModuleSteel[]) => {
    setModuleSteelsState(steels);
    saveToStorage(STORAGE_KEYS.MODULE_STEELS, steels);
  }, [saveToStorage]);

  const addModuleSteel = useCallback((steel: ModuleSteel) => {
    const newSteels = [...moduleSteels, steel];
    setModuleSteels(newSteels);
  }, [moduleSteels, setModuleSteels]);

  const updateModuleSteel = useCallback((id: string, updates: Partial<ModuleSteel>) => {
    const newSteels = moduleSteels.map(steel => 
      steel.id === id ? { ...steel, ...updates } : steel
    );
    setModuleSteels(newSteels);
  }, [moduleSteels, setModuleSteels]);

  const removeModuleSteel = useCallback((id: string) => {
    const newSteels = moduleSteels.filter(steel => steel.id !== id);
    setModuleSteels(newSteels);
  }, [moduleSteels, setModuleSteels]);

  const setConstraints = useCallback((newConstraints: OptimizationConstraints) => {
    setConstraintsState(newConstraints);
    saveToStorage(STORAGE_KEYS.CONSTRAINTS, newConstraints);
  }, [saveToStorage]);

  // 优化方法 - 更新为异步任务模式
  const startOptimization = useCallback(async () => {
    if (designSteels.length === 0 || moduleSteels.length === 0) {
      setError('请先添加设计钢材和模数钢材');
      return;
    }

    try {
      // 关键修复：在提交新任务之前，立即清除上一次的优化结果。
      // 这是整个流程的起点，确保状态纯净。
      setCurrentOptimization(null);
      saveToStorage(STORAGE_KEYS.CURRENT_OPTIMIZATION, null);
      
      console.log('🚀 启动异步优化任务');
      setError(null);
      
      const requestData = {
        designSteels: designSteels,
        moduleSteels: moduleSteels,
        constraints: {
          ...constraints,
          // 将用户输入的焊接次数转换为后端需要的焊接段数（次数+1=段数）
          maxWeldingSegments: constraints.maxWeldingSegments + 1,
          timeLimit: Math.round(constraints.timeLimit * 1000)
        }
      };
      
      // 使用异步优化Hook提交任务
      const result = await asyncOptimization.submitOptimization(requestData);
      
      if (!result.success) {
        throw new Error(result.error || '提交优化任务失败');
      }
      
      console.log('✅ 异步优化任务已提交，TaskID:', result.taskId);
      
    } catch (err: any) {
      console.error('❌ 启动异步优化失败:', err);
      const errorMessage = err.message || '启动优化过程中发生错误';
      setError(errorMessage);
    }
  }, [designSteels, moduleSteels, constraints, asyncOptimization, saveToStorage]);

  const cancelOptimization = useCallback(async () => {
    console.log('🛑 取消优化任务');
    
    // 如果有活跃的异步任务，取消它
    if (asyncOptimization.isActive) {
      const result = await asyncOptimization.cancelTask();
      if (result.success) {
        console.log('✅ 异步任务已取消');
      } else {
        console.error('❌ 取消异步任务失败:', result.error);
      }
    }
    
    // 重置本地状态
    setIsOptimizing(false);
    setProgress(0);
    setError('优化已取消');
  }, [asyncOptimization]);

  const clearOptimizationData = useCallback(() => {
    setCurrentOptimization(null);
    localStorage.removeItem(STORAGE_KEYS.CURRENT_OPTIMIZATION);
  }, []);

  const clearHistory = useCallback(() => {
    setOptimizationHistory([]);
    localStorage.removeItem(STORAGE_KEYS.OPTIMIZATION_HISTORY);
  }, []);

  const getOptimizationById = useCallback((id: string) => {
    return optimizationHistory.find(opt => opt.id === id);
  }, [optimizationHistory]);

  const contextValue: OptimizationContextType = {
    // 钢材数据
    designSteels,
    moduleSteels,
    constraints,
    
    // 优化结果
    currentOptimization,
    optimizationHistory,
    
    // 状态
    isOptimizing,
    progress,
    error,
    isDataLoaded,
    
    // 新增：异步任务状态
    currentTaskId: asyncOptimization.currentTask.taskId,
    taskStatus: asyncOptimization.currentTask.status,
    taskMessage: asyncOptimization.currentTask.message,
    isPolling: asyncOptimization.isPolling,
    
    // 钢材管理方法
    setDesignSteels,
    addDesignSteel,
    updateDesignSteel,
    removeDesignSteel,
    
    setModuleSteels,
    addModuleSteel,
    updateModuleSteel,
    removeModuleSteel,
    
    setConstraints,
    
    // 优化方法
    startOptimization,
    cancelOptimization,
    resetTask: asyncOptimization.resetTask,
    clearOptimizationData,
    clearHistory,
    getOptimizationById,
    
    // 新增：异步任务方法
    getTaskHistory: asyncOptimization.getTaskHistory,
    
    // 调试方法
    debugContext,
    forceReloadData,
  };

  return (
    <OptimizationContext.Provider value={contextValue}>
      {children}
    </OptimizationContext.Provider>
  );
};

// 使用上下文的hook
export const useOptimizationContext = () => {
  const context = useContext(OptimizationContext);
  if (context === undefined) {
    throw new Error('useOptimizationContext must be used within an OptimizationProvider');
  }
  return context;
}; 