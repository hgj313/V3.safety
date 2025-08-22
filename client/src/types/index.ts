/**
 * 钢材采购优化系统 V3.0 - 前端类型定义
 * 对应后端API的TypeScript接口
 */

// ==================== 基础数据类型 ====================

export interface DesignSteel {
  id: string;
  length: number;
  quantity: number;
  crossSection: number;
  displayId?: string;
  componentNumber?: string;
  specification?: string;
  partNumber?: string;
  groupKey?: string; // 规格+截面面积组合键，用于调试和分组
}

export interface ModuleSteel {
  id: string;
  name: string;
  length: number;
}

export interface RemainderV3 {
  id: string;
  length: number;
  type: 'pending' | 'pseudo' | 'real' | 'waste';
  isConsumed: boolean;
  sourceChain: string[];
  crossSection: number;
  createdAt: string;
  consumedAt?: string;
  originalLength?: number;
  parentId?: string;
}

// ==================== 约束系统 ====================

export interface OptimizationConstraints {
  wasteThreshold: number;        // 废料阈值 (mm) - 余料长度小于此值时视为废料
  targetLossRate: number;        // 目标损耗率 (%) - 优化目标参考值（软约束）
  timeLimit: number;             // 计算时间限制 (ms) - 算法最大运行时间
  maxWeldingSegments: number;    // 最大焊接次数 (次) - 单根设计钢材允许的最大焊接次数
}

export interface ConstraintViolation {
  type: string;
  message: string;
  current?: any;
  suggested?: any;
  steelIndex?: number;
  details?: any;
}

export interface ConstraintSuggestion {
  type: 'addLongerModule' | 'increaseWelding';
  priority: number;
  title: string;
  description: string;
  details: any;
  implementation: {
    action: string;
    [key: string]: any;
  };
}

export interface ConstraintValidation {
  isValid: boolean;
  violations: ConstraintViolation[];
  suggestions: ConstraintSuggestion[];
  warnings: any[];
}

// ==================== 切割和优化结果 ====================

export interface CuttingDetail {
  sourceType: 'module' | 'remainder';
  sourceId: string;
  sourceLength: number;
  moduleType?: string;
  moduleLength?: number;
  designId: string;
  length: number;
  quantity: number;
  remainderInfo?: any;
  weldingCount: number;
}

export interface CuttingPlan {
  sourceType: 'module' | 'remainder';
  sourceDescription: string;
  sourceLength: number;
  moduleType?: string;
  moduleLength?: number;
  cuts: Array<{
    designId: string;
    length: number;
    quantity: number;
  }>;
  newRemainders: RemainderV3[];
  pseudoRemainders: RemainderV3[];
  realRemainders: RemainderV3[];
  waste: number;
  usedRemainders: RemainderV3[];
}

export interface OptimizationSolution {
  cuttingPlans: CuttingPlan[];
  totalModuleUsed: number;
  totalMaterial?: number; // 新增：总材料长度
  totalWaste: number;
  totalPseudoRemainder: number;
  totalRealRemainder: number;
  details: CuttingDetail[];
  lossRateBreakdown: Record<string, any>;
}

export interface LossRateValidation {
  isValid: boolean;
  totalLossRate: number;
  weightedAverage: number;
  difference: number;
  specResults: Array<{
    lossRate: number;
    weight: number;
    contribution: number;
  }>;
  errorMessage?: string;
}

// 新增：完整统计数据类型定义
export interface CompleteStatsData {
  global: {
    totalModuleCount: number;
    totalModuleLength: number;
    totalWaste: number;
    totalRealRemainder: number;
    totalPseudoRemainder: number;
    totalDesignSteelLength: number;
    overallLossRate: number;
    materialUtilizationRate: number;
  };
  chartData: {
    lossRateData: Array<{
      specification: string;
      groupKey: string;
      lossRate: number;
      moduleUsed: number;
      waste: number;
      realRemainder: number;
      pseudoRemainder: number;
      utilization: number;
    }>;
    pieData: Array<{
      name: string;
      value: number;
      fill: string;
    }>;
  };
  requirementValidation: {
    items: Array<{
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
      satisfactionRate: number;
    }>;
    summary: {
      total: number;
      satisfied: number;
      unsatisfied: number;
      allSatisfied: boolean;
      overallSatisfactionRate: number;
    };
  };
  moduleUsageStats: {
    bySpecification: Record<string, any>;
    grandTotal: { count: number; totalLength: number };
  };
  specificationDetails: Record<string, any>;
  consistencyCheck: {
    isConsistent: boolean;
    errors: string[];
    warnings: string[];
    checkTime: string;
  };
}

export interface OptimizationResult {
  solutions: Record<string, OptimizationSolution>;
  totalLossRate: number;
  totalModuleUsed: number;
  totalWaste: number;
  totalPseudoRemainder: number;
  totalRealRemainder: number;
  totalMaterial: number;
  executionTime: number;
  lossRateValidation?: LossRateValidation;
  constraintValidation?: ConstraintValidation;
  completeStats?: CompleteStatsData; // 新增：完整的预计算统计数据
  processingStatus?: {
    isCompleted: boolean;
    remaindersFinalized: boolean;
    readyForRendering: boolean;
    completedAt: string;
    dataConsistencyChecked?: boolean;
  };
}

// ==================== API响应类型 ====================

export interface ApiResponse<T = any> {
  success: boolean;
  result?: T;
  error?: string;
  details?: any;
  suggestions?: any[];
  validation?: ConstraintValidation;
  warnings?: any[];
  stack?: string;
}

export interface OptimizationRequest {
  designSteels: DesignSteel[];
  moduleSteels: ModuleSteel[];
  constraints: OptimizationConstraints;
}

export interface OptimizationResponse extends ApiResponse<OptimizationResult> {
  optimizationId?: string;
  stats?: {
    totalCuts: number;
    remaindersGenerated: number;
    remaindersReused: number;
    weldingOperations: number;
  };
  executionTime?: number;
}

// ==================== UI状态类型 ====================

export interface UIState {
  isLoading: boolean;
  error?: string;
  currentStep: number;
  isDarkMode: boolean;
  sidebarCollapsed: boolean;
}

export interface OptimizationState {
  designSteels: DesignSteel[];
  moduleSteels: ModuleSteel[];
  constraints: OptimizationConstraints;
  result?: OptimizationResult;
  isOptimizing: boolean;
  lastOptimizationId?: string;
  validationResult?: ConstraintValidation;
}

export interface FileUploadState {
  isUploading: boolean;
  uploadProgress: number;
  uploadError?: string;
  parsedData?: DesignSteel[];
}

// ==================== 图表数据类型 ====================

export interface ChartData {
  name: string;
  value: number;
  color?: string;
  percentage?: number;
}

export interface LossRateBreakdownData {
  crossSection: string;
  totalMaterial: number;
  wasteLength: number;
  pseudoRemainderLength: number;
  realRemainderLength: number;
  lossRate: number;
  efficiency: number;
}

export interface OptimizationStats {
  activeOptimizations: number;
  totalOptimizations: number;
  successfulOptimizations: number;
  successRate: string;
  averageExecutionTime: number;
}

export interface OptimizationHistory {
  id: string;
  timestamp: string;
  success: boolean;
  executionTime: number;
  totalLossRate: number;
  totalModuleUsed: number;
}

// ==================== 主题系统 ====================

export interface ThemeColors {
  primary: string;
  secondary: string;
  success: string;
  warning: string;
  error: string;
  info: string;
  background: string;
  surface: string;
  text: {
    primary: string;
    secondary: string;
    disabled: string;
  };
  border: string;
  shadow: string;
}

export interface Theme {
  colors: ThemeColors;
  spacing: Record<string, string>;
  borderRadius: Record<string, string>;
  fontSize: Record<string, string>;
  fontWeight: Record<string, number>;
  transition: Record<string, string>;
  zIndex: Record<string, number>;
}

// ==================== 表单类型 ====================

export interface DesignSteelFormData {
  length: string;
  quantity: string;
  crossSection: string;
  displayId?: string;
  componentNumber?: string;
  specification?: string;
  partNumber?: string;
}

export interface ModuleSteelFormData {
  name: string;
  length: string;
}

export interface ConstraintsFormData {
  wasteThreshold: string;
  targetLossRate: string;
  timeLimit: string;
  maxWeldingSegments: string;
}

// ==================== 导出类型 ====================

export interface ExportOptions {
  format: 'excel' | 'pdf';
  includeCharts: boolean;
  includeDetails: boolean;
  includeLossRateBreakdown: boolean;
  customTitle?: string;
}

export interface ExportResponse {
  success: boolean;
  downloadUrl?: string;
  filename?: string;
  error?: string;
}

// ==================== 系统配置 ====================

export interface SystemConfig {
  apiBaseUrl: string;
  uploadMaxSize: number;
  supportedFileTypes: string[];
  defaultConstraints: OptimizationConstraints;
  uiConfig: {
    itemsPerPage: number;
    chartAnimationDuration: number;
    autoSaveInterval: number;
  };
}

// ==================== 工具类型 ====================

export type LoadingState = 'idle' | 'loading' | 'success' | 'error';

export interface AsyncAction<T = any> {
  state: LoadingState;
  data?: T;
  error?: string;
}

export interface PaginationConfig {
  current: number;
  pageSize: number;
  total: number;
  showSizeChanger: boolean;
  showQuickJumper: boolean;
}

// ==================== 事件类型 ====================

export interface OptimizationEvent {
  type: 'start' | 'progress' | 'complete' | 'error' | 'cancel';
  payload?: any;
  timestamp: string;
}

export interface SystemEvent {
  type: 'notification' | 'warning' | 'error';
  message: string;
  duration?: number;
  action?: () => void;
}

// ==================== 常量 ====================

export const REMAINDER_TYPES = {
  PENDING: 'pending' as const,
  PSEUDO: 'pseudo' as const,
  REAL: 'real' as const,
  WASTE: 'waste' as const,
};

export const SOURCE_TYPES = {
  MODULE: 'module' as const,
  REMAINDER: 'remainder' as const,
};

// 🔧 修复：前端约束默认值，对应后端约束配置中心
// 注意：timeLimit在前端以秒为单位显示，后端以毫秒处理
export const DEFAULT_CONSTRAINTS: OptimizationConstraints = {
  wasteThreshold: 100,        // 废料阈值 (mm)
  targetLossRate: 5,          // 目标损耗率 (%)
  timeLimit: 30,              // 计算时间限制 (秒) - 注意单位
  maxWeldingSegments: 0,      // 最大焊接次数 (次)
};

export const UI_STEPS = {
  DESIGN_INPUT: 0,
  MODULE_CONFIG: 1,
  CONSTRAINT_SETTING: 2,
  VALIDATION: 3,
  OPTIMIZATION: 4,
  RESULTS: 5,
} as const; 