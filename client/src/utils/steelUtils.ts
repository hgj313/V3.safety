import { DesignSteel, ModuleSteel } from '../types';

// 生成设计钢材显示编号 - 真正的规格化设计（规格+截面面积组合分组）
export const generateDisplayIds = (steels: DesignSteel[]): DesignSteel[] => {
  // 三级排序：规格 → 截面面积 → 长度
  const sorted = [...steels].sort((a, b) => {
    const specA = a.specification || '未知规格';
    const specB = b.specification || '未知规格';
    
    // 第一级：按规格排序
    if (specA !== specB) {
      return specA.localeCompare(specB);
    }
    
    // 第二级：同规格内按截面面积排序
    if (a.crossSection !== b.crossSection) {
      return a.crossSection - b.crossSection;
    }
    
    // 第三级：同规格同截面面积内按长度排序
    return a.length - b.length;
  });

  // 按规格+截面面积组合分组
  const groups: Record<string, DesignSteel[]> = {};
  sorted.forEach(steel => {
    const specification = steel.specification || '未知规格';
    const crossSection = Math.round(steel.crossSection); // 四舍五入处理浮点数
    const groupKey = `${specification}_${crossSection}`; // 组合键
    
    if (!groups[groupKey]) {
      groups[groupKey] = [];
    }
    groups[groupKey].push(steel);
  });

  // 生成字母前缀（支持超过26个组合）
  const generateLetterPrefix = (index: number): string => {
    if (index < 26) {
      return String.fromCharCode(65 + index); // A, B, C, ..., Z
    } else {
      const firstLetter = Math.floor(index / 26) - 1;
      const secondLetter = index % 26;
      return String.fromCharCode(65 + firstLetter) + String.fromCharCode(65 + secondLetter); // AA, AB, AC...
    }
  };

  // 按组合键排序（确保字母前缀分配的一致性）
  const sortedGroupKeys = Object.keys(groups).sort();
  
  const result: DesignSteel[] = [];
  sortedGroupKeys.forEach((groupKey, groupIndex) => {
    const letterPrefix = generateLetterPrefix(groupIndex);
    const groupSteels = groups[groupKey];
    
    groupSteels.forEach((steel, itemIndex) => {
      result.push({
        ...steel,
        displayId: `${letterPrefix}${itemIndex + 1}`, // A1, A2, B1, B2, AA1, AB1...
        groupKey: groupKey // 添加组合键信息用于调试
      });
    });
  });

  console.log('🎯 真正的规格化编号系统完成:', {
    总钢材数: result.length,
    分组数: sortedGroupKeys.length,
    分组详情: sortedGroupKeys.map((key, index) => ({
      组合键: key,
      字母前缀: generateLetterPrefix(index),
      钢材数量: groups[key].length
    })),
    前5个示例: result.slice(0, 5).map(s => ({ 
      id: s.id, 
      displayId: s.displayId, 
      specification: s.specification,
      crossSection: s.crossSection,
      length: s.length,
      groupKey: s.groupKey
    }))
  });

  return result;
};

// V3规格化数据重组 - 直接按规格组织，无需映射转换
export const regroupOptimizationResultsBySpecification = (
  solutions: Record<string, any>
): Record<string, any> => {
  // V3系统中，solutions的key应该直接是规格名称，而不是截面面积
  // 如果后端还在使用截面面积作为key，这里需要适配
  const specificationResults: Record<string, any> = {};
  
  Object.entries(solutions).forEach(([key, solution]) => {
    // 尝试从solution中获取规格信息
    const specification = solution.specification || key;
    
    console.log(`🔍 V3规格化重组: ${key} → 规格: ${specification}`);
    
    specificationResults[specification] = {
      ...solution,
      specification: specification
    };
  });
  
  return specificationResults;
};

// 按规格分组（V3核心功能）
export const groupBySpecification = (steels: DesignSteel[]): Record<string, DesignSteel[]> => {
  const groups: Record<string, DesignSteel[]> = {};
  steels.forEach(steel => {
    const spec = steel.specification || '未知规格';
    if (!groups[spec]) {
      groups[spec] = [];
    }
    groups[spec].push(steel);
  });
  return groups;
};

// V3规格化统计 - 按规格计算实际使用情况
export const calculateActualUsageBySpecification = (
  designSteels: DesignSteel[], 
  moduleSteels: ModuleSteel[]
): Record<string, {
  specification: string;
  totalDesignLength: number;
  actualModuleCount: number;
  actualTotalLength: number;
  actualWaste: number;
  actualLossRate: number;
}> => {
  const specGroups = groupBySpecification(designSteels);
  const results: Record<string, any> = {};
  
  Object.entries(specGroups).forEach(([specification, steels]) => {
    // 计算该规格的设计钢材总长度
    const totalDesignLength = steels.reduce((sum, steel) => 
      sum + (steel.length * steel.quantity), 0
    );
    
    // 找到对应规格的模数钢材（目前ModuleSteel没有specification字段，使用第一个）
    // TODO: 后续需要给ModuleSteel添加specification字段以支持规格匹配
    const matchingModule = moduleSteels[0]; // 暂时使用第一个模数钢材
    
    const moduleLength = matchingModule?.length || 12000;
    
    // 计算实际需要的模数钢材数量
    const actualModuleCount = Math.ceil(totalDesignLength / moduleLength);
    
    // 计算实际总长度和损耗
    const actualTotalLength = actualModuleCount * moduleLength;
    const actualWaste = actualTotalLength - totalDesignLength;
    const actualLossRate = (actualWaste / actualTotalLength) * 100;
    
    results[specification] = {
      specification,
      totalDesignLength,
      actualModuleCount,
      actualTotalLength,
      actualWaste,
      actualLossRate
    };
  });
  
  return results;
};

// 格式化数字显示
export const formatNumber = (num: number, decimals: number = 2): string => {
  return num.toLocaleString('zh-CN', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals
  });
};

// 验证设计钢材数据
export const validateDesignSteel = (steel: Partial<DesignSteel>): string | null => {
  if (!steel.length || steel.length <= 0) {
    return '长度必须大于0';
  }
  if (!steel.quantity || steel.quantity <= 0) {
    return '数量必须大于0';
  }
  if (!steel.crossSection || steel.crossSection <= 0) {
    return '截面面积必须大于0';
  }
  if (!steel.specification || steel.specification.trim() === '') {
    return '规格不能为空';
  }
  return null;
};

// 计算设计钢材总体积
export const calculateTotalVolume = (steels: DesignSteel[]): number => {
  return steels.reduce((total, steel) => {
    return total + steel.length * steel.quantity * steel.crossSection;
  }, 0);
};

// 生成唯一ID
export const generateId = (): string => {
  return Date.now().toString(36) + Math.random().toString(36).substr(2);
};

// 建立规格到截面面积的映射（兼容功能）
export const buildSpecificationMapping = (steels: DesignSteel[]): Record<string, number> => {
  const mapping: Record<string, number> = {};
  steels.forEach(steel => {
    if (steel.specification && steel.crossSection) {
      if (!mapping[steel.specification]) {
        mapping[steel.specification] = steel.crossSection;
      }
    }
  });
  return mapping;
};

// 计算总体实际使用情况（简化版本，推荐使用calculateActualUsageBySpecification）
export const calculateActualUsage = (
  designSteels: DesignSteel[], 
  moduleSteels: ModuleSteel[]
): {
  actualModuleCount: number;
  actualTotalLength: number;
  actualWaste: number;
  actualLossRate: number;
  totalDesignLength: number;
} => {
  const totalDesignLength = designSteels.reduce((sum, steel) => 
    sum + (steel.length * steel.quantity), 0
  );
  
  const moduleLength = moduleSteels[0]?.length || 12000;
  const actualModuleCount = Math.ceil(totalDesignLength / moduleLength);
  const actualTotalLength = actualModuleCount * moduleLength;
  const actualWaste = actualTotalLength - totalDesignLength;
  const actualLossRate = (actualWaste / actualTotalLength) * 100;
  
  return {
    actualModuleCount,
    actualTotalLength,
    actualWaste,
    actualLossRate,
    totalDesignLength
  };
};

/**
 * 根据设计钢材列表和总损耗率计算总采购金额
 * @param designSteels 设计钢材列表
 * @param totalLossRate 总损耗率
 * @returns 总采购金额
 */
export const calculateTotalPurchaseAmount = (
  designSteels: DesignSteel[],
  totalLossRate: number
): number => {
  const totalAmount = calculateTotalVolume(designSteels);
  return totalAmount * (1 + totalLossRate / 100);
};

/**
 * 格式化数字为千分位
 * @param num 数字
 * @returns 格式化后的字符串
 */
export const formatNumberToThousands = (num: number): string => {
  return num.toString().replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,');
}; 