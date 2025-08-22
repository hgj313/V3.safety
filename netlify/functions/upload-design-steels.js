/**
 * Netlify Function - 设计钢材文件上传 (V3增强版)
 * 支持智能字段识别、数据清洗、容错处理
 */
const multiparty = require('multiparty');

// 智能字段映射字典
const FIELD_MAPPING = {
  length: {
    keywords: ['长度', 'length', '长度(mm)', '长度 mm', '长度（mm）', '长度mm', '钢材长度', 'Length(mm)', 'LENGTH', '长度/mm'],
    required: true,
    defaultValue: null
  },
  quantity: {
    keywords: ['数量', 'quantity', 'qty', '件数', '根数', 'Quantity', '数量(件)', '数量（件）', 'QTY', '数量/件'],
    required: true,
    defaultValue: null
  },
  crossSection: {
    keywords: ['截面面积', '截面', 'cross section', '面积', '截面积(mm²)', '截面积（mm²）', 'CrossSection', 'crossSection', '截面面积/mm²'],
    required: false,
    defaultValue: 1000
  },
  specification: {
    keywords: ['规格', 'spec', 'specification', '材质', '钢材规格', '型号', 'Specification', 'SPEC', '规格型号'],
    required: false,
    defaultValue: '未知规格'
  },
  componentNumber: {
    keywords: ['构件编号', '构件号', 'ComponentNumber', 'componentNumber', '编号', '构件', '零件编号'],
    required: false,
    defaultValue: null
  },
  partNumber: {
    keywords: ['部件编号', '部件号', 'PartNumber', 'partNumber', '零件号', '部件'],
    required: false,
    defaultValue: null
  },
  material: {
    keywords: ['材质', 'Material', 'material', '钢材材质', '材料', '钢材类型'],
    required: false,
    defaultValue: ''
  },
  note: {
    keywords: ['备注', 'Note', 'note', '说明', '备注说明', '描述', 'description'],
    required: false,
    defaultValue: ''
  }
};

// 智能字段识别器
function identifyFields(headers) {
  const fieldMapping = {};
  const unidentified = [];
  const confidence = {};

  console.log('🔍 开始智能字段识别，表头:', headers);

  // 为每个字段寻找最佳匹配
  Object.keys(FIELD_MAPPING).forEach(fieldKey => {
    const fieldConfig = FIELD_MAPPING[fieldKey];
    let bestMatch = null;
    let bestScore = 0;

    headers.forEach(header => {
      const cleanHeader = header.trim().toLowerCase();
      
      // 精确匹配
      const exactMatch = fieldConfig.keywords.find(keyword => 
        cleanHeader === keyword.toLowerCase() || 
        cleanHeader.includes(keyword.toLowerCase())
      );
      
      if (exactMatch) {
        const score = cleanHeader === exactMatch.toLowerCase() ? 100 : 80;
        if (score > bestScore) {
          bestMatch = header;
          bestScore = score;
        }
      }
    });

    if (bestMatch) {
      fieldMapping[fieldKey] = bestMatch;
      confidence[fieldKey] = bestScore;
      console.log(`✅ ${fieldKey}: "${bestMatch}" (置信度: ${bestScore}%)`);
    } else if (fieldConfig.required) {
      console.log(`❌ 必需字段 ${fieldKey} 未找到匹配`);
    } else {
      console.log(`⚠️ 可选字段 ${fieldKey} 未找到，将使用默认值`);
    }
  });

  // 找出未识别的列
  headers.forEach(header => {
    const isUsed = Object.values(fieldMapping).includes(header);
    if (!isUsed) {
      unidentified.push(header);
    }
  });

  return {
    fieldMapping,
    confidence,
    unidentified,
    requiredFieldsMissing: Object.keys(FIELD_MAPPING)
      .filter(key => FIELD_MAPPING[key].required && !fieldMapping[key])
  };
}

// 数据清洗引擎
function cleanData(value, fieldType) {
  if (value === null || value === undefined || value === '') {
    return null;
  }

  // 转为字符串处理
  let cleanValue = String(value).trim();

  switch (fieldType) {
    case 'length':
    case 'crossSection':
      // 数值类型：提取数字，去除单位
      cleanValue = cleanValue
        .replace(/[^\d.-]/g, '') // 只保留数字、小数点、负号
        .replace(/,/g, ''); // 去除千分位符号
      
      const numValue = parseFloat(cleanValue);
      return isNaN(numValue) ? null : numValue;

    case 'quantity':
      // 整数类型：提取整数
      cleanValue = cleanValue
        .replace(/[^\d]/g, '') // 只保留数字
        .replace(/,/g, ''); // 去除千分位符号
      
      const intValue = parseInt(cleanValue);
      return isNaN(intValue) ? null : intValue;

    case 'specification':
    case 'material':
    case 'componentNumber':
    case 'partNumber':
    case 'note':
      // 文本类型：去除多余空格，处理特殊字符
      cleanValue = cleanValue
        .replace(/\s+/g, ' ') // 多个空格变成一个
        .replace(/^[-\s]*$/, '') // 如果只有横线和空格，视为空
        .trim();
      
      return cleanValue === '' ? null : cleanValue;

    default:
      return cleanValue;
  }
}

// 生成自动编号
function generateAutoNumber(index, prefix = 'AUTO') {
  return `${prefix}${String(index + 1).padStart(3, '0')}`;
}

// 生成显示ID - 使用与前端完全一致的逻辑
function generateDisplayIds(steels) {
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
  const groups = {};
  sorted.forEach(steel => {
    const specification = steel.specification || '未知规格';
    const crossSection = Math.round(steel.crossSection);
    const groupKey = `${specification}_${crossSection}`;
    
    if (!groups[groupKey]) {
      groups[groupKey] = [];
    }
    groups[groupKey].push(steel);
  });

  // 生成字母前缀
  const generateLetterPrefix = (index) => {
    if (index < 26) {
      return String.fromCharCode(65 + index); // A, B, C, ..., Z
    } else {
      const firstLetter = Math.floor(index / 26) - 1;
      const secondLetter = index % 26;
      return String.fromCharCode(65 + firstLetter) + String.fromCharCode(65 + secondLetter);
    }
  };

  // 按组合键排序
  const sortedGroupKeys = Object.keys(groups).sort();
  
  const result = [];
  sortedGroupKeys.forEach((groupKey, groupIndex) => {
    const letterPrefix = generateLetterPrefix(groupIndex);
    const groupSteels = groups[groupKey];
    
    groupSteels.forEach((steel, itemIndex) => {
      result.push({
        ...steel,
        displayId: `${letterPrefix}${itemIndex + 1}`,
        groupKey: groupKey
      });
    });
  });

  return result;
}

// 处理Excel文件 - 增强版
function processExcelFile(fileBuffer, filename) {
  try {
    const XLSX = require('xlsx');
    
    console.log('=== Excel文件智能解析开始 ===');
    console.log('文件名:', filename);
    console.log('文件大小:', fileBuffer.length, '字节');

    const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // 读取原始数据
    const rawData = XLSX.utils.sheet_to_json(worksheet);
    console.log('原始数据行数:', rawData.length);

    if (rawData.length === 0) {
      throw new Error('Excel文件为空或无法读取数据');
    }

    // 获取表头并进行智能字段识别
    const headers = Object.keys(rawData[0]);
    const fieldAnalysis = identifyFields(headers);

    console.log('字段识别结果:', fieldAnalysis);

    // 检查必需字段
    if (fieldAnalysis.requiredFieldsMissing.length > 0) {
      throw new Error(`缺少必需字段: ${fieldAnalysis.requiredFieldsMissing.join(', ')}`);
    }

    // 数据处理统计
    const stats = {
      totalRows: rawData.length,
      validRows: 0,
      skippedRows: 0,
      cleaningActions: {
        unitRemoved: 0,
        spacesCleaned: 0,
        autoNumberGenerated: 0,
        defaultValuesUsed: 0
      }
    };

    // 逐行处理数据
    const processedSteels = [];
    rawData.forEach((row, index) => {
      try {
        const steel = {
          id: `design_${Date.now()}_${index}`
        };

        // 处理每个字段
        Object.keys(FIELD_MAPPING).forEach(fieldKey => {
          const fieldConfig = FIELD_MAPPING[fieldKey];
          const sourceColumn = fieldAnalysis.fieldMapping[fieldKey];
          
          let rawValue = null;
          if (sourceColumn && row[sourceColumn] !== undefined) {
            rawValue = row[sourceColumn];
          }

          // 数据清洗
          const cleanedValue = cleanData(rawValue, fieldKey);
          
          if (cleanedValue !== null) {
            steel[fieldKey] = cleanedValue;
            
            // 统计清洗动作
            if (rawValue !== cleanedValue) {
              if (typeof rawValue === 'string' && /mm|cm|m²/.test(rawValue)) {
                stats.cleaningActions.unitRemoved++;
              }
              if (typeof rawValue === 'string' && rawValue.trim() !== rawValue) {
                stats.cleaningActions.spacesCleaned++;
              }
            }
          } else if (fieldConfig.required) {
            throw new Error(`第${index + 1}行缺少必需字段: ${fieldKey}`);
          } else if (fieldConfig.defaultValue !== null) {
            steel[fieldKey] = fieldConfig.defaultValue;
            stats.cleaningActions.defaultValuesUsed++;
          }
        });

        // 为空的编号字段生成自动编号
        if (!steel.componentNumber) {
          steel.componentNumber = generateAutoNumber(index, 'GJ');
          stats.cleaningActions.autoNumberGenerated++;
        }
        if (!steel.partNumber) {
          steel.partNumber = generateAutoNumber(index, 'BJ');
          stats.cleaningActions.autoNumberGenerated++;
        }

        // 验证关键数据
        if (steel.length > 0 && steel.quantity > 0 && steel.crossSection > 0) {
          processedSteels.push(steel);
          stats.validRows++;
        } else {
          console.log(`跳过第${index + 1}行: 数据不完整`, steel);
          stats.skippedRows++;
        }

      } catch (error) {
        console.log(`跳过第${index + 1}行: ${error.message}`);
        stats.skippedRows++;
      }
    });

    // 生成最终的显示ID
    const finalSteels = generateDisplayIds(processedSteels);

    console.log('=== Excel文件智能解析完成 ===');

    return {
      success: true,
      message: `智能解析成功！处理了 ${stats.validRows} 条有效数据`,
      designSteels: finalSteels,
      analysisReport: {
        fieldMapping: fieldAnalysis.fieldMapping,
        confidence: fieldAnalysis.confidence,
        unidentifiedColumns: fieldAnalysis.unidentified,
        dataStats: stats,
        cleaningReport: [
          stats.cleaningActions.unitRemoved > 0 ? `自动去除了 ${stats.cleaningActions.unitRemoved} 个单位标识` : null,
          stats.cleaningActions.spacesCleaned > 0 ? `清理了 ${stats.cleaningActions.spacesCleaned} 个格式问题` : null,
          stats.cleaningActions.defaultValuesUsed > 0 ? `使用了 ${stats.cleaningActions.defaultValuesUsed} 个默认值` : null,
          stats.cleaningActions.autoNumberGenerated > 0 ? `生成了 ${stats.cleaningActions.autoNumberGenerated} 个自动编号` : null
        ].filter(Boolean)
      },
      debugInfo: {
        原始行数: stats.totalRows,
        有效数据: stats.validRows,
        跳过行数: stats.skippedRows,
        字段识别: Object.keys(fieldAnalysis.fieldMapping).length,
        处理时间: new Date().toISOString(),
        版本信息: 'V3智能解析增强版'
      }
    };
  } catch (error) {
    console.error('=== Excel文件解析错误 ===', error);
    throw error;
  }
}

exports.handler = async (event, context) => {
  // 处理CORS预检请求
  if (event.httpMethod === 'OPTIONS') {
    return {
      statusCode: 200,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Headers': 'Content-Type',
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS'
      },
      body: ''
    };
  }

  try {
    if (event.httpMethod !== 'POST') {
      return {
        statusCode: 405,
        headers: {
          'Content-Type': 'application/json',
          'Access-Control-Allow-Origin': '*'
        },
        body: JSON.stringify({
          success: false,
          error: '仅支持POST请求'
        })
      };
    }

    // 处理JSON格式的base64数据
    if (event.headers['content-type']?.includes('application/json')) {
      const requestData = JSON.parse(event.body);
      const { filename, data, type } = requestData;
      
      console.log('智能解析上传:', { filename, type, size: data.length });
      
      // 解析base64数据
      const buffer = Buffer.from(data, 'base64');
      
      // 智能处理文件
      const result = processExcelFile(buffer, filename);
      
      return {
        statusCode: 200,
        headers: {
          'Content-Type': 'application/json',
          'Access-Control-Allow-Origin': '*'
        },
        body: JSON.stringify(result)
      };
    }

    // 处理multipart/form-data格式 (暂时简化处理)
    return {
      statusCode: 400,
      headers: {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*'
      },
      body: JSON.stringify({
        success: false,
        error: 'Netlify Functions目前仅支持JSON格式的base64文件上传'
      })
    };

  } catch (error) {
    console.error('❌ 智能解析失败:', error);
    return {
      statusCode: 500,
      headers: {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*'
      },
      body: JSON.stringify({
        success: false,
        error: `智能解析失败: ${error.message}`
      })
    };
  }
}; 