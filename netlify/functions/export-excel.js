const { Readable } = require('stream');
const ExcelJS = require('exceljs');

// 处理CORS
const handleCors = (headers = {}) => ({
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization',
  'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
  ...headers
});

// 从优化结果中提取采购清单数据
function extractProcurementData(results) {
  const procurementData = {
    purchaseList: [],
    totalDemand: 0,
    actualPurchase: 0,
    overallUtilization: 0,
    totalLossRate: 0,
    algorithm: '贪心算法'
  };

  try {
    console.log('🔍 开始提取采购数据:', {
      hasSolutions: !!results.solutions,
      solutionsCount: results.solutions?.length || 0,
      hasModuleUsageStats: !!results.moduleUsageStats
    });

    // 优先使用moduleUsageStats（如果有的话）
    if (results.moduleUsageStats && Array.isArray(results.moduleUsageStats)) {
      console.log('✅ 使用moduleUsageStats数据');
      procurementData.purchaseList = results.moduleUsageStats.map((item, index) => ({
        specification: item.specification || '',
        length: item.length || 0,
        quantity: item.totalUsed || 0,
        utilization: item.averageUtilization || 0,
        remark: `利用率: ${((item.averageUtilization || 0) * 100).toFixed(1)}%`
      }));
    } 
    // 回退到从solutions提取
    else if (results.solutions && Array.isArray(results.solutions)) {
      console.log('✅ 从solutions提取数据');
      const moduleUsageMap = new Map();
      
      results.solutions.forEach((solution, solutionIndex) => {
        console.log(`处理解决方案 ${solutionIndex}:`, {
          hasModuleUsage: !!solution.moduleUsage,
          moduleUsageCount: solution.moduleUsage?.length || 0
        });
        
        if (solution.moduleUsage && Array.isArray(solution.moduleUsage)) {
          solution.moduleUsage.forEach(usage => {
            if (usage && usage.specification && usage.length !== undefined) {
              const key = `${usage.specification}_${usage.length}`;
              if (moduleUsageMap.has(key)) {
                const existing = moduleUsageMap.get(key);
                existing.quantity += usage.quantity || 0;
              } else {
                moduleUsageMap.set(key, {
                  specification: usage.specification,
                  length: usage.length,
                  quantity: usage.quantity || 0,
                  utilization: usage.utilization || 0,
                  remark: usage.remark || ''
                });
              }
            }
          });
        }
      });

      procurementData.purchaseList = Array.from(moduleUsageMap.values());
    }

    // 计算统计数据
    procurementData.actualPurchase = procurementData.purchaseList.reduce((sum, item) => sum + item.quantity, 0);
    if (procurementData.purchaseList.length > 0) {
      procurementData.overallUtilization = procurementData.purchaseList.reduce((sum, item) => sum + item.utilization, 0) / procurementData.purchaseList.length;
    }

    console.log('📊 提取结果:', {
      purchaseListCount: procurementData.purchaseList.length,
      actualPurchase: procurementData.actualPurchase,
      overallUtilization: procurementData.overallUtilization
    });

    return procurementData;
  } catch (error) {
    console.error('提取采购数据失败:', error);
    return procurementData;
  }
}

// 生成Excel报告的函数
async function generateExcelReport(data) {
  const workbook = new ExcelJS.Workbook();
  
  // 创建工作表
  const worksheet1 = workbook.addWorksheet('采购清单');
  const worksheet2 = workbook.addWorksheet('优化信息');
  
  // 设置列标题和格式
  worksheet1.columns = [
    { header: '序号', key: 'index', width: 8 },
    { header: '规格', key: 'specification', width: 15 },
    { header: '长度(mm)', key: 'length', width: 12 },
    { header: '数量', key: 'quantity', width: 10 },
    { header: '材料利用率', key: 'utilization', width: 15 },
    { header: '备注', key: 'remark', width: 20 }
  ];
  
  worksheet2.columns = [
    { header: '项目', key: 'item', width: 20 },
    { header: '数值', key: 'value', width: 15 },
    { header: '单位', key: 'unit', width: 10 },
    { header: '说明', key: 'description', width: 30 }
  ];
  
  // 填充采购清单数据
  if (data.purchaseList && Array.isArray(data.purchaseList)) {
    data.purchaseList.forEach((item, index) => {
      worksheet1.addRow({
        index: index + 1,
        specification: item.specification || '',
        length: item.length || 0,
        quantity: item.quantity || 0,
        utilization: item.utilization ? `${(item.utilization * 100).toFixed(1)}%` : '0%',
        remark: item.remark || ''
      });
    });
  }
  
  // 添加采购统计信息
  const stats = [
    { item: '实际采购量', value: data.actualPurchase || 0, unit: '根', description: '实际需要采购的模块钢材数量' },
    { item: '材料利用率', value: data.overallUtilization ? (data.overallUtilization * 100).toFixed(1) : 0, unit: '%', description: '整体材料利用率' },
    { item: '采购规格数', value: data.purchaseList?.length || 0, unit: '种', description: '需要采购的不同规格数量' }
  ];
  
  stats.forEach(stat => {
    worksheet2.addRow(stat);
  });
  
  return workbook;
}

exports.handler = async (event, context) => {
  // 处理OPTIONS预检请求
  if (event.httpMethod === 'OPTIONS') {
    return {
      statusCode: 200,
      headers: handleCors()
    };
  }
  
  if (event.httpMethod !== 'POST') {
    return {
      statusCode: 405,
      headers: handleCors({ 'Content-Type': 'application/json' }),
      body: JSON.stringify({ error: 'Method not allowed' })
    };
  }
  
  // 添加请求基本信息日志
  console.log('📡 收到请求:', {
    httpMethod: event.httpMethod,
    path: event.path,
    headers: event.headers,
    bodyLength: event.body ? event.body.length : 0
  });
  
  try {
    const data = JSON.parse(event.body);
    
    // 添加详细调试日志
    console.log('📊 收到导出请求数据:', {
      hasResults: !!data.results,
      hasExportOptions: !!data.exportOptions,
      resultsType: typeof data.results,
      resultsKeys: data.results ? Object.keys(data.results) : [],
      bodyLength: event.body ? event.body.length : 0,
      fullData: JSON.stringify(data, null, 2)
    });
    
    // 验证数据并提供默认值
    const results = data.results || {};
    const exportOptions = data.exportOptions || {};
    
    if (!results.solutions || !Array.isArray(results.solutions)) {
      console.log('⚠️ 没有解决方案数据，使用空数据');
      results.solutions = [];
    }
    
    // 从优化结果中提取采购清单数据
    const procurementData = extractProcurementData(results);
    
    // 生成Excel
    const workbook = await generateExcelReport(procurementData);
    
    // 写入缓冲区
    const buffer = await workbook.xlsx.writeBuffer();
    
    // 修复文件名编码问题
    const filename = encodeURIComponent(`钢材优化报告_${new Date().toISOString().split('T')[0]}.xlsx`);
    
    return {
      statusCode: 200,
      headers: handleCors({
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename*=UTF-8''${filename}`,
        'Content-Length': buffer.length
      }),
      body: buffer.toString('base64'),
      isBase64Encoded: true
    };
    
  } catch (error) {
    console.error('Export error:', error);
    return {
      statusCode: 500,
      headers: handleCors({ 'Content-Type': 'application/json' }),
      body: JSON.stringify({ error: 'Failed to generate Excel report', details: error.message })
    };
  }
};