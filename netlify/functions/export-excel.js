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
    // 从优化结果中提取模块钢材使用统计
    if (results.solutions && Array.isArray(results.solutions)) {
      const moduleUsageMap = new Map();
      
      results.solutions.forEach(solution => {
        if (solution.moduleUsage) {
          solution.moduleUsage.forEach(usage => {
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
          });
        }
      });

      procurementData.purchaseList = Array.from(moduleUsageMap.values());
      procurementData.actualPurchase = procurementData.purchaseList.reduce((sum, item) => sum + item.quantity, 0);
      
      // 计算平均利用率
      if (procurementData.purchaseList.length > 0) {
        procurementData.overallUtilization = procurementData.purchaseList.reduce((sum, item) => sum + item.utilization, 0) / procurementData.purchaseList.length;
      }
    }

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
  
  try {
    const data = JSON.parse(event.body);
    
    // 验证必需的数据
    if (!data.results || !data.exportOptions) {
      return {
        statusCode: 400,
        headers: handleCors({ 'Content-Type': 'application/json' }),
        body: JSON.stringify({ error: 'Missing required data' })
      };
    }
    
    // 从优化结果中提取采购清单数据
    const procurementData = extractProcurementData(data.results);
    
    // 生成Excel
    const workbook = await generateExcelReport(procurementData);
    
    // 写入缓冲区
    const buffer = await workbook.xlsx.writeBuffer();
    
    return {
      statusCode: 200,
      headers: handleCors({
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="钢材优化报告_${new Date().toISOString().split('T')[0]}.xlsx"`,
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