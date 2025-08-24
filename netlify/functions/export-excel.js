const { Readable } = require('stream');
const ExcelJS = require('exceljs');

// 处理CORS
const handleCors = (headers = {}) => ({
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization',
  'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
  ...headers
});

// 重新设计的采购数据提取函数 - 按规格汇总数量
function extractProcurementData(results) {
  const procurementData = {
    purchaseList: [],
    totalDemand: 0,
    actualPurchase: 0,
    overallUtilization: 0.95,
    totalLossRate: 5,
    algorithm: '贪心算法'
  };

  try {
    console.log('🔍 重新提取采购清单数据，按规格汇总数量');
    
    let rawData = [];
    
    // 获取原始数据
    if (results.moduleUsageStats && Array.isArray(results.moduleUsageStats.sortedStats)) {
      rawData = results.moduleUsageStats.sortedStats;
      console.log('✅ 使用sortedStats数据，共', rawData.length, '条记录');
    } else if (results.moduleUsageStats && Array.isArray(results.moduleUsageStats)) {
      rawData = results.moduleUsageStats;
      console.log('✅ 使用moduleUsageStats数据，共', rawData.length, '条记录');
    } else {
      console.log('⚠️ 没有找到模块使用数据');
      return procurementData;
    }

    // 按规格和长度分组汇总
    const specificationMap = new Map();
    
    rawData.forEach(item => {
      const spec = item.specification || '';
      const length = Number(item.length) || 0;
      const count = Number(item.count) || Number(item.totalUsed) || 0;
      const totalLength = Number(item.totalLength) || (length * count);
      
      const key = `${spec}_${length}`;
      
      if (specificationMap.has(key)) {
        const existing = specificationMap.get(key);
        existing.quantity += count;
        existing.totalLength += totalLength;
      } else {
        specificationMap.set(key, {
          specification: spec,
          length: length,
          quantity: count,
          utilization: Number(item.averageUtilization) || 0.95,
          totalLength: totalLength,
          remark: `规格: ${spec}, 长度: ${length}mm`
        });
      }
    });

    // 转换为数组并排序
    procurementData.purchaseList = Array.from(specificationMap.values())
      .sort((a, b) => {
        if (a.specification !== b.specification) {
          return a.specification.localeCompare(b.specification);
        }
        return a.length - b.length;
      });

    // 计算总数量
    procurementData.actualPurchase = procurementData.purchaseList.reduce((sum, item) => sum + item.quantity, 0);
    procurementData.totalDemand = procurementData.purchaseList.reduce((sum, item) => sum + item.totalLength, 0);

    console.log('📊 汇总后的采购清单:', {
      规格种类: procurementData.purchaseList.length,
      总数量: procurementData.actualPurchase,
      总长度: procurementData.totalDemand,
      明细: procurementData.purchaseList.map(item => ({
        规格: item.specification,
        长度: item.length,
        数量: item.quantity
      }))
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
  let totalCost = 0;
  let totalQuantity = 0;
  let totalMaterial = 0;

  if (data.purchaseList && Array.isArray(data.purchaseList)) {
    data.purchaseList.forEach((item, index) => {
      const totalLength = item.totalLength || (item.length * item.quantity);
      const row = {
        index: index + 1,
        specification: item.specification || '',
        length: item.length || 0,
        quantity: item.quantity || 0,
        utilization: item.utilization ? `${(item.utilization * 100).toFixed(1)}%` : '0%',
        remark: item.remark || ''
      };
      
      const dataRow = worksheet1.addRow(row);
      dataRow.height = 20;
      
      // 交替行颜色
      if (index % 2 === 0) {
        dataRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8F9FA' } };
      }
      
      // 数据格式化
      dataRow.getCell('length').numFmt = '#,##0';
      dataRow.alignment = { horizontal: 'center', vertical: 'middle' };
      
      totalCost += totalLength * 0.007;
      totalQuantity += item.quantity;
      totalMaterial += totalLength;
    });

    // 添加汇总行
    const summaryRow = worksheet1.addRow({
      index: '',
      specification: '合计',
      length: '',
      quantity: totalQuantity || data.actualPurchase || 0,
      utilization: data.overallUtilization ? `${(data.overallUtilization * 100).toFixed(1)}%` : '0%',
      remark: '总采购成本'
    });
    summaryRow.font = { bold: true };
    summaryRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE9ECEF' } };
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