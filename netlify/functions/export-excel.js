const { Readable } = require('stream');
const ExcelJS = require('exceljs');

// 处理CORS
const handleCors = (headers = {}) => ({
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization',
  'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
  ...headers
});

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
  
  // 填充数据
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
  
  // 添加优化统计信息
  const stats = [
    { item: '总需求量', value: data.totalDemand || 0, unit: '根', description: '设计钢材总数量' },
    { item: '实际采购量', value: data.actualPurchase || 0, unit: '根', description: '实际需要采购数量' },
    { item: '材料利用率', value: data.overallUtilization ? (data.overallUtilization * 100).toFixed(1) : 0, unit: '%', description: '整体材料利用率' },
    { item: '总损耗率', value: data.totalLossRate ? (data.totalLossRate * 100).toFixed(1) : 0, unit: '%', description: '整体损耗率' },
    { item: '优化算法', value: data.algorithm || '贪心算法', unit: '', description: '使用的优化算法' }
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
    
    // 生成Excel
    const workbook = await generateExcelReport(data.results);
    
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