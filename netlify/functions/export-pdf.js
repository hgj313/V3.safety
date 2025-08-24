// 处理CORS
const handleCors = (headers = {}) => ({
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization',
  'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
  ...headers
});

// 生成PDF HTML内容
function generatePDFHTML(data) {
  const { results, exportOptions } = data;
  
  return `
    <!DOCTYPE html>
    <html lang="zh-CN">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>钢材优化报告</title>
      <style>
        body { font-family: Arial, sans-serif; margin: 20px; color: #333; }
        .header { text-align: center; margin-bottom: 30px; border-bottom: 2px solid #007bff; padding-bottom: 10px; }
        .section { margin-bottom: 25px; }
        .section h2 { color: #007bff; border-bottom: 1px solid #eee; padding-bottom: 5px; }
        table { width: 100%; border-collapse: collapse; margin: 10px 0; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f8f9fa; font-weight: bold; }
        .stats { display: flex; justify-content: space-between; margin: 15px 0; }
        .stat-box { background: #f8f9fa; padding: 10px; border-radius: 5px; text-align: center; flex: 1; margin: 0 5px; }
        .highlight { color: #007bff; font-weight: bold; }
      </style>
    </head>
    <body>
      <div class="header">
        <h1>钢材优化报告</h1>
        <p>生成时间: ${new Date().toLocaleString('zh-CN')}</p>
      </div>

      <div class="section">
        <h2>📊 优化结果总览</h2>
        <div class="stats">
          <div class="stat-box">
            <div>总需求量</div>
            <div class="highlight">${results.totalDemand || 0} 根</div>
          </div>
          <div class="stat-box">
            <div>实际采购量</div>
            <div class="highlight">${results.actualPurchase || 0} 根</div>
          </div>
          <div class="stat-box">
            <div>材料利用率</div>
            <div class="highlight">${results.overallUtilization ? (results.overallUtilization * 100).toFixed(1) : 0}%</div>
          </div>
          <div class="stat-box">
            <div>总损耗率</div>
            <div class="highlight">${results.totalLossRate ? (results.totalLossRate * 100).toFixed(1) : 0}%</div>
          </div>
        </div>
      </div>

      ${exportOptions.includePurchaseList && results.purchaseList ? `
      <div class="section">
        <h2>📋 采购清单</h2>
        <table>
          <thead>
            <tr>
              <th>序号</th>
              <th>规格</th>
              <th>长度(mm)</th>
              <th>数量</th>
              <th>材料利用率</th>
            </tr>
          </thead>
          <tbody>
            ${results.purchaseList.map((item, index) => `
              <tr>
                <td>${index + 1}</td>
                <td>${item.specification || '-'}</td>
                <td>${item.length || 0}</td>
                <td>${item.quantity || 0}</td>
                <td>${item.utilization ? (item.utilization * 100).toFixed(1) + '%' : '0%'}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
      ` : ''}

      ${exportOptions.includeDesignSteels && results.designSteels ? `
      <div class="section">
        <h2>🔧 设计钢材清单</h2>
        <table>
          <thead>
            <tr>
              <th>序号</th>
              <th>规格</th>
              <th>需求长度(mm)</th>
              <th>数量</th>
              <th>备注</th>
            </tr>
          </thead>
          <tbody>
            ${results.designSteels.map((item, index) => `
              <tr>
                <td>${index + 1}</td>
                <td>${item.specification || '-'}</td>
                <td>${item.requiredLength || 0}</td>
                <td>${item.quantity || 0}</td>
                <td>${item.remark || '-'}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
      ` : ''}

      <div class="section">
        <h2>⚙️ 系统信息</h2>
        <p><strong>优化算法:</strong> ${results.algorithm || '贪心算法'}</p>
        <p><strong>生成时间:</strong> ${new Date().toLocaleString('zh-CN')}</p>
        <p><strong>系统版本:</strong> V3.0</p>
      </div>
    </body>
    </html>
  `;
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
    
    // 生成PDF HTML
    const htmlContent = generatePDFHTML(data);
    
    return {
      statusCode: 200,
      headers: handleCors({
        'Content-Type': 'text/html',
        'Content-Disposition': `attachment; filename="钢材优化报告_${new Date().toISOString().split('T')[0]}.html"`
      }),
      body: htmlContent
    };
    
  } catch (error) {
    console.error('PDF export error:', error);
    return {
      statusCode: 500,
      headers: handleCors({ 'Content-Type': 'application/json' }),
      body: JSON.stringify({ error: 'Failed to generate PDF report', details: error.message })
    };
  }
};