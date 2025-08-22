/**
 * Netlify Function - 系统健康检查
 */
exports.handler = async (event, context) => {
  try {
    return {
      statusCode: 200,
      headers: {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Headers': 'Content-Type',
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS'
      },
      body: JSON.stringify({
        success: true,
        message: '钢材采购优化系统 V3.0 运行正常 (Netlify)',
        version: '3.0.0',
        timestamp: new Date().toISOString(),
        platform: 'Netlify Functions',
        features: ['优化算法', '文件上传', '数据分析', '结果导出']
      })
    };
  } catch (error) {
    return {
      statusCode: 500,
      headers: {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*'
      },
      body: JSON.stringify({
        success: false,
        error: `健康检查失败: ${error.message}`
      })
    };
  }
}; 