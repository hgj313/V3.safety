/**
 * Netlify Function - 约束条件验证
 */
let OptimizationService;
try {
  OptimizationService = require('../../api/services/OptimizationService');
} catch (error) {
  console.warn('优化服务未找到，将使用模拟模式');
}

// 创建模拟约束验证服务
function createMockValidationService() {
  return {
    validateWeldingConstraints: async (data) => ({
      success: true,
      validation: {
        isValid: true,
        violations: [],
        suggestions: [
          '建议使用更高强度的焊接材料',
          '建议增加焊接点数量以提高结构稳定性'
        ]
      }
    })
  };
}

// 获取验证服务实例
function getValidationService() {
  if (OptimizationService) {
    return new OptimizationService();
  } else {
    return createMockValidationService();
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

    const requestData = JSON.parse(event.body);
    console.log('🔍 收到约束验证请求 (Netlify)');

    const service = getValidationService();
    const result = await service.validateWeldingConstraints(requestData);

    return {
      statusCode: 200,
      headers: {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*'
      },
      body: JSON.stringify(result)
    };
  } catch (error) {
    console.error('❌ 约束验证API错误:', error);
    return {
      statusCode: 500,
      headers: {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*'
      },
      body: JSON.stringify({
        success: false,
        error: `约束验证失败: ${error.message}`
      })
    };
  }
}; 