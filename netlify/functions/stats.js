/**
 * Netlify Function - 系统统计信息
 * 提供系统运行状态、性能指标和使用统计
 */
const { neon } = require('@neondatabase/serverless');

// 数据库连接
let sql = null;

async function initDatabase() {
  if (!sql) {
    const databaseUrl = process.env.NETLIFY_DATABASE_URL || process.env.DATABASE_URL;
    if (!databaseUrl) {
      throw new Error('DATABASE_URL 环境变量未设置');
    }
    sql = neon(databaseUrl);
  }
}

// 获取系统统计信息
async function getSystemStats() {
  try {
    await initDatabase();
    
    // 获取任务统计
    const taskStats = await sql`
      SELECT 
        COUNT(*) as total_tasks,
        COUNT(*) FILTER (WHERE status = 'pending') as pending_tasks,
        COUNT(*) FILTER (WHERE status = 'running') as running_tasks,
        COUNT(*) FILTER (WHERE status = 'completed') as completed_tasks,
        COUNT(*) FILTER (WHERE status = 'failed') as failed_tasks,
        COUNT(*) FILTER (WHERE status = 'cancelled') as cancelled_tasks,
        CASE 
          WHEN COUNT(*) > 0 THEN 
            ROUND((COUNT(*) FILTER (WHERE status = 'completed')::float / COUNT(*) * 100), 2)
          ELSE 0
        END as success_rate,
        ROUND(AVG(execution_time) FILTER (WHERE execution_time IS NOT NULL), 2) as avg_execution_time,
        COUNT(*) FILTER (WHERE created_at > NOW() - INTERVAL '24 hours') as last_24h_tasks,
        COUNT(*) FILTER (WHERE created_at > NOW() - INTERVAL '7 days') as last_7d_tasks,
        COUNT(*) FILTER (WHERE created_at > NOW() - INTERVAL '30 days') as last_30d_tasks
      FROM optimization_tasks
    `;

    // 获取最近任务
    const recentTasks = await sql`
      SELECT 
        id,
        status,
        progress,
        message,
        execution_time,
        created_at
      FROM optimization_tasks
      ORDER BY created_at DESC
      LIMIT 10
    `;

    // 获取性能统计
    const performanceStats = await sql`
      SELECT 
        DATE_TRUNC('day', created_at) as date,
        COUNT(*) as total_tasks,
        COUNT(*) FILTER (WHERE status = 'completed') as completed_tasks,
        COUNT(*) FILTER (WHERE status = 'failed') as failed_tasks,
        ROUND(AVG(execution_time) FILTER (WHERE execution_time IS NOT NULL), 2) as avg_execution_time
      FROM optimization_tasks
      WHERE created_at > NOW() - INTERVAL '7 days'
      GROUP BY DATE_TRUNC('day', created_at)
      ORDER BY date DESC
    `;

    // 获取系统日志统计
    const logStats = await sql`
      SELECT 
        level,
        COUNT(*) as count
      FROM system_logs
      WHERE created_at > NOW() - INTERVAL '24 hours'
      GROUP BY level
      ORDER BY count DESC
    `;

    return {
      summary: taskStats[0] || {},
      recentTasks: recentTasks || [],
      performanceStats: performanceStats || [],
      logStats: logStats || [],
      systemInfo: {
        version: process.env.REACT_APP_VERSION || '3.0.0',
        platform: 'Netlify Functions',
        nodeVersion: process.version,
        uptime: process.uptime(),
        memoryUsage: process.memoryUsage(),
        timestamp: new Date().toISOString()
      }
    };
  } catch (error) {
    console.error('获取系统统计失败:', error);
    throw error;
  }
}

// 获取健康状态
async function getHealthStatus() {
  try {
    await initDatabase();
    
    // 测试数据库连接
    const dbTest = await sql`SELECT 1 as test`;
    const dbHealthy = dbTest.length > 0;
    
    // 检查运行中的任务
    const runningTasks = await sql`
      SELECT COUNT(*) as count 
      FROM optimization_tasks 
      WHERE status = 'running'
    `;
    
    const runningCount = runningTasks[0]?.count || 0;
    
    return {
      status: 'healthy',
      database: dbHealthy ? 'connected' : 'disconnected',
      runningTasks: runningCount,
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    return {
      status: 'unhealthy',
      error: error.message,
      timestamp: new Date().toISOString()
    };
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
    if (event.httpMethod !== 'GET') {
      return {
        statusCode: 405,
        headers: {
          'Content-Type': 'application/json',
          'Access-Control-Allow-Origin': '*'
        },
        body: JSON.stringify({
          success: false,
          error: '仅支持GET请求'
        })
      };
    }

    const queryParams = event.queryStringParameters || {};
    const type = queryParams.type || 'full';

    let result;
    
    switch (type) {
      case 'health':
        result = await getHealthStatus();
        break;
      case 'full':
      default:
        result = await getSystemStats();
        break;
    }

    return {
      statusCode: 200,
      headers: {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*',
        'Cache-Control': 'no-cache, no-store, must-revalidate'
      },
      body: JSON.stringify({
        success: true,
        data: result,
        timestamp: new Date().toISOString()
      })
    };
  } catch (error) {
    console.error('❌ 系统统计API错误:', error);
    return {
      statusCode: 500,
      headers: {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*'
      },
      body: JSON.stringify({
        success: false,
        error: `获取系统统计失败: ${error.message}`
      })
    };
  }
}; 