/**
 * Netlify Function - 获取单个任务状态
 */
const TaskManager = require('./utils/TaskManager');

exports.handler = async (event, context) => {
  // 预检请求处理
  if (event.httpMethod === 'OPTIONS') {
    return {
      statusCode: 200,
      headers: { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Headers': 'Content-Type', 'Access-Control-Allow-Methods': 'GET' },
      body: ''
    };
  }

  try {
    const taskId = event.path.split('/').pop();
    if (!taskId) {
      return { statusCode: 400, body: JSON.stringify({ success: false, error: '缺少任务ID' }) };
    }
    
    const taskManager = new TaskManager();
    
    // 非阻塞地触发清理任务，不影响主流程
    taskManager.cleanupExpiredTasks().catch(err => console.error('清理过期任务时发生错误:', err));
    
    // 获取当前任务状态
    const task = await taskManager.getTask(taskId);
    
    if (!task) {
      return { statusCode: 404, body: JSON.stringify({ success: false, error: '任务不存在' }) };
    }
    
    // 关键修复：返回一个扁平化的对象，而不是嵌套的task对象，以匹配前端的期望
    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' },
      body: JSON.stringify({
        success: true,
        id: task.id,
        status: task.status,
        progress: task.progress,
        message: task.message,
        results: task.results,
        error: task.error,
        executionTime: task.executionTime,
        createdAt: task.createdAt,
        updatedAt: task.updatedAt
      })
    };

  } catch (error) {
    console.error('❌ 获取任务状态失败:', error);
    return {
      statusCode: 500,
      headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' },
      body: JSON.stringify({ success: false, error: `服务器内部错误: ${error.message}` })
    };
  }
}; 