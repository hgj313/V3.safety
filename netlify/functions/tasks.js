/**
 * Netlify Function - 获取任务列表
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
    const taskManager = new TaskManager();
    const tasks = await taskManager.getTaskList(event.queryStringParameters);
    
    return {
      statusCode: 200,
      headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' },
      body: JSON.stringify({ success: true, tasks })
    };

  } catch (error) {
    console.error('❌ 获取任务列表失败:', error);
    return {
      statusCode: 500,
      headers: { 'Content-T ype': 'application/json', 'Access-Control-Allow-Origin': '*' },
      body: JSON.stringify({ success: false, error: `服务器内部错误: ${error.message}` })
    };
  }
}; 