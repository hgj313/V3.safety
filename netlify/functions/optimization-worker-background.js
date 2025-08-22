/**
 * Netlify后台函数 - 优化任务工作者
 * 负责执行耗时长的钢材优化计算
 */
const TaskManager = require('./utils/TaskManager');

// 使用精确的相对路径，确保在任何环境下都能可靠地加载模块
const OptimizationService = require('../../api/services/OptimizationService');

exports.handler = async (event, context) => {
  const taskManager = new TaskManager();
  
  try {
    const { taskId, optimizationData } = JSON.parse(event.body);
    console.log(`[${taskId}] 后台工作者已启动，开始执行优化...`);

    // 确保在异步执行上下文中初始化
    await taskManager.initialize();
    
    // 步骤 1: (已在主函数中完成) -> 更新任务为"运行中"
    await taskManager.updateTaskStatus(taskId, 'running', {
      progress: 10,
      message: '优化算法已启动...'
    });

    // 步骤 2: 正确地实例化优化服务
    const service = new OptimizationService();
    console.log(`[${taskId}] OptimizationService 实例化成功`);

    // 步骤 3: 定义进度回调
    const progressCallback = async (progress, message) => {
      const newProgress = Math.max(10, Math.round(progress));
      console.log(`[${taskId}] 进度更新: ${newProgress}% - ${message}`);
      await taskManager.updateTaskProgress(taskId, newProgress, message);
    };
    
    // 步骤 4: 运行优化算法
    console.log(`[${taskId}] 调用 service.run()...`);
    const startTime = Date.now();
    const results = await service.run(optimizationData, progressCallback);
    const executionTime = Date.now() - startTime;
    console.log(`[${taskId}] service.run() 完成，耗时: ${executionTime}ms`);
    
    // 步骤 5: 设置最终结果
    await taskManager.setTaskResults(taskId, {
      ...results,
      executionTime: `${(executionTime / 1000).toFixed(2)}s`
    });

    console.log(`[${taskId}] 任务成功完成`);

  } catch (error) {
    const taskId = JSON.parse(event.body)?.taskId || 'unknown_task';
    const errorMessage = `后台工作者捕获到致命错误: ${error.message}. Stack: ${error.stack}`;
    console.error(`[${taskId}] 优化任务执行失败:`, error);
    
    try {
      await taskManager.setTaskError(taskId, errorMessage);
    } catch (dbError) {
      console.error(`[${taskId}] 在捕获到执行错误后，更新数据库也失败了:`, dbError);
    }
    
    // 即使失败，也需要成功返回，避免Netlify重试
    return {
      statusCode: 200,
      body: JSON.stringify({ success: false, error: errorMessage })
    };
  }
  
  return {
    statusCode: 200,
    body: JSON.stringify({ success: true })
  };
}; 