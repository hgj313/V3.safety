/**
 * V3并行计算性能监控器
 * 跟踪并行优化任务的执行状态和性能指标
 */

class ParallelOptimizationMonitor {
  constructor() {
    this.tasks = new Map(); // 任务状态跟踪
    this.globalStats = {
      totalTasks: 0,
      completedTasks: 0,
      failedTasks: 0,
      totalExecutionTime: 0,
      averageExecutionTime: 0,
      maxExecutionTime: 0,
      minExecutionTime: Infinity,
      parallelEfficiency: 0
    };
    this.startTime = null;
    this.endTime = null;
  }

  /**
   * 开始监控并行计算
   */
  startMonitoring(taskCount) {
    this.startTime = Date.now();
    this.globalStats.totalTasks = taskCount;
    
    console.log(`📊 V3并行计算监控启动: ${taskCount}个任务将并行执行`);
    
    // 初始化任务状态
    for (let i = 0; i < taskCount; i++) {
      this.tasks.set(`task_${i}`, {
        index: i,
        status: 'pending',
        startTime: null,
        endTime: null,
        executionTime: 0,
        groupKey: null,
        steelsCount: 0,
        cuts: 0,
        error: null
      });
    }
  }

  /**
   * 记录任务开始
   */
  recordTaskStart(taskIndex, groupKey, steelsCount) {
    const task = this.tasks.get(`task_${taskIndex}`);
    if (task) {
      task.status = 'running';
      task.startTime = Date.now();
      task.groupKey = groupKey;
      task.steelsCount = steelsCount;
      
      console.log(`🚀 任务${taskIndex}开始: ${groupKey} (${steelsCount}种钢材)`);
    }
  }

  /**
   * 记录任务完成
   */
  recordTaskCompletion(taskIndex, stats) {
    const task = this.tasks.get(`task_${taskIndex}`);
    if (task) {
      task.status = 'completed';
      task.endTime = Date.now();
      task.executionTime = task.endTime - task.startTime;
      task.cuts = stats.cuts;
      
      this.globalStats.completedTasks++;
      this.updateGlobalStats(task.executionTime);
      
      console.log(`✅ 任务${taskIndex}完成: ${task.groupKey} (${task.executionTime}ms, ${stats.cuts}次切割)`);
    }
  }

  /**
   * 记录任务失败
   */
  recordTaskFailure(taskIndex, error) {
    const task = this.tasks.get(`task_${taskIndex}`);
    if (task) {
      task.status = 'failed';
      task.endTime = Date.now();
      task.executionTime = task.endTime - task.startTime;
      task.error = error;
      
      this.globalStats.failedTasks++;
      
      console.error(`❌ 任务${taskIndex}失败: ${task.groupKey} - ${error}`);
    }
  }

  /**
   * 更新全局统计
   */
  updateGlobalStats(executionTime) {
    this.globalStats.totalExecutionTime += executionTime;
    this.globalStats.averageExecutionTime = 
      this.globalStats.totalExecutionTime / this.globalStats.completedTasks;
    
    if (executionTime > this.globalStats.maxExecutionTime) {
      this.globalStats.maxExecutionTime = executionTime;
    }
    
    if (executionTime < this.globalStats.minExecutionTime) {
      this.globalStats.minExecutionTime = executionTime;
    }
  }

  /**
   * 完成监控
   */
  finishMonitoring() {
    this.endTime = Date.now();
    const totalWallTime = this.endTime - this.startTime;
    
    // 计算并行效率
    const sequentialTime = this.globalStats.totalExecutionTime;
    this.globalStats.parallelEfficiency = 
      sequentialTime > 0 ? (sequentialTime / totalWallTime) * 100 : 0;
    
    console.log(`🏁 V3并行计算监控完成:`);
    console.log(`   总墙钟时间: ${totalWallTime}ms`);
    console.log(`   总CPU时间: ${sequentialTime}ms`);
    console.log(`   并行效率: ${this.globalStats.parallelEfficiency.toFixed(2)}%`);
    console.log(`   成功任务: ${this.globalStats.completedTasks}/${this.globalStats.totalTasks}`);
    console.log(`   失败任务: ${this.globalStats.failedTasks}/${this.globalStats.totalTasks}`);
    
    return this.generateReport();
  }

  /**
   * 生成详细报告
   */
  generateReport() {
    const report = {
      summary: {
        totalTasks: this.globalStats.totalTasks,
        successfulTasks: this.globalStats.completedTasks,
        failedTasks: this.globalStats.failedTasks,
        successRate: (this.globalStats.completedTasks / this.globalStats.totalTasks) * 100,
        totalWallTime: this.endTime - this.startTime,
        totalCpuTime: this.globalStats.totalExecutionTime,
        parallelEfficiency: this.globalStats.parallelEfficiency,
        averageTaskTime: this.globalStats.averageExecutionTime,
        maxTaskTime: this.globalStats.maxExecutionTime,
        minTaskTime: this.globalStats.minExecutionTime === Infinity ? 0 : this.globalStats.minExecutionTime
      },
      taskDetails: [],
      performance: {
        speedup: this.calculateSpeedup(),
        efficiency: this.calculateEfficiency(),
        scalability: this.calculateScalability()
      }
    };

    // 添加任务详情
    this.tasks.forEach((task, taskId) => {
      report.taskDetails.push({
        taskId: taskId,
        groupKey: task.groupKey,
        status: task.status,
        executionTime: task.executionTime,
        steelsCount: task.steelsCount,
        cuts: task.cuts,
        error: task.error
      });
    });

    // 按执行时间排序
    report.taskDetails.sort((a, b) => b.executionTime - a.executionTime);

    return report;
  }

  /**
   * 计算加速比
   */
  calculateSpeedup() {
    const wallTime = this.endTime - this.startTime;
    const cpuTime = this.globalStats.totalExecutionTime;
    return cpuTime > 0 ? cpuTime / wallTime : 1;
  }

  /**
   * 计算效率
   */
  calculateEfficiency() {
    const speedup = this.calculateSpeedup();
    return speedup / this.globalStats.totalTasks;
  }

  /**
   * 计算可扩展性指标
   */
  calculateScalability() {
    const efficiency = this.calculateEfficiency();
    return {
      efficiency: efficiency,
      rating: efficiency > 0.8 ? 'excellent' : 
              efficiency > 0.6 ? 'good' : 
              efficiency > 0.4 ? 'fair' : 'poor'
    };
  }

  /**
   * 获取实时状态
   */
  getRealtimeStatus() {
    const runningTasks = Array.from(this.tasks.values()).filter(t => t.status === 'running');
    const completedTasks = Array.from(this.tasks.values()).filter(t => t.status === 'completed');
    const failedTasks = Array.from(this.tasks.values()).filter(t => t.status === 'failed');
    
    return {
      total: this.globalStats.totalTasks,
      running: runningTasks.length,
      completed: completedTasks.length,
      failed: failedTasks.length,
      progress: ((completedTasks.length + failedTasks.length) / this.globalStats.totalTasks) * 100
    };
  }

  /**
   * 输出性能报告到控制台
   */
  printPerformanceReport() {
    const report = this.generateReport();
    
    console.log('\n📊 ==================== V3并行计算性能报告 ====================');
    console.log(`🎯 总体表现:`);
    console.log(`   任务总数: ${report.summary.totalTasks}`);
    console.log(`   成功率: ${report.summary.successRate.toFixed(2)}%`);
    console.log(`   总墙钟时间: ${report.summary.totalWallTime}ms`);
    console.log(`   并行效率: ${report.summary.parallelEfficiency.toFixed(2)}%`);
    console.log(`   加速比: ${report.performance.speedup.toFixed(2)}x`);
    console.log(`   效率评级: ${report.performance.scalability.rating}`);
    
    console.log(`\n⏱️ 任务时间分析:`);
    console.log(`   平均执行时间: ${report.summary.averageTaskTime.toFixed(2)}ms`);
    console.log(`   最长执行时间: ${report.summary.maxTaskTime}ms`);
    console.log(`   最短执行时间: ${report.summary.minTaskTime}ms`);
    
    console.log(`\n🏆 性能最佳任务:`);
    const bestTasks = report.taskDetails
      .filter(t => t.status === 'completed')
      .slice(0, 3);
    bestTasks.forEach((task, index) => {
      console.log(`   ${index + 1}. ${task.groupKey}: ${task.executionTime}ms (${task.cuts}次切割)`);
    });
    
    if (report.summary.failedTasks > 0) {
      console.log(`\n❌ 失败任务:`);
      const failedTasks = report.taskDetails.filter(t => t.status === 'failed');
      failedTasks.forEach(task => {
        console.log(`   ${task.groupKey}: ${task.error}`);
      });
    }
    
    console.log('============================================================\n');
  }
}

module.exports = ParallelOptimizationMonitor; 