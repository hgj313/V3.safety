/**
 * Netlify异步任务管理器 - Neon PostgreSQL版本
 * 最终修复版：提供所有必需方法，职责单一
 */
const { neon } = require('@neondatabase/serverless');

class TaskManager {
  constructor() {
    this.databaseUrl = process.env.NETLIFY_DATABASE_URL || process.env.DATABASE_URL;
    this.sql = null;
    this.isInitialized = false;
  }

  async initialize() {
    if (this.isInitialized) return;
    if (!this.databaseUrl) throw new Error('DATABASE_URL 环境变量未设置');
    this.sql = neon(this.databaseUrl);
    await this.ensureTablesExist();
    this.isInitialized = true;
    console.log('🔧 任务管理器初始化完成 (Neon PostgreSQL)');
  }

  async ensureTablesExist() {
    await this.sql`
      CREATE TABLE IF NOT EXISTS optimization_tasks (
        id TEXT PRIMARY KEY, type TEXT NOT NULL DEFAULT 'optimization',
        status TEXT NOT NULL CHECK (status IN ('pending', 'running', 'completed', 'failed', 'cancelled')),
        progress INTEGER DEFAULT 0, message TEXT, input_data JSONB, results JSONB,
        error_message TEXT, execution_time INTEGER,
        created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(), updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
      )`;
    await this.sql`CREATE INDEX IF NOT EXISTS idx_optimization_tasks_status ON optimization_tasks(status)`;
    await this.sql`CREATE INDEX IF NOT EXISTS idx_optimization_tasks_created_at ON optimization_tasks(created_at DESC)`;
  }

  generateTaskId() {
    return `task_${Date.now()}_${Math.floor(Math.random() * 900000) + 100000}`;
  }

  async createPendingTask(optimizationData) {
    await this.initialize();
    const taskId = this.generateTaskId();
    await this.sql`
      INSERT INTO optimization_tasks (id, type, status, progress, message, input_data) 
      VALUES (${taskId}, 'optimization', 'pending', 0, '任务已创建，等待后台处理', ${JSON.stringify(optimizationData)})`;
    console.log(`✅ 创建待处理任务: ${taskId}`);
    return taskId;
  }

  async getTask(taskId) {
    await this.initialize();
    const result = await this.sql`SELECT * FROM optimization_tasks WHERE id = ${taskId}`;
    if (result.length === 0) return null;
    const task = result[0];
    return {
      id: task.id, type: task.type, status: task.status, progress: task.progress, message: task.message,
      inputData: task.input_data, results: task.results, error: task.error_message,
      executionTime: task.execution_time, createdAt: task.created_at, updatedAt: task.updated_at
    };
  }

  async updateTaskStatus(taskId, status, updates = {}) {
    await this.initialize();
    const updateFields = { status, updated_at: new Date(), ...updates };
    if (updates.results) updateFields.results = JSON.stringify(updates.results);
    
    await this.sql`
      UPDATE optimization_tasks SET 
        status = ${updateFields.status},
        progress = COALESCE(${updateFields.progress}, progress),
        message = COALESCE(${updateFields.message}, message),
        results = COALESCE(${updateFields.results}, results),
        error_message = COALESCE(${updateFields.error}, error_message),
        execution_time = COALESCE(${updateFields.executionTime}, execution_time),
        updated_at = NOW()
      WHERE id = ${taskId}`;
    console.log(`📝 更新任务状态: ${taskId} -> ${status}`);
  }

  async updateTaskProgress(taskId, progress, message) {
    await this.updateTaskStatus(taskId, 'running', { progress, message });
  }

  async setTaskResults(taskId, results) {
    const task = await this.getTask(taskId);
    if (!task) throw new Error(`任务 ${taskId} 不存在`);
    const executionTime = new Date(task.createdAt).getTime();
    await this.updateTaskStatus(taskId, 'completed', {
      progress: 100, message: '优化完成', results,
      executionTime: Date.now() - executionTime
    });
  }

  async setTaskError(taskId, error) {
    await this.updateTaskStatus(taskId, 'failed', { error });
  }

  async getTaskList(options = {}) {
    await this.initialize();
    const { limit = 20, status = null } = options;
    let query;
    if (status) {
      query = this.sql`SELECT id, status, progress, message, execution_time, created_at FROM optimization_tasks WHERE status = ${status} ORDER BY created_at DESC LIMIT ${limit}`;
    } else {
      query = this.sql`SELECT id, status, progress, message, execution_time, created_at FROM optimization_tasks ORDER BY created_at DESC LIMIT ${limit}`;
    }
    return await query;
  }

  async cleanupExpiredTasks() {
    await this.initialize();
    const result = await this.sql`
      DELETE FROM optimization_tasks 
      WHERE created_at < NOW() - INTERVAL '24 hours' AND status IN ('completed', 'failed', 'cancelled')
      RETURNING id`;
    if (result.length > 0) {
      console.log(`🧹 清理了 ${result.length} 个过期任务。`);
    }
    return result.length;
  }
}

module.exports = TaskManager; 