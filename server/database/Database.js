const { Low } = require('lowdb');
const { JSONFile } = require('lowdb/node');
const fs = require('fs');
const path = require('path');

/**
 * JSON数据库管理类
 * 使用lowdb存储钢材优化系统数据
 * 优点：无需编译、跨平台、易于调试
 */
class DatabaseManager {
  constructor() {
    this.db = null;
    this.dbPath = path.join(__dirname, 'steel_system.json');
    this.backupDir = path.join(__dirname, 'backups');
  }

  /**
   * 初始化数据库连接
   */
  async init() {
    try {
      console.log('🗄️ 正在初始化JSON数据库...');
      
      // 确保备份目录存在
      if (!fs.existsSync(this.backupDir)) {
        fs.mkdirSync(this.backupDir, { recursive: true });
      }
      
      // 创建数据库适配器
      const adapter = new JSONFile(this.dbPath);
      this.db = new Low(adapter, this.getDefaultData());
      
      // 读取数据库文件
      await this.db.read();

      // 如果文件不存在（首次运行或被删除），立即写入默认数据到磁盘
      if (!fs.existsSync(this.dbPath)) {
        await this.save();
      }
      
      console.log('✅ JSON数据库初始化成功');
      console.log(`📍 数据库文件位置: ${this.dbPath}`);
      
      return true;
    } catch (error) {
      console.error('❌ 数据库初始化失败:', error);
      return false;
    }
  }

  /**
   * 获取默认数据结构
   */
  getDefaultData() {
    return {
      designSteels: [],
      moduleSteels: [
        {
          id: 'default_1',
          name: '12米标准钢材',
          length: 12000,
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString()
        },
        {
          id: 'default_2', 
          name: '9米标准钢材',
          length: 9000,
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString()
        },
        {
          id: 'default_3',
          name: '6米标准钢材', 
          length: 6000,
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString()
        }
      ],
      optimizationTasks: [],
      systemStats: {
        totalOptimizations: 0,
        totalDesignSteels: 0,
        totalModuleSteels: 3,
        totalSavedCost: 0,
        lastUpdated: new Date().toISOString()
      },
      operationLogs: [],
      settings: {
        autoBackup: true,
        maxLogEntries: 1000,
        maxBackups: 10
      }
    };
  }

  /**
   * 获取数据库连接
   */
  getConnection() {
    if (!this.db) {
      throw new Error('数据库未初始化，请先调用 init() 方法');
    }
    return this.db;
  }

  /**
   * 保存数据到文件
   */
  async save() {
    try {
      await this.db.write();
      return true;
    } catch (error) {
      console.error('保存数据失败:', error);
      return false;
    }
  }

  /**
   * 重新加载数据
   */
  async reload() {
    try {
      await this.db.read();
      return true;
    } catch (error) {
      console.error('重新加载数据失败:', error);
      return false;
    }
  }

  /**
   * 获取数据库统计信息
   */
  getStats() {
    try {
      const data = this.db.data;
      
      const stats = {
        designSteels: data.designSteels?.length || 0,
        moduleSteels: data.moduleSteels?.length || 0,
        optimizationTasks: data.optimizationTasks?.length || 0,
        completedTasks: data.optimizationTasks?.filter(task => task.status === 'completed')?.length || 0,
        operationLogs: data.operationLogs?.length || 0,
        databaseSize: this.getDatabaseSize(),
        lastUpdated: data.systemStats?.lastUpdated || new Date().toISOString()
      };
      
      return stats;
    } catch (error) {
      console.error('获取数据库统计信息失败:', error);
      return null;
    }
  }

  /**
   * 获取数据库文件大小
   */
  getDatabaseSize() {
    try {
      const stats = fs.statSync(this.dbPath);
      const sizeInBytes = stats.size;
      const sizeInKB = (sizeInBytes / 1024).toFixed(2);
      return `${sizeInKB} KB`;
    } catch (error) {
      return '未知';
    }
  }

  /**
   * 备份数据库
   */
  async backup(backupPath) {
    try {
      if (!backupPath) {
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        backupPath = path.join(this.backupDir, `backup_${timestamp}.json`);
      }
      
      // 确保数据是最新的
      await this.db.read();
      
      // 复制文件
      fs.copyFileSync(this.dbPath, backupPath);
      
      console.log(`✅ 数据库备份成功: ${backupPath}`);
      
      // 清理旧备份（保留最新10个）
      this.cleanOldBackups();
      
      return backupPath;
    } catch (error) {
      console.error('❌ 数据库备份失败:', error);
      throw error;
    }
  }

  /**
   * 清理旧备份文件
   */
  cleanOldBackups() {
    try {
      const files = fs.readdirSync(this.backupDir)
        .filter(file => file.startsWith('backup_') && file.endsWith('.json'))
        .map(file => ({
          name: file,
          path: path.join(this.backupDir, file),
          time: fs.statSync(path.join(this.backupDir, file)).mtime
        }))
        .sort((a, b) => b.time - a.time);

      // 保留最新的10个备份
      const maxBackups = this.db.data.settings?.maxBackups || 10;
      if (files.length > maxBackups) {
        const filesToDelete = files.slice(maxBackups);
        filesToDelete.forEach(file => {
          fs.unlinkSync(file.path);
          console.log(`🗑️ 删除旧备份: ${file.name}`);
        });
      }
    } catch (error) {
      console.error('清理旧备份失败:', error);
    }
  }

  /**
   * 记录操作日志
   */
  async logOperation(type, description, details = null, req = null) {
    try {
      const log = {
        id: Date.now().toString(),
        operationType: type,
        description: description,
        details: details,
        ipAddress: req ? (req.ip || req.connection.remoteAddress) : null,
        userAgent: req ? req.get('User-Agent') : null,
        createdAt: new Date().toISOString()
      };

      this.db.data.operationLogs.push(log);
      
      // 限制日志数量
      const maxLogs = this.db.data.settings?.maxLogEntries || 1000;
      if (this.db.data.operationLogs.length > maxLogs) {
        this.db.data.operationLogs = this.db.data.operationLogs.slice(-maxLogs);
      }
      
      await this.save();
    } catch (error) {
      console.error('记录操作日志失败:', error);
    }
  }

  /**
   * 更新系统统计
   */
  async updateSystemStats(updates) {
    try {
      this.db.data.systemStats = {
        ...this.db.data.systemStats,
        ...updates,
        lastUpdated: new Date().toISOString()
      };
      await this.save();
    } catch (error) {
      console.error('更新系统统计失败:', error);
    }
  }

  /**
   * 设计钢材CRUD操作
   */
  async saveDesignSteels(steels) {
    try {
      this.db.data.designSteels = steels.map(steel => ({
        ...steel,
        updatedAt: new Date().toISOString(),
        createdAt: steel.createdAt || new Date().toISOString()
      }));
      
      await this.updateSystemStats({
        totalDesignSteels: steels.length
      });
      
      await this.save();
      return true;
    } catch (error) {
      console.error('保存设计钢材失败:', error);
      return false;
    }
  }

  getDesignSteels() {
    return this.db.data.designSteels || [];
  }

  /**
   * 模数钢材CRUD操作
   */
  async saveModuleSteels(steels) {
    try {
      this.db.data.moduleSteels = steels.map(steel => ({
        ...steel,
        updatedAt: new Date().toISOString(),
        createdAt: steel.createdAt || new Date().toISOString()
      }));
      
      await this.updateSystemStats({
        totalModuleSteels: steels.length
      });
      
      await this.save();
      return true;
    } catch (error) {
      console.error('保存模数钢材失败:', error);
      return false;
    }
  }

  getModuleSteels() {
    return this.db.data.moduleSteels || [];
  }

  /**
   * 优化任务CRUD操作
   */
  async saveOptimizationTask(task) {
    try {
      const existingIndex = this.db.data.optimizationTasks.findIndex(t => t.id === task.id);
      
      const taskWithTimestamp = {
        ...task,
        updatedAt: new Date().toISOString(),
        createdAt: task.createdAt || new Date().toISOString()
      };
      
      if (existingIndex >= 0) {
        this.db.data.optimizationTasks[existingIndex] = taskWithTimestamp;
      } else {
        this.db.data.optimizationTasks.push(taskWithTimestamp);
      }
      
      // 如果是完成状态，更新总优化次数
      if (task.status === 'completed') {
        const completedCount = this.db.data.optimizationTasks.filter(t => t.status === 'completed').length;
        await this.updateSystemStats({
          totalOptimizations: completedCount
        });
      }
      
      await this.save();
      return true;
    } catch (error) {
      console.error('保存优化任务失败:', error);
      return false;
    }
  }

  getOptimizationTasks() {
    return this.db.data.optimizationTasks || [];
  }

  getOptimizationTask(id) {
    return this.db.data.optimizationTasks?.find(task => task.id === id) || null;
  }

  /**
   * 创建新的优化任务
   */
  async createOptimizationTask(taskData) {
    try {
      const taskId = `task_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
      
      const newTask = {
        id: taskId,
        status: 'pending',
        progress: 0,
        message: '任务已创建，等待执行',
        results: null,
        error: null,
        executionTime: 0,
        startTime: new Date().toISOString(),
        endTime: null,
        inputData: {
          designSteelsCount: taskData.designSteels?.length || 0,
          moduleSteelsCount: taskData.moduleSteels?.length || 0,
          constraints: taskData.constraints || {}
        },
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      };
      
      this.db.data.optimizationTasks.push(newTask);
      await this.save();
      
      console.log(`✅ 创建优化任务: ${taskId}`);
      return taskId;
    } catch (error) {
      console.error('创建优化任务失败:', error);
      throw error;
    }
  }

  /**
   * 更新任务状态
   */
  async updateTaskStatus(taskId, status, updates = {}) {
    try {
      const taskIndex = this.db.data.optimizationTasks.findIndex(t => t.id === taskId);
      
      if (taskIndex === -1) {
        throw new Error(`任务不存在: ${taskId}`);
      }
      
      const task = this.db.data.optimizationTasks[taskIndex];
      
      // 更新任务数据
      this.db.data.optimizationTasks[taskIndex] = {
        ...task,
        status: status,
        updatedAt: new Date().toISOString(),
        ...updates
      };
      
      // 如果任务完成或失败，设置结束时间
      if (status === 'completed' || status === 'failed') {
        this.db.data.optimizationTasks[taskIndex].endTime = new Date().toISOString();
        
        // 计算执行时间
        const startTime = new Date(task.startTime).getTime();
        const endTime = Date.now();
        this.db.data.optimizationTasks[taskIndex].executionTime = endTime - startTime;
      }
      
      await this.save();
      
      console.log(`📝 任务状态更新: ${taskId} -> ${status}`);
      return true;
    } catch (error) {
      console.error('更新任务状态失败:', error);
      return false;
    }
  }

  /**
   * 更新任务进度
   */
  async updateTaskProgress(taskId, progress, message = '') {
    try {
      const taskIndex = this.db.data.optimizationTasks.findIndex(t => t.id === taskId);
      
      if (taskIndex === -1) {
        console.warn(`任务不存在: ${taskId}`);
        return false;
      }
      
      this.db.data.optimizationTasks[taskIndex].progress = progress;
      this.db.data.optimizationTasks[taskIndex].message = message;
      this.db.data.optimizationTasks[taskIndex].updatedAt = new Date().toISOString();
      
      await this.save();
      
      console.log(`📊 任务进度更新: ${taskId} -> ${progress}% (${message})`);
      return true;
    } catch (error) {
      console.error('更新任务进度失败:', error);
      return false;
    }
  }

  /**
   * 设置任务结果
   */
  async setTaskResults(taskId, results) {
    try {
      return await this.updateTaskStatus(taskId, 'completed', {
        results: JSON.stringify(results),
        progress: 100,
        message: '优化完成'
      });
    } catch (error) {
      console.error('设置任务结果失败:', error);
      return false;
    }
  }

  /**
   * 设置任务错误
   */
  async setTaskError(taskId, error) {
    try {
      return await this.updateTaskStatus(taskId, 'failed', {
        error: error.message || String(error),
        message: `优化失败: ${error.message || String(error)}`
      });
    } catch (error) {
      console.error('设置任务错误失败:', error);
      return false;
    }
  }

  /**
   * 获取活跃任务（正在执行的任务）
   */
  getActiveTasks() {
    return this.db.data.optimizationTasks?.filter(task => 
      task.status === 'pending' || task.status === 'running'
    ) || [];
  }

  /**
   * 清理过期任务（超过24小时的已完成任务）
   */
  async cleanupExpiredTasks() {
    try {
      const oneDayAgo = new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString();
      
      const beforeCount = this.db.data.optimizationTasks.length;
      this.db.data.optimizationTasks = this.db.data.optimizationTasks.filter(task => {
        // 保留活跃任务和24小时内的任务
        return (task.status === 'pending' || task.status === 'running') || 
               (task.updatedAt > oneDayAgo);
      });
      
      const afterCount = this.db.data.optimizationTasks.length;
      const cleanedCount = beforeCount - afterCount;
      
      if (cleanedCount > 0) {
        await this.save();
        console.log(`🧹 清理了 ${cleanedCount} 个过期任务`);
      }
      
      return cleanedCount;
    } catch (error) {
      console.error('清理过期任务失败:', error);
      return 0;
    }
  }

  /**
   * 获取操作日志
   */
  getOperationLogs(limit = 100) {
    const logs = this.db.data.operationLogs || [];
    return logs.slice(-limit).reverse(); // 返回最新的日志
  }

  /**
   * 导出数据
   */
  async exportData() {
    try {
      await this.db.read();
      return {
        exportTime: new Date().toISOString(),
        version: '3.0.0',
        data: this.db.data
      };
    } catch (error) {
      console.error('导出数据失败:', error);
      return null;
    }
  }

  /**
   * 导入数据
   */
  async importData(data) {
    try {
      // 备份当前数据
      await this.backup();
      
      // 验证数据格式
      if (!data.data || typeof data.data !== 'object') {
        throw new Error('无效的数据格式');
      }
      
      // 合并数据
      this.db.data = {
        ...this.getDefaultData(),
        ...data.data
      };
      
      await this.save();
      console.log('✅ 数据导入成功');
      return true;
    } catch (error) {
      console.error('❌ 数据导入失败:', error);
      return false;
    }
  }
}

// 创建单例实例
const databaseManager = new DatabaseManager();

module.exports = databaseManager; 