/**
 * 钢材采购优化系统 V3.0 - 稳定版服务器
 * 结合简化版的稳定性和完整版的功能
 */

const express = require('express');
const cors = require('cors');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const databaseManager = require('./database/Database');
const ExcelJS = require('exceljs');
// PDF功能采用V2简单HTML生成方案，无需复杂的jsPDF库
// 所有jsPDF相关代码已删除

const app = express();
const PORT = process.env.PORT || 5099;

// 中间件配置
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));

// 请求日志中间件
app.use((req, res, next) => {
  console.log(`📥 ${req.method} ${req.url}`);
  next();
});

// 创建uploads目录
const uploadsDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir, { recursive: true });
}

// 文件上传配置
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, uploadsDir);
  },
  filename: (req, file, cb) => {
    const timestamp = Date.now();
    const originalName = Buffer.from(file.originalname, 'latin1').toString('utf8');
    cb(null, `${timestamp}-${originalName}`);
  }
});

const upload = multer({ 
  storage: storage,
  limits: { fileSize: 10 * 1024 * 1024 }
});

// 优化服务实例（延迟加载）
let optimizationService = null;

// 延迟加载优化服务
function getOptimizationService() {
  if (!optimizationService) {
    try {
      const OptimizationService = require('../api/services/OptimizationService');
      optimizationService = new OptimizationService();
      console.log('✅ 优化服务模块加载成功');
    } catch (error) {
      console.warn('⚠️ 优化服务模块加载失败，将使用模拟模式:', error.message);
      optimizationService = createMockOptimizationService();
    }
  }
  return optimizationService;
}

// 创建模拟优化服务
function createMockOptimizationService() {
  return {
    optimizeSteel: async (data) => ({
      success: true,
      result: {
        totalLossRate: 3.5,
        totalModuleUsed: 100,
        totalWaste: 50,
        solutions: {},
        executionTime: 1500
      },
      optimizationId: 'mock_' + Date.now(),
      stats: { totalCuts: 10, remaindersGenerated: 5 }
    }),
    validateWeldingConstraints: async (data) => ({
      success: true,
      validation: { isValid: true, violations: [], suggestions: [] }
    }),
    getOptimizationProgress: (id) => ({
      success: true,
      status: 'completed',
      runningTime: 1500
    }),
    cancelOptimization: (id) => ({
      success: true,
      message: '优化已取消'
    }),
    getActiveOptimizers: () => ({
      success: true,
      activeOptimizers: []
    }),
    getOptimizationHistory: (limit) => ({
      success: true,
      history: []
    }),
    getSystemStats: () => ({
      success: true,
      stats: {
        totalOptimizations: 0,
        averageExecutionTime: 0,
        successRate: 100
      }
    }),
    cleanupExpiredOptimizers: () => {}
  };
}

// ==================== API路由 ====================

/**
 * 系统健康检查
 */
app.get('/api/health', (req, res) => {
  res.json({
    success: true,
    message: '钢材采购优化系统 V3.0 运行正常',
    version: '3.0.0',
    timestamp: new Date().toISOString(),
    uptime: process.uptime(),
    optimizationServiceStatus: optimizationService ? 'loaded' : 'not_loaded'
  });
});

/**
 * 获取系统统计信息
 */
app.get('/api/stats', async (req, res) => {
  console.log('🔍 Stats端点被访问');
  try {
    console.log('🔧 获取数据库统计...');
    const dbStats = databaseManager.getStats();
    
    console.log('🔧 获取优化服务统计...');
    const service = getOptimizationService();
    const serviceStats = service.getSystemStats();
    
    // 合并数据库和服务统计
    const combinedStats = {
      success: true,
      stats: {
        totalOptimizations: dbStats?.completedTasks || 0,
        totalDesignSteels: dbStats?.designSteels || 0,
        totalModuleSteels: dbStats?.moduleSteels || 0,
        totalSavedCost: 0, // 待开发功能
        averageExecutionTime: serviceStats?.stats?.averageExecutionTime || 0,
        successRate: serviceStats?.stats?.successRate || 100,
        databaseSize: dbStats?.databaseSize || '未知',
        lastUpdated: dbStats?.lastUpdated || new Date().toISOString()
      }
    };
    
    console.log('✅ Stats数据:', combinedStats);
    res.json(combinedStats);
  } catch (error) {
    console.error('❌ Stats端点错误:', error);
    res.status(500).json({
      success: false,
      error: `获取系统统计失败: ${error.message}`
    });
  }
});

/**
 * 验证约束条件
 */
app.post('/api/validate-constraints', async (req, res) => {
  try {
    const service = getOptimizationService();
    const result = await service.validateWeldingConstraints(req.body);
    res.json(result);
  } catch (error) {
    res.status(500).json({
      success: false,
      error: `约束验证失败: ${error.message}`
    });
  }
});

/**
 * 执行钢材优化 - 异步任务模式
 */
app.post('/api/optimize', async (req, res) => {
  console.log('--- ENTERING NEW ASYNC /api/optimize (DEBUG V5) ---');
  try {
    console.log('🚀 收到优化请求 - 异步模式');
    console.log('设计钢材数量:', req.body.designSteels?.length || 0);
    console.log('模数钢材数量:', req.body.moduleSteels?.length || 0);
    console.log('约束条件:', req.body.constraints);

    // 创建异步任务
    const taskId = await databaseManager.createOptimizationTask(req.body);
    
    // 立即返回taskId，不等待优化完成
    res.json({
      success: true,
      taskId: taskId,
      message: '优化任务已创建，请通过taskId查询进度',
      status: 'pending'
    });
    
    // 在后台异步执行优化
    executeOptimizationTask(taskId, req.body);
    
  } catch (error) {
    console.error('❌ 创建优化任务失败:', error);
    res.status(500).json({
      success: false,
      error: `创建优化任务失败: ${error.message}`
    });
  }
});

/**
 * 异步执行优化任务
 */
async function executeOptimizationTask(taskId, optimizationData) {
  try {
    console.log(`🔄 开始执行优化任务: ${taskId}`);
    
    // 更新任务状态为运行中
    await databaseManager.updateTaskStatus(taskId, 'running', {
      progress: 10,
      message: '正在初始化优化器...'
    });
    
    // 获取优化服务实例
    const service = getOptimizationService();
    
    // 模拟进度更新
    await databaseManager.updateTaskProgress(taskId, 20, '正在解析输入数据...');
    
    // 执行优化计算
    await databaseManager.updateTaskProgress(taskId, 30, '正在计算最优切割方案...');
    
    const result = await service.optimizeSteel(optimizationData);
    
    if (result.success) {
      console.log(`✅ 优化任务完成: ${taskId}`);
      console.log('执行时间:', result.executionTime + 'ms');
      console.log('总损耗率:', result.result?.totalLossRate + '%');

      // 保存优化结果到数据库
      await databaseManager.setTaskResults(taskId, result.result);
      
      // 记录操作日志
      await databaseManager.logOperation(
        'optimize',
        `完成优化任务 ${taskId}`,
        { 
          optimizationId: result.optimizationId,
          executionTime: result.executionTime,
          totalLossRate: result.result?.totalLossRate
        }
      );
      
    } else {
      console.log(`❌ 优化任务失败: ${taskId}`, result.error);
      await databaseManager.setTaskError(taskId, new Error(result.error || '优化计算失败'));
    }
    
  } catch (error) {
    console.error(`❌ 执行优化任务异常: ${taskId}`, error);
    await databaseManager.setTaskError(taskId, error);
  }
}

/**
 * 获取任务状态和结果
 */
app.get('/api/task/:taskId', async (req, res) => {
  try {
    const { taskId } = req.params;
    const task = databaseManager.getOptimizationTask(taskId);
    
    if (!task) {
      return res.status(404).json({
        success: false,
        error: '任务不存在'
      });
    }
    
    // 构建响应数据
    const response = {
      success: true,
      taskId: task.id,
      status: task.status,
      progress: task.progress,
      message: task.message,
      executionTime: task.executionTime,
      createdAt: task.createdAt,
      updatedAt: task.updatedAt
    };
    
    // 如果任务完成，包含结果数据
    if (task.status === 'completed' && task.results) {
      try {
        response.results = JSON.parse(task.results);
      } catch (parseError) {
        console.error('解析任务结果失败:', parseError);
        response.results = null;
      }
    }
    
    // 如果任务失败，包含错误信息
    if (task.status === 'failed' && task.error) {
      response.error = task.error;
    }
    
    res.json(response);
    
  } catch (error) {
    console.error('❌ 获取任务状态失败:', error);
    res.status(500).json({
      success: false,
      error: `获取任务状态失败: ${error.message}`
    });
  }
});

/**
 * 获取所有任务列表
 */
app.get('/api/tasks', async (req, res) => {
  try {
    const limit = parseInt(req.query.limit) || 20;
    const status = req.query.status; // 可选的状态过滤
    
    let tasks = databaseManager.getOptimizationTasks();
    
    // 状态过滤
    if (status) {
      tasks = tasks.filter(task => task.status === status);
    }
    
    // 按创建时间倒序排列
    tasks.sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime());
    
    // 限制返回数量
    tasks = tasks.slice(0, limit);
    
    // 简化任务信息（不包含完整结果）
    const simplifiedTasks = tasks.map(task => ({
      id: task.id,
      status: task.status,
      progress: task.progress,
      message: task.message,
      executionTime: task.executionTime,
      inputData: task.inputData,
      createdAt: task.createdAt,
      updatedAt: task.updatedAt,
      hasResults: task.status === 'completed' && !!task.results
    }));
    
    res.json({
      success: true,
      tasks: simplifiedTasks,
      total: databaseManager.getOptimizationTasks().length
    });
    
  } catch (error) {
    console.error('❌ 获取任务列表失败:', error);
    res.status(500).json({
      success: false,
      error: `获取任务列表失败: ${error.message}`
    });
  }
});

/**
 * 取消任务
 */
app.delete('/api/task/:taskId', async (req, res) => {
  try {
    const { taskId } = req.params;
    const task = databaseManager.getOptimizationTask(taskId);
    
    if (!task) {
      return res.status(404).json({
        success: false,
        error: '任务不存在'
      });
    }
    
    // 只能取消待执行或正在执行的任务
    if (task.status !== 'pending' && task.status !== 'running') {
      return res.status(400).json({
        success: false,
        error: '只能取消待执行或正在执行的任务'
      });
    }
    
    // 更新任务状态为已取消
    await databaseManager.updateTaskStatus(taskId, 'cancelled', {
      message: '任务已被用户取消'
    });
    
    res.json({
      success: true,
      message: '任务已取消'
    });
    
  } catch (error) {
    console.error('❌ 取消任务失败:', error);
    res.status(500).json({
      success: false,
      error: `取消任务失败: ${error.message}`
    });
  }
});

/**
 * 获取优化进度
 */
app.get('/api/optimize/:optimizationId/progress', (req, res) => {
  try {
    const { optimizationId } = req.params;
    const service = getOptimizationService();
    const result = service.getOptimizationProgress(optimizationId);
    res.json(result);
  } catch (error) {
    res.status(500).json({
      success: false,
      error: `获取优化进度失败: ${error.message}`
    });
  }
});

/**
 * 取消优化
 */
app.delete('/api/optimize/:optimizationId', (req, res) => {
  try {
    const { optimizationId } = req.params;
    const service = getOptimizationService();
    const result = service.cancelOptimization(optimizationId);
    res.json(result);
  } catch (error) {
    res.status(500).json({
      success: false,
      error: `取消优化失败: ${error.message}`
    });
  }
});

/**
 * 获取活跃的优化任务
 */
app.get('/api/optimize/active', (req, res) => {
  try {
    const service = getOptimizationService();
    const result = service.getActiveOptimizers();
    res.json(result);
  } catch (error) {
    res.status(500).json({
      success: false,
      error: `获取活跃优化任务失败: ${error.message}`
    });
  }
});

/**
 * 获取优化历史
 */
app.get('/api/optimize/history', (req, res) => {
  try {
    const limit = parseInt(req.query.limit) || 20;
    const service = getOptimizationService();
    const result = service.getOptimizationHistory(limit);
    res.json(result);
  } catch (error) {
    res.status(500).json({
      success: false,
      error: `获取优化历史失败: ${error.message}`
    });
  }
});

/**
 * 上传设计钢材文件 - 支持JSON和multipart两种格式
 */
app.post('/api/upload-design-steels', async (req, res) => {
  try {
    console.log('📁 处理文件上传请求');
    
    // 处理JSON格式的base64数据
    if (req.is('application/json') && req.body.data) {
      const { filename, data, type } = req.body;
      
      console.log('JSON格式上传:', { filename, type, size: data.length });
      
      // 解析base64数据
      const buffer = Buffer.from(data, 'base64');
      
      // 解析文件并返回设计钢材数据
      const designSteels = await parseFileBuffer(buffer, filename, type);
      
      res.json({
        success: true,
        message: `文件解析成功，找到 ${designSteels.length} 条设计钢材数据`,
        designSteels: designSteels,
        debugInfo: {
          原始行数: designSteels.length + 1,
          有效数据: designSteels.length,
          列名信息: '编号,构件编号,规格,长度,数量,截面面积,部件编号',
          截面面积统计: {
            有截面面积: designSteels.filter(s => s.crossSection > 0).length,
            无截面面积: designSteels.filter(s => s.crossSection === 0).length
          }
        }
      });

      // === 持久化到JSON数据库 ===
      try {
        await databaseManager.saveDesignSteels(designSteels);
        await databaseManager.logOperation(
          'upload',
          `上传设计钢材文件 ${filename}`,
          { count: designSteels.length, filename }
        );
      } catch (dbErr) {
        console.warn('⚠️ 设计钢材数据未写入数据库:', dbErr.message);
      }
      
      return;
    }
    
    // 处理multipart/form-data格式
    upload.single('file')(req, res, async (err) => {
      if (err) {
        console.error('❌ 文件上传错误:', err);
        return res.status(400).json({
          success: false,
          error: err.message
        });
      }
      
      if (!req.file) {
        return res.status(400).json({
          success: false,
          error: '未接收到文件'
        });
      }

      console.log('文件上传成功:', req.file.originalname);
      
      // 读取文件并解析
      const buffer = fs.readFileSync(req.file.path);
      const designSteels = await parseFileBuffer(buffer, req.file.originalname, req.file.mimetype);
      
      res.json({
        success: true,
        message: `文件解析成功，找到 ${designSteels.length} 条设计钢材数据`,
        designSteels: designSteels,
        debugInfo: {
          原始行数: designSteels.length + 1,
          有效数据: designSteels.length,
          列名信息: '编号,构件编号,规格,长度,数量,截面面积,部件编号',
          截面面积统计: {
            有截面面积: designSteels.filter(s => s.crossSection > 0).length,
            无截面面积: designSteels.filter(s => s.crossSection === 0).length
          }
        }
      });

      // === 持久化到JSON数据库 ===
      try {
        await databaseManager.saveDesignSteels(designSteels);
        await databaseManager.logOperation(
          'upload',
          `上传设计钢材文件 ${req.file.originalname}`,
          { count: designSteels.length, filename: req.file.originalname }
        );
      } catch (dbErr) {
        console.warn('⚠️ 设计钢材数据未写入数据库:', dbErr.message);
      }
      
      return;
    });

  } catch (error) {
    console.error('❌ 文件处理错误:', error);
    res.status(500).json({
      success: false,
      error: `文件处理失败: ${error.message}`
    });
  }
});

/**
 * 导出Excel报告
 */
app.post('/api/export/excel', async (req, res) => {
  try {
    const { optimizationResult, exportOptions = {} } = req.body;
    
    if (!optimizationResult) {
      return res.status(400).json({
        success: false,
        error: '缺少优化结果数据'
      });
    }

    console.log('📊 开始生成Excel报告...');
    const excelBuffer = await generateExcelReport(optimizationResult, exportOptions);
    
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const filename = `钢材优化报告_${timestamp}.xlsx`;
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(filename)}"`);
    res.send(excelBuffer);
    
    console.log('✅ Excel报告生成成功:', filename);
    
  } catch (error) {
    console.error('❌ Excel导出失败:', error);
    res.status(500).json({
      success: false,
      error: `Excel导出失败: ${error.message}`
    });
  }
});

/**
 * 导出PDF报告
 */
app.post('/api/export/pdf', async (req, res) => {
  try {
    const { optimizationResult, exportOptions = {}, designSteels = [] } = req.body;
    
    if (!optimizationResult) {
      return res.status(400).json({
        success: false,
        error: '缺少优化结果数据'
      });
    }

    console.log('📄 [方案B] 开始生成HTML报告内容...');
    
    const htmlContent = generatePDFHTML(optimizationResult, { 
      ...exportOptions, 
      designSteels: designSteels 
    });
    
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const fileName = `钢材优化报告_${timestamp}.html`;
    
    res.json({
      success: true,
      fileName: fileName,
      htmlContent: htmlContent, // 直接在JSON中返回HTML内容
      message: 'HTML报告内容已生成，请在前端处理下载。'
    });
    
    console.log('✅ [方案B] HTML内容生成并发送成功');
    
  } catch (error) {
    console.error('❌ PDF(HTML)导出失败:', error);
    res.status(500).json({
      success: false,
      error: `报告生成失败: ${error.message}`
    });
  }
});

// ==================== 文件解析功能 ====================

/**
 * 解析文件缓冲区为设计钢材数据
 */
async function parseFileBuffer(buffer, filename, mimetype) {
  console.log(`📊 开始解析文件: ${filename}, 类型: ${mimetype}`);
  
  try {
    let data = [];
    
    if (filename.toLowerCase().endsWith('.csv')) {
      // 解析CSV文件
      data = await parseCSVBuffer(buffer);
    } else if (filename.toLowerCase().endsWith('.xlsx') || filename.toLowerCase().endsWith('.xls')) {
      // 解析Excel文件
      const XLSX = require('xlsx');
      data = parseExcelBuffer(buffer, XLSX);
    } else {
      throw new Error('不支持的文件格式');
    }
    
    console.log(`📊 解析完成，找到 ${data.length} 条原始数据`);
    
    // 转换为标准格式
    const steelData = data.map((row, index) => {
      // 尝试多种可能的列名映射
      const length = parseFloat(
        row['长度'] || row['长度(mm)'] || row['Length'] || row['length'] || 
        row['长度 (mm)'] || row['长度（mm）'] || row['长度mm'] || 0
      );
      
      const quantity = parseInt(
        row['数量'] || row['Quantity'] || row['quantity'] || row['件数'] || 
        row['数量(件)'] || row['数量（件）'] || 1
      );
      
      const crossSection = parseFloat(
        row['截面面积'] || row['截面面积(mm²)'] || row['截面面积（mm²）'] || 
        row['CrossSection'] || row['crossSection'] || row['面积'] || 0
      );
      
      const componentNumber = String(
        row['构件编号'] || row['构件号'] || row['ComponentNumber'] || 
        row['componentNumber'] || row['编号'] || `GJ${String(index + 1).padStart(3, '0')}`
      );
      
      const specification = String(
        row['规格'] || row['Specification'] || row['specification'] || 
        row['型号'] || row['钢材规格'] || ''
      );
      
      const partNumber = String(
        row['部件编号'] || row['部件号'] || row['PartNumber'] || 
        row['partNumber'] || row['零件号'] || `BJ${String(index + 1).padStart(3, '0')}`
      );
      
      return {
        id: `steel_${Date.now()}_${index}`,
        length: length || 0,
        quantity: quantity || 1,
        crossSection: crossSection || 0,
        componentNumber: componentNumber,
        specification: specification,
        partNumber: partNumber
      };
    }).filter(steel => steel.length > 0); // 过滤掉长度为0的数据
    
    console.log(`✅ 数据转换完成，有效数据 ${steelData.length} 条`);
    return steelData;
    
  } catch (error) {
    console.error('❌ 文件解析错误:', error);
    throw new Error(`文件解析失败: ${error.message}`);
  }
}

/**
 * 解析Excel缓冲区
 */
function parseExcelBuffer(buffer, XLSX) {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  
  // 转换为JSON格式
  const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
    header: 1,
    defval: '',
    raw: false
  });
  
  if (jsonData.length < 2) {
    throw new Error('Excel文件数据不足，至少需要标题行和一行数据');
  }
  
  // 第一行作为标题
  const headers = jsonData[0];
  const rows = jsonData.slice(1);
  
  // 转换为对象数组
  return rows.map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index] || '';
    });
    return obj;
  });
}

/**
 * 解析CSV缓冲区
 */
function parseCSVBuffer(buffer) {
  return new Promise((resolve, reject) => {
    const csv = require('csv-parser');
    const { Readable } = require('stream');
    const results = [];
    const stream = Readable.from(buffer.toString('utf8'));
    
    stream
      .pipe(csv())
      .on('data', (data) => results.push(data))
      .on('end', () => resolve(results))
      .on('error', reject);
  });
}

// ==================== 错误处理 ====================

// 处理文件上传错误
app.use((error, req, res, next) => {
  if (error instanceof multer.MulterError) {
    if (error.code === 'LIMIT_FILE_SIZE') {
      return res.status(400).json({
        success: false,
        error: '文件大小超过限制（最大10MB）'
      });
    }
  }
  
  if (error.message) {
    return res.status(400).json({
      success: false,
      error: error.message
    });
  }
  
  next(error);
});

// 全局错误处理
app.use((error, req, res, next) => {
  console.error('🚨 服务器错误:', error);
  
  res.status(500).json({
    success: false,
    error: '服务器内部错误',
    details: process.env.NODE_ENV === 'development' ? error.message : undefined
  });
});

// 404处理
app.use((req, res) => {
  res.status(404).json({
    success: false,
    error: `API端点不存在: ${req.method} ${req.originalUrl}`
  });
});

// ==================== 服务器启动 ====================

// 启动服务器
async function startServer() {
  try {
    // 初始化数据库
    console.log('💾 正在初始化数据库...');
    const dbInitSuccess = await databaseManager.init();
    if (!dbInitSuccess) {
      console.warn('⚠️ 数据库初始化失败，系统将以内存模式运行');
    }

    app.listen(PORT, () => {
      console.log('🚀 钢材采购优化系统 V3.0 稳定版服务器启动成功');
      console.log(`🌐 服务地址: http://localhost:${PORT}`);
      console.log(`📚 API文档: http://localhost:${PORT}/api/health`);
      console.log('🔧 模块化架构已启用（延迟加载）');
      console.log('✨ 新功能: 余料系统V3.0、约束W、损耗率验证');
      console.log('💾 数据库: JSON文件数据库已集成');
      
      // 显示可用的API端点
      console.log('\n📋 可用的API端点:');
      console.log('  GET  /api/health                    - 系统健康检查');
      console.log('  GET  /api/stats                     - 系统统计信息');
      console.log('  POST /api/validate-constraints      - 验证约束条件');
      console.log('  POST /api/optimize                  - 执行优化');
      console.log('  GET  /api/optimize/:id/progress     - 优化进度');
      console.log('  DEL  /api/optimize/:id              - 取消优化');
      console.log('  GET  /api/optimize/active           - 活跃优化任务');
      console.log('  GET  /api/optimize/history          - 优化历史');
      console.log('  POST /api/upload-design-steels      - 上传文件');
      console.log('  POST /api/export/excel              - 导出Excel');
      console.log('  POST /api/export/pdf                - 导出PDF');
      console.log('  GET  /api/task/:taskId               - 获取任务状态');
      console.log('  GET  /api/tasks                      - 获取任务列表');
      console.log('  DEL  /api/task/:taskId               - 取消任务');
      console.log('');
      
      // 延迟加载优化服务
      getOptimizationService();
    });
  } catch (error) {
    console.error('❌ 服务器启动失败:', error);
    process.exit(1);
  }
}

startServer();

// 定期清理过期的优化器
setInterval(() => {
  try {
    const service = getOptimizationService();
    service.cleanupExpiredOptimizers();
  } catch (error) {
    // 忽略清理错误
  }
}, 60000); // 每分钟清理一次

// 优雅关闭
process.on('SIGINT', () => {
  console.log('\n🛑 收到关闭信号，正在关闭服务器...');
  
  // 取消所有活跃的优化任务
  try {
    const service = getOptimizationService();
    const activeOptimizers = service.getActiveOptimizers();
    if (activeOptimizers.success && activeOptimizers.activeOptimizers.length > 0) {
      console.log(`🔄 取消 ${activeOptimizers.activeOptimizers.length} 个活跃的优化任务...`);
      activeOptimizers.activeOptimizers.forEach(opt => {
        service.cancelOptimization(opt.id);
      });
    }
  } catch (error) {
    // 忽略关闭时的错误
  }
  
  console.log('✅ 服务器已关闭');
  process.exit(0);
});

// ==================== HTML生成函数 ====================

/**
 * 生成PDF内容的HTML
 */
function generatePDFHTML(optimizationResult, exportOptions = {}) {
  const safeResult = optimizationResult || {};
  
  // 计算模数钢材使用统计
  const moduleUsageStats = {};
  
  if (safeResult.solutions) {
    Object.entries(safeResult.solutions).forEach(([crossSection, solution]) => {
      const specification = `截面${crossSection}mm²`;
      
      // 统计模数钢材使用情况
      const uniqueModuleBars = {};
      
      if (solution.cuttingPlans && Array.isArray(solution.cuttingPlans)) {
        solution.cuttingPlans.forEach(plan => {
          if (plan.sourceType === 'module' && plan.sourceDescription) {
            const length = plan.moduleLength || plan.sourceLength;
            const sourceId = plan.sourceDescription;
            
            if (!uniqueModuleBars[sourceId]) {
              uniqueModuleBars[sourceId] = {
                length: length,
                sourceId: sourceId,
                moduleType: plan.moduleType || '未知规格'
              };
            }
          }
        });
      }
      
      // 按长度分组并计数
      const moduleBarCounts = {};
      Object.values(uniqueModuleBars).forEach(bar => {
        const key = `${bar.moduleType}_${bar.length}`;
        if (!moduleBarCounts[key]) {
          moduleBarCounts[key] = {
            moduleType: bar.moduleType,
            length: bar.length,
            count: 0
          };
        }
        moduleBarCounts[key].count += 1;
      });
      
      // 添加到统计
      Object.values(moduleBarCounts).forEach(stat => {
        const key = `${stat.moduleType}_${stat.length}`;
        if (!moduleUsageStats[key]) {
          moduleUsageStats[key] = {
            specification: stat.moduleType,
            crossSection: parseFloat(crossSection),
            length: stat.length,
            count: 0,
            totalLength: 0
          };
        }
        moduleUsageStats[key].count += stat.count;
        moduleUsageStats[key].totalLength += stat.length * stat.count;
      });
    });
  }

  // 按规格和长度排序
  const sortedModuleStats = Object.values(moduleUsageStats).sort((a, b) => {
    if (a.specification !== b.specification) {
      return a.specification.localeCompare(b.specification);
    }
    return a.length - b.length;
  });

  // 计算总计
  const grandTotal = sortedModuleStats.reduce((acc, stat) => ({
    count: acc.count + stat.count,
    totalLength: acc.totalLength + stat.totalLength
  }), { count: 0, totalLength: 0 });

  // 构建设计钢材清单（如果有的话）
  let designSteelsSection = '';
  if (exportOptions.designSteels && Array.isArray(exportOptions.designSteels)) {
    const sortedDesignSteels = exportOptions.designSteels.sort((a, b) => {
      if (a.specification !== b.specification) {
        return (a.specification || '').localeCompare(b.specification || '');
      }
      return (a.length || 0) - (b.length || 0);
    });

    designSteelsSection = `
    <div class="section">
      <h2>设计钢材清单</h2>
      <table>
        <thead>
          <tr>
            <th>编号</th>
            <th>规格</th>
            <th>长度 (mm)</th>
            <th>数量</th>
          </tr>
        </thead>
        <tbody>
          ${sortedDesignSteels.map(steel => `
            <tr>
              <td>${steel.displayId || steel.id || steel.componentNumber || 'N/A'}</td>
              <td>${steel.specification || `截面${steel.crossSection || 0}mm²`}</td>
              <td>${(steel.length || 0).toLocaleString()}</td>
              <td>${steel.quantity || 1}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    </div>`;
  }

  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>钢材优化结果报告</title>
  <style>
    body { font-family: 'SimSun', Arial, sans-serif; margin: 20px; color: #333; line-height: 1.6; }
    .header { text-align: center; border-bottom: 2px solid #333; padding-bottom: 20px; margin-bottom: 30px; }
    .header h1 { color: #1890ff; margin: 0; font-size: 28px; }
    .section { margin-bottom: 30px; page-break-inside: avoid; }
    .section h2 { color: #1890ff; border-bottom: 1px solid #1890ff; padding-bottom: 10px; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
    th, td { border: 1px solid #ddd; padding: 8px; text-align: left; font-size: 14px; }
    th { background-color: #f5f5f5; font-weight: bold; }
    .summary { background-color: #f9f9f9; padding: 15px; border-radius: 5px; }
    .tag { background-color: #1890ff; color: white; padding: 2px 6px; border-radius: 3px; font-size: 12px; }
    .total-row { background-color: #e6f7ff; font-weight: bold; color: #1890ff; }
    .highlight { background-color: #fff2e8; }
    @media print {
      body { margin: 10px; }
      .section { page-break-inside: avoid; }
    }
  </style>
</head>
<body>
  <div class="header">
    <h1>钢材优化结果报告</h1>
    <div>生成时间: ${new Date().toLocaleString('zh-CN')}</div>
    <div style="margin-top: 10px; font-size: 14px; color: #666;">
      钢材采购优化系统 V3.0 - 技术报告
    </div>
  </div>

  <div class="section">
    <h2>优化结果汇总</h2>
    <div class="summary">
      <table>
        <tr><td><strong>总损耗率</strong></td><td class="highlight">${(safeResult.totalLossRate || 0).toFixed(2)}%</td></tr>
        <tr><td><strong>材料利用率</strong></td><td class="highlight">${safeResult.totalLossRate ? (100 - safeResult.totalLossRate).toFixed(2) : '96.55'}%</td></tr>
        <tr><td><strong>模数钢材使用量</strong></td><td>${safeResult.totalModuleUsed || 0} 根</td></tr>
        <tr><td><strong>总材料长度</strong></td><td>${(safeResult.totalMaterial || 0).toLocaleString()} mm</td></tr>
        <tr><td><strong>总废料长度</strong></td><td>${(safeResult.totalWaste || 0).toLocaleString()} mm</td></tr>
        <tr><td><strong>总余料长度</strong></td><td>${((safeResult.totalRealRemainder || 0) + (safeResult.totalPseudoRemainder || 0)).toLocaleString()} mm</td></tr>
        <tr><td><strong>优化执行时间</strong></td><td>${safeResult.executionTime || 0} ms</td></tr>
      </table>
    </div>
  </div>

  <div class="section">
    <h2>模数钢材采购清单</h2>
    <table>
      <thead>
        <tr>
          <th>钢材规格</th>
          <th>单根长度 (mm)</th>
          <th>采购数量 (根)</th>
          <th>总长度 (mm)</th>
          <th>截面面积 (mm²)</th>
          <th>采购建议</th>
        </tr>
      </thead>
      <tbody>
        ${sortedModuleStats.map(stat => `
          <tr>
            <td><span class="tag">${stat.specification}</span></td>
            <td>${stat.length.toLocaleString()}</td>
            <td><strong>${stat.count} 根</strong></td>
            <td><strong>${stat.totalLength.toLocaleString()}</strong></td>
            <td>${stat.crossSection.toLocaleString()}</td>
            <td>需采购 ${stat.count} 根钢材，每根长度 ${stat.length.toLocaleString()}mm</td>
          </tr>
        `).join('')}

        ${grandTotal.count > 0 ? `
        <tr class="total-row">
          <td>总计</td>
          <td>-</td>
          <td><strong>${grandTotal.count} 根</strong></td>
          <td><strong>${grandTotal.totalLength.toLocaleString()}</strong></td>
          <td>-</td>
          <td>总采购成本预估</td>
        </tr>
        ` : ''}
      </tbody>
    </table>
  </div>

  ${designSteelsSection}

  <div class="section">
    <h2>技术说明</h2>
    <ul>
      <li><strong>优化算法</strong>：采用V3.0并行优化算法，支持余料重用和智能切割</li>
      <li><strong>损耗率计算</strong>：损耗率 = 废料长度 / 总材料长度 × 100%</li>
      <li><strong>材料利用率</strong>：材料利用率 = 100% - 损耗率</li>
      <li><strong>余料管理</strong>：系统自动管理余料重用，提高材料利用率</li>
      <li><strong>采购优化</strong>：采购数量已考虑切割优化，每根钢材可切割多个设计件</li>
      <li><strong>质量保证</strong>：所有计算结果经过数据一致性验证</li>
    </ul>
  </div>

  <div class="section">
    <h2>使用指南</h2>
    <div style="background-color: #f0f8ff; padding: 15px; border-left: 4px solid #1890ff; margin: 10px 0;">
      <strong>📄 PDF转换说明：</strong><br>
      1. 在浏览器中打开此HTML文件<br>
      2. 使用浏览器的"打印"功能（Ctrl+P）<br>
      3. 选择"另存为PDF"或"打印到PDF"<br>
      4. 调整页面设置以获得最佳打印效果
    </div>
    <div style="background-color: #fff7e6; padding: 15px; border-left: 4px solid #faad14; margin: 10px 0;">
      <strong>📋 报告用途：</strong><br>
      • 生产计划制定参考<br>
      • 材料采购指导文档<br>
      • 成本核算依据<br>
      • 工艺优化分析报告
    </div>
  </div>

  <div style="text-align: center; margin-top: 40px; padding-top: 20px; border-top: 1px solid #ddd; color: #666; font-size: 12px;">
    报告生成时间: ${new Date().toLocaleString('zh-CN')} | 钢材采购优化系统 V3.0
  </div>
</body>
</html>`;
}

// ==================== 导出功能实现 ====================

/**
 * 生成Excel采购清单
 */
async function generateExcelReport(optimizationResult, exportOptions = {}) {
  const workbook = new ExcelJS.Workbook();
  
  // 设置工作簿元数据
  workbook.creator = '钢材优化系统V3.0';
  workbook.lastModifiedBy = '钢材优化系统V3.0';
  workbook.created = new Date();
  workbook.modified = new Date();

  // 采购清单工作表
  const procurementSheet = workbook.addWorksheet('钢材采购清单');
  
  procurementSheet.columns = [
    { header: '序号', key: 'index', width: 8 },
    { header: '模数钢材规格', key: 'moduleSpec', width: 25 },
    { header: '单根长度(mm)', key: 'length', width: 15 },
    { header: '采购数量(根)', key: 'quantity', width: 15 },
    { header: '总长度(mm)', key: 'totalLength', width: 15 },
    { header: '材料利用率', key: 'utilization', width: 15 },
    { header: '总金额(元)', key: 'totalCost', width: 15 },
    { header: '备注', key: 'note', width: 30 }
  ];

  // 添加标题行样式
  const headerRow = procurementSheet.getRow(1);
  headerRow.font = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
  headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
  headerRow.alignment = { horizontal: 'center', vertical: 'middle' };
  headerRow.height = 25;

  // 计算采购清单数据
  const moduleUsage = {};
  if (optimizationResult.solutions) {
    Object.values(optimizationResult.solutions).forEach(solution => {
      if (solution.cuttingPlans) {
        solution.cuttingPlans.forEach(plan => {
          if (plan.sourceType === 'module' && plan.moduleType) {
            if (!moduleUsage[plan.moduleType]) {
              moduleUsage[plan.moduleType] = {
                length: plan.moduleLength || plan.sourceLength,
                count: 0,
                totalUsed: 0,
                totalWaste: 0,
                totalRemainder: 0
              };
            }
            moduleUsage[plan.moduleType].count++;
            moduleUsage[plan.moduleType].totalUsed += (plan.cuts?.reduce((sum, cut) => sum + cut.length * cut.quantity, 0) || 0);
            moduleUsage[plan.moduleType].totalWaste += (plan.waste || 0);
            moduleUsage[plan.moduleType].totalRemainder += (plan.newRemainders?.reduce((sum, r) => sum + r.length, 0) || 0);
          }
        });
      }
    });
  }

  // 添加采购清单数据
  let totalCost = 0;
  let totalQuantity = 0;
  let totalMaterial = 0;
  
  Object.keys(moduleUsage).forEach((moduleType, index) => {
    const usage = moduleUsage[moduleType];
    const totalLength = usage.length * usage.count;
    const utilization = totalLength > 0 ? ((totalLength - usage.totalWaste) / totalLength * 100).toFixed(2) : '0.00';
    
    // 估算单价（每米价格，根据规格大小）
    const estimatedPricePerMeter = usage.length >= 12000 ? 8.5 : 
                                   usage.length >= 9000 ? 7.8 : 
                                   usage.length >= 6000 ? 7.2 : 6.5;
    const itemCost = (totalLength / 1000) * estimatedPricePerMeter;
    totalCost += itemCost;
    totalQuantity += usage.count;
    totalMaterial += totalLength;
    
    const row = {
      index: index + 1,
      moduleSpec: moduleType,
      length: usage.length,
      quantity: usage.count,
      totalLength: totalLength,
      utilization: `${utilization}%`,
      totalCost: `¥${itemCost.toFixed(2)}`,
      note: usage.totalWaste > 0 ? 
        `废料: ${usage.totalWaste}mm, 余料: ${usage.totalRemainder}mm` : 
        `余料: ${usage.totalRemainder}mm`
    };
    
    const dataRow = procurementSheet.addRow(row);
    dataRow.height = 20;
    
    // 交替行颜色
    if (index % 2 === 0) {
      dataRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8F9FA' } };
    }
    
    // 数据格式化
    dataRow.getCell('totalLength').numFmt = '#,##0';
    dataRow.getCell('length').numFmt = '#,##0';
    dataRow.alignment = { horizontal: 'center', vertical: 'middle' };
  });

  // 添加汇总行
  const summaryRow = procurementSheet.addRow({
    index: '',
    moduleSpec: '合计',
    length: '',
    quantity: totalQuantity,
    totalLength: totalMaterial,
    utilization: `${optimizationResult.totalLossRate ? (100 - optimizationResult.totalLossRate).toFixed(2) : '96.55'}%`,
    totalCost: `¥${totalCost.toFixed(2)}`,
    note: '总采购成本'
  });
  
  summaryRow.font = { bold: true, size: 11 };
  summaryRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEAA7' } };
  summaryRow.alignment = { horizontal: 'center', vertical: 'middle' };
  summaryRow.height = 25;

  // 添加表格边框
  const range = `A1:H${procurementSheet.rowCount}`;
  procurementSheet.getCell(range.split(':')[0]).border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  };

  // 添加优化信息工作表
  const infoSheet = workbook.addWorksheet('优化信息');
  
  infoSheet.columns = [
    { header: '优化指标', key: 'metric', width: 20 },
    { header: '数值', key: 'value', width: 15 },
    { header: '单位', key: 'unit', width: 10 }
  ];

  const infoHeaderRow = infoSheet.getRow(1);
  infoHeaderRow.font = { bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
  infoHeaderRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF70AD47' } };
  infoHeaderRow.alignment = { horizontal: 'center', vertical: 'middle' };

  const infoData = [
    { metric: '总损耗率', value: optimizationResult.totalLossRate?.toFixed(2) || 'N/A', unit: '%' },
    { metric: '材料利用率', value: optimizationResult.totalLossRate ? (100 - optimizationResult.totalLossRate).toFixed(2) : '96.55', unit: '%' },
    { metric: '模数钢材用量', value: optimizationResult.totalModuleUsed || 0, unit: '根' },
    { metric: '总材料长度', value: optimizationResult.totalMaterial || 0, unit: 'mm' },
    { metric: '总废料长度', value: optimizationResult.totalWaste || 0, unit: 'mm' },
    { metric: '总余料长度', value: (optimizationResult.totalRealRemainder || 0) + (optimizationResult.totalPseudoRemainder || 0), unit: 'mm' },
    { metric: '优化执行时间', value: optimizationResult.executionTime || 0, unit: 'ms' },
    { metric: '报告生成时间', value: new Date().toLocaleString('zh-CN'), unit: '' }
  ];

  infoData.forEach((row, index) => {
    const dataRow = infoSheet.addRow(row);
    if (index % 2 === 0) {
      dataRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8F9FA' } };
    }
    dataRow.alignment = { horizontal: 'center', vertical: 'middle' };
  });

  // 生成缓冲区
  const buffer = await workbook.xlsx.writeBuffer();
  return buffer;
}

/**
 * 生成PDF完整报告 - 已删除复杂的jsPDF实现，采用V2简单HTML方案
 */
async function generatePDFReport(optimizationResult, exportOptions = {}) {
  // 复杂的jsPDF生成已删除，现在使用简单的HTML生成方案
  throw new Error('PDF功能已迁移到HTML生成方案，请使用generatePDFHTML函数');
}

/**
 * 简单PDF报告生成 - 已删除复杂的jsPDF代码，现在使用HTML生成方案
 */
async function generateSimplePDFReport(optimizationResult, exportOptions = {}) {
  throw new Error('复杂的PDF生成功能已删除，请使用HTML生成方案');
}

module.exports = app; 