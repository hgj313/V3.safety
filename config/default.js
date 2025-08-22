/**
 * 钢材采购优化系统 V3.0 - 默认配置
 */

module.exports = {
  // 服务器配置
  server: {
    port: process.env.PORT || 3001,
    host: process.env.HOST || 'localhost',
    cors: {
      origin: process.env.CORS_ORIGIN || '*',
      credentials: true
    },
    bodyParser: {
      limit: '10mb'
    }
  },

  // 优化算法配置
  optimization: {
    // 🔧 修复：默认约束条件现在由约束配置中心统一管理
    // 这里保留配置是为了向后兼容，实际使用时会优先使用约束配置中心
    defaultConstraints: {
      wasteThreshold: 100,        // 废料阈值 (mm) - 与约束配置中心保持一致
      expectedLossRate: 5,        // 期望损耗率 (%) - 与约束配置中心保持一致
      timeLimit: 30000,           // 时间限制 (ms) - 与约束配置中心保持一致
      weldingSegments: 1          // 允许焊接段数W - 与约束配置中心保持一致
    },
    
    // 算法参数
    algorithm: {
      maxIterations: 1000,        // 最大迭代次数
      convergenceThreshold: 0.01, // 收敛阈值
      mutationRate: 0.1,          // 变异率
      populationSize: 50          // 种群大小
    },

    // 性能配置
    performance: {
      maxActiveOptimizers: 10,    // 最大并发优化器数量
      optimizerTimeout: 300000,   // 优化器超时时间 (5分钟)
      cleanupInterval: 60000,     // 清理间隔 (1分钟)
      historyLimit: 100           // 历史记录数量限制
    }
  },

  // 余料系统配置
  remainder: {
    // 余料ID生成配置
    idGeneration: {
      letterLimit: 50,            // 单字母数量限制
      maxLetters: 26              // 最大字母数量
    },
    
    // 余料分类参数
    classification: {
      wasteSizeThreshold: 100,    // 废料尺寸阈值 (mm)
      lengthDecayFactor: 0.9      // 长度衰减因子
    }
  },

  // 文件处理配置
  fileProcessing: {
    upload: {
      maxSize: 10 * 1024 * 1024,  // 最大文件大小 (10MB)
      allowedTypes: ['.csv', '.xlsx', '.xls'],
      tempDir: './server/uploads',
      retentionDays: 7            // 文件保留天数
    },
    
    parsing: {
      csvDelimiter: ',',          // CSV分隔符
      encoding: 'utf8',           // 文件编码
      headerRow: true,            // 是否包含标题行
      maxRows: 10000              // 最大行数
    }
  },

  // 导出配置
  export: {
    excel: {
      defaultTemplate: 'standard',
      maxWorksheets: 10,
      compression: true
    },
    
    pdf: {
      pageSize: 'A4',
      orientation: 'portrait',
      margins: {
        top: 20,
        right: 20,
        bottom: 20,
        left: 20
      },
      fontSize: 12,
      includeCharts: true
    }
  },

  // 数据库配置（如果需要）
  database: {
    type: process.env.DB_TYPE || 'memory',
    connection: {
      host: process.env.DB_HOST || 'localhost',
      port: process.env.DB_PORT || 5432,
      database: process.env.DB_NAME || 'hgj_steel_v3',
      username: process.env.DB_USER || 'postgres',
      password: process.env.DB_PASS || ''
    },
    pool: {
      min: 2,
      max: 10,
      idle: 10000
    }
  },

  // 日志配置
  logging: {
    level: process.env.LOG_LEVEL || 'info',
    file: {
      enabled: true,
      path: './logs/app.log',
      maxSize: '10MB',
      maxFiles: 5
    },
    console: {
      enabled: true,
      colorize: true,
      timestamp: true
    }
  },

  // 缓存配置
  cache: {
    type: 'memory',             // 缓存类型: memory, redis
    ttl: 3600,                  // 默认TTL (秒)
    maxSize: 100,               // 最大缓存条目数
    checkPeriod: 600            // 检查周期 (秒)
  },

  // API配置
  api: {
    version: '3.0.0',
    prefix: '/api',
    rateLimit: {
      windowMs: 15 * 60 * 1000, // 15分钟
      max: 100                  // 最大请求数
    },
    validation: {
      stripUnknown: true,       // 移除未知字段
      abortEarly: false         // 不在第一个错误时停止
    }
  },

  // 前端配置
  client: {
    build: {
      outputDir: './client/build',
      publicPath: '/',
      staticDir: './client/public'
    },
    
    ui: {
      theme: {
        defaultMode: 'light',   // light, dark, auto
        primaryColor: '#1890ff',
        borderRadius: 8
      },
      
      animation: {
        duration: 300,          // 默认动画时长 (ms)
        easing: 'ease-in-out'   // 缓动函数
      },
      
      pagination: {
        defaultPageSize: 20,    // 默认页面大小
        showSizeChanger: true,  // 显示页面大小选择器
        showQuickJumper: true   // 显示快速跳转
      }
    }
  },

  // 安全配置
  security: {
    helmet: {
      enabled: true,
      contentSecurityPolicy: false
    },
    
    session: {
      secret: process.env.SESSION_SECRET || 'hgj-steel-v3-secret',
      resave: false,
      saveUninitialized: false,
      cookie: {
        maxAge: 24 * 60 * 60 * 1000 // 24小时
      }
    }
  },

  // 监控配置
  monitoring: {
    metrics: {
      enabled: true,
      interval: 5000,           // 收集间隔 (ms)
      retention: 24 * 60 * 60   // 数据保留时间 (秒)
    },
    
    health: {
      enabled: true,
      checks: [
        'server',
        'optimizer',
        'memory',
        'disk'
      ]
    }
  },

  // 开发配置
  development: {
    hotReload: true,
    debugMode: process.env.NODE_ENV === 'development',
    mockData: {
      enabled: false,
      dataSet: 'sample'
    }
  },

  // 生产配置
  production: {
    compression: true,
    staticCaching: true,
    errorReporting: {
      enabled: true,
      level: 'error'
    }
  }
}; 