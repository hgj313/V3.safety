-- 钢材采购优化系统 V3.0 数据库表结构
-- 创建时间: 2024-01-20
-- 说明: 使用SQLite数据库存储系统数据

-- 1. 设计钢材表
CREATE TABLE IF NOT EXISTS design_steels (
    id TEXT PRIMARY KEY,
    display_id TEXT,
    length INTEGER NOT NULL,
    quantity INTEGER NOT NULL,
    cross_section REAL NOT NULL,
    component_number TEXT,
    specification TEXT,
    part_number TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
);

-- 2. 模数钢材表
CREATE TABLE IF NOT EXISTS module_steels (
    id TEXT PRIMARY KEY,
    name TEXT NOT NULL UNIQUE,
    length INTEGER NOT NULL,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
);

-- 3. 优化任务表
CREATE TABLE IF NOT EXISTS optimization_tasks (
    id TEXT PRIMARY KEY,
    status TEXT NOT NULL CHECK (status IN ('idle', 'running', 'completed', 'error')),
    progress INTEGER DEFAULT 0,
    error_message TEXT,
    waste_threshold INTEGER,
    target_loss_rate REAL,
    time_limit INTEGER,
    max_welding_segments INTEGER,
    results TEXT, -- JSON格式存储优化结果
    execution_time INTEGER,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
);

-- 4. 优化任务与设计钢材关联表
CREATE TABLE IF NOT EXISTS task_design_steels (
    task_id TEXT,
    steel_id TEXT,
    PRIMARY KEY (task_id, steel_id),
    FOREIGN KEY (task_id) REFERENCES optimization_tasks(id) ON DELETE CASCADE,
    FOREIGN KEY (steel_id) REFERENCES design_steels(id) ON DELETE CASCADE
);

-- 5. 优化任务与模数钢材关联表
CREATE TABLE IF NOT EXISTS task_module_steels (
    task_id TEXT,
    steel_id TEXT,
    PRIMARY KEY (task_id, steel_id),
    FOREIGN KEY (task_id) REFERENCES optimization_tasks(id) ON DELETE CASCADE,
    FOREIGN KEY (steel_id) REFERENCES module_steels(id) ON DELETE CASCADE
);

-- 6. 系统统计表
CREATE TABLE IF NOT EXISTS system_stats (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    total_optimizations INTEGER DEFAULT 0,
    total_design_steels INTEGER DEFAULT 0,
    total_module_steels INTEGER DEFAULT 0,
    total_saved_cost REAL DEFAULT 0,
    last_updated DATETIME DEFAULT CURRENT_TIMESTAMP
);

-- 7. 操作日志表
CREATE TABLE IF NOT EXISTS operation_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    operation_type TEXT NOT NULL, -- 'upload', 'optimize', 'export', 'delete' etc.
    description TEXT,
    details TEXT, -- JSON格式存储详细信息
    ip_address TEXT,
    user_agent TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
);

-- 初始化系统统计数据
INSERT OR IGNORE INTO system_stats (id, total_optimizations) VALUES (1, 0);

-- 创建索引以提高查询性能
CREATE INDEX IF NOT EXISTS idx_design_steels_length ON design_steels(length);
CREATE INDEX IF NOT EXISTS idx_design_steels_created_at ON design_steels(created_at);
CREATE INDEX IF NOT EXISTS idx_module_steels_length ON module_steels(length);
CREATE INDEX IF NOT EXISTS idx_optimization_tasks_status ON optimization_tasks(status);
CREATE INDEX IF NOT EXISTS idx_optimization_tasks_created_at ON optimization_tasks(created_at);
CREATE INDEX IF NOT EXISTS idx_operation_logs_created_at ON operation_logs(created_at);
CREATE INDEX IF NOT EXISTS idx_operation_logs_type ON operation_logs(operation_type); 