-- 钢材采购优化系统 V3.0 - Netlify数据库初始化脚本
-- 适用于Neon PostgreSQL

-- 删除现有表（如果存在）
DROP TABLE IF EXISTS optimization_tasks CASCADE;
DROP TABLE IF EXISTS system_logs CASCADE;
DROP TABLE IF EXISTS user_preferences CASCADE;

-- 创建优化任务表
CREATE TABLE optimization_tasks (
    id TEXT PRIMARY KEY,
    type TEXT NOT NULL DEFAULT 'optimization',
    status TEXT NOT NULL CHECK (status IN ('pending', 'running', 'completed', 'failed', 'cancelled')),
    progress INTEGER DEFAULT 0 CHECK (progress >= 0 AND progress <= 100),
    message TEXT,
    input_data JSONB,
    results JSONB,
    error_message TEXT,
    execution_time INTEGER,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- 创建系统日志表
CREATE TABLE system_logs (
    id SERIAL PRIMARY KEY,
    level TEXT NOT NULL CHECK (level IN ('info', 'warn', 'error', 'debug')),
    message TEXT NOT NULL,
    context JSONB,
    source TEXT,
    user_id TEXT,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- 创建用户偏好设置表
CREATE TABLE user_preferences (
    id SERIAL PRIMARY KEY,
    user_id TEXT NOT NULL,
    preference_key TEXT NOT NULL,
    preference_value JSONB,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    UNIQUE(user_id, preference_key)
);

-- 创建索引以提高查询性能
CREATE INDEX idx_optimization_tasks_status ON optimization_tasks(status);
CREATE INDEX idx_optimization_tasks_created_at ON optimization_tasks(created_at DESC);
CREATE INDEX idx_optimization_tasks_updated_at ON optimization_tasks(updated_at DESC);
CREATE INDEX idx_optimization_tasks_type ON optimization_tasks(type);

CREATE INDEX idx_system_logs_level ON system_logs(level);
CREATE INDEX idx_system_logs_created_at ON system_logs(created_at DESC);
CREATE INDEX idx_system_logs_source ON system_logs(source);

CREATE INDEX idx_user_preferences_user_id ON user_preferences(user_id);
CREATE INDEX idx_user_preferences_key ON user_preferences(preference_key);

-- 创建更新时间戳的触发器函数
CREATE OR REPLACE FUNCTION update_updated_at_column()
RETURNS TRIGGER AS $$
BEGIN
    NEW.updated_at = NOW();
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

-- 为需要的表创建触发器
CREATE TRIGGER update_optimization_tasks_updated_at
    BEFORE UPDATE ON optimization_tasks
    FOR EACH ROW
    EXECUTE FUNCTION update_updated_at_column();

CREATE TRIGGER update_user_preferences_updated_at
    BEFORE UPDATE ON user_preferences
    FOR EACH ROW
    EXECUTE FUNCTION update_updated_at_column();

-- 插入初始数据
INSERT INTO system_logs (level, message, context, source) VALUES 
('info', '数据库初始化完成', '{"version": "3.0.0", "platform": "Netlify"}', 'system'),
('info', '表结构创建完成', '{"tables": ["optimization_tasks", "system_logs", "user_preferences"]}', 'system');

-- 创建清理过期任务的存储过程
CREATE OR REPLACE FUNCTION cleanup_expired_tasks()
RETURNS INTEGER AS $$
DECLARE
    deleted_count INTEGER;
BEGIN
    -- 删除7天前的已完成任务
    DELETE FROM optimization_tasks 
    WHERE status IN ('completed', 'failed', 'cancelled') 
    AND created_at < NOW() - INTERVAL '7 days';
    
    GET DIAGNOSTICS deleted_count = ROW_COUNT;
    
    -- 记录清理日志
    INSERT INTO system_logs (level, message, context, source) 
    VALUES ('info', '清理过期任务完成', 
            json_build_object('deleted_count', deleted_count)::jsonb, 
            'cleanup_job');
    
    RETURN deleted_count;
END;
$$ LANGUAGE plpgsql;

-- 创建获取系统统计信息的存储过程
CREATE OR REPLACE FUNCTION get_system_stats()
RETURNS JSONB AS $$
DECLARE
    stats JSONB;
BEGIN
    SELECT json_build_object(
        'total_tasks', COUNT(*),
        'pending_tasks', COUNT(*) FILTER (WHERE status = 'pending'),
        'running_tasks', COUNT(*) FILTER (WHERE status = 'running'),
        'completed_tasks', COUNT(*) FILTER (WHERE status = 'completed'),
        'failed_tasks', COUNT(*) FILTER (WHERE status = 'failed'),
        'cancelled_tasks', COUNT(*) FILTER (WHERE status = 'cancelled'),
        'success_rate', CASE 
            WHEN COUNT(*) > 0 THEN 
                ROUND((COUNT(*) FILTER (WHERE status = 'completed')::float / COUNT(*) * 100), 2)
            ELSE 0
        END,
        'avg_execution_time', ROUND(AVG(execution_time) FILTER (WHERE execution_time IS NOT NULL), 2),
        'last_24h_tasks', COUNT(*) FILTER (WHERE created_at > NOW() - INTERVAL '24 hours'),
        'database_size', pg_size_pretty(pg_database_size(current_database())),
        'last_updated', NOW()
    )::jsonb INTO stats
    FROM optimization_tasks;
    
    RETURN stats;
END;
$$ LANGUAGE plpgsql;

-- 创建数据库版本信息表
CREATE TABLE database_metadata (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL,
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- 插入版本信息
INSERT INTO database_metadata (key, value) VALUES 
('version', '3.0.0'),
('platform', 'Netlify'),
('database_type', 'Neon PostgreSQL'),
('initialized_at', NOW()::text),
('schema_version', '1.0');

-- 创建性能监控视图
CREATE VIEW task_performance_summary AS
SELECT 
    DATE_TRUNC('day', created_at) as date,
    COUNT(*) as total_tasks,
    COUNT(*) FILTER (WHERE status = 'completed') as completed_tasks,
    COUNT(*) FILTER (WHERE status = 'failed') as failed_tasks,
    ROUND(AVG(execution_time) FILTER (WHERE execution_time IS NOT NULL), 2) as avg_execution_time,
    MAX(execution_time) as max_execution_time,
    MIN(execution_time) as min_execution_time
FROM optimization_tasks
WHERE created_at > NOW() - INTERVAL '30 days'
GROUP BY DATE_TRUNC('day', created_at)
ORDER BY date DESC;

-- 输出初始化完成信息
SELECT 
    '数据库初始化完成' as status,
    NOW() as initialized_at,
    '3.0.0' as version,
    'Netlify' as platform,
    current_database() as database_name;

-- 验证表结构
SELECT 
    schemaname,
    tablename,
    tableowner
FROM pg_tables 
WHERE schemaname = 'public' 
    AND tablename IN ('optimization_tasks', 'system_logs', 'user_preferences', 'database_metadata')
ORDER BY tablename; 