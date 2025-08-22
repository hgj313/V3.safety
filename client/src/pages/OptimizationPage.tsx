import React, { useState, useEffect, useRef } from 'react';
import { 
  Card, 
  Button, 
  Space, 
  Upload, 
  message, 
  Form, 
  Input, 
  Row, 
  Col, 
  Divider,
  Table,
  Progress,
  Alert,
  Typography,
  Tag,
  Modal,
  InputNumber,
  Collapse,
  Popconfirm
} from 'antd';
import { 
  PlayCircleOutlined, 
  UploadOutlined, 
  FileExcelOutlined,
  SettingOutlined,
  DeleteOutlined,

  PlusOutlined,
  EditOutlined,
  DownOutlined,
  RightOutlined,
  ReloadOutlined
} from '@ant-design/icons';
import styled from 'styled-components';
import { motion } from 'framer-motion';
// import type { UploadProps, UploadFile } from 'antd/es/upload/interface';
import { useOptimizationContext, DesignSteel, ModuleSteel } from '../contexts/OptimizationContext';
import { DEFAULT_CONSTRAINTS } from '../types';
import { useNavigate } from 'react-router-dom';
import { generateDisplayIds } from '../utils/steelUtils';

const { Title, Text } = Typography;
const { Panel } = Collapse;

const PageContainer = styled.div`
  height: 100%;
  overflow: auto;
  padding: 0;
`;

const StepCard = styled(Card)`
  margin-bottom: 24px;
  border-radius: 12px;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.06);
  
  .ant-card-head {
    border-bottom: 1px solid #f0f0f0;
  }
`;

const ActionButton = styled(Button)`
  border-radius: 8px;
  height: 40px;
  font-weight: 500;
`;

const OptimizationPage: React.FC = () => {
  const [form] = Form.useForm();
  const [moduleForm] = Form.useForm();
  
  // 从Context获取数据和方法
  const {
    designSteels,
    moduleSteels,
    constraints,
    isOptimizing,
    progress,
    error,
    setDesignSteels,
    removeDesignSteel,
    addModuleSteel,
    updateModuleSteel,
    removeModuleSteel,
    setConstraints,
    startOptimization,
    clearOptimizationData,
    currentOptimization // 确保我们获取了currentOptimization
  } = useOptimizationContext();
  
  const navigate = useNavigate();
  // 关键修复：使用useRef替代useState，防止不必要的重渲染导致跳转被取消
  const lastNavigatedTaskId = useRef<string | null>(null);
  
  // 监听优化任务状态，完成后自动跳转
  useEffect(() => {
    // 只有当存在一个已完成的优化任务，并且我们尚未为该任务导航过时，才执行跳转
    if (currentOptimization && currentOptimization.status === 'completed' && currentOptimization.id !== lastNavigatedTaskId.current) {
      message.success('优化完成！正在跳转到结果页面...', 1.5);
      
      // 记录我们已经为这个任务ID导航过，防止重复跳转
      lastNavigatedTaskId.current = currentOptimization.id;
      
      const timer = setTimeout(() => {
        navigate('/results');
      }, 1000);

      // 清理函数保持不变，以防组件在计时器完成前被卸载
      return () => clearTimeout(timer);
    }
  }, [currentOptimization, navigate]); // 依赖项中不再需要lastNavigatedTaskId，因为useRef的更新不会触发重渲染

  // 关键修复：在组件挂载时重置导航状态
  useEffect(() => {
    // 当用户进入优化页面时，重置导航记录
    lastNavigatedTaskId.current = null;
  }, []);
  
  // 本地UI状态
  const [designCollapsed, setDesignCollapsed] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [showDesignModal, setShowDesignModal] = useState(false);
  const [editingDesignSteel, setEditingDesignSteel] = useState<DesignSteel | null>(null);
  const [showModuleModal, setShowModuleModal] = useState(false);
  const [editingModuleSteel, setEditingModuleSteel] = useState<ModuleSteel | null>(null);
  const [selectedRowKeys, setSelectedRowKeys] = React.useState<React.Key[]>([]);

  // 统一使用稳健的、来自 utils 的显示ID生成逻辑
  const designSteelsForDisplay = React.useMemo(() => {
    return generateDisplayIds(designSteels);
  }, [designSteels]);

  // 生成唯一ID
  const generateId = () => Math.random().toString(36).substr(2, 9);

  const rowSelection = {
    selectedRowKeys,
    onChange: (newSelectedRowKeys: React.Key[]) => setSelectedRowKeys(newSelectedRowKeys),
  };

  const handleBatchDelete = () => {
    Modal.confirm({
      title: `确认删除选中的${selectedRowKeys.length}条设计钢材吗？`,
      content: '删除后数据不可恢复，请谨慎操作。',
      okText: '确认删除',
      okType: 'danger',
      cancelText: '取消',
      onOk: () => {
        setDesignSteels(designSteels.filter((item: DesignSteel) => !selectedRowKeys.includes(item.id)));
        setSelectedRowKeys([]);
        message.success('批量删除成功');
      },
    });
  };

  // ==================== 设计钢材管理 ====================
  
  // 处理设计钢材文件上传
  const handleDesignSteelUpload = async (file: File) => {
    setUploading(true);
    try {
      console.log('=== 设计钢材文件上传开始 ===');
      console.log('文件信息:', {
        name: file.name,
        size: file.size,
        type: file.type
      });
      
      // 读取文件内容
      const arrayBuffer = await file.arrayBuffer();
      const uint8Array = new Uint8Array(arrayBuffer);
      const base64 = btoa(Array.from(uint8Array, byte => String.fromCharCode(byte)).join(''));
      
      // API调用
      const response = await fetch('/api/upload-design-steels', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          filename: file.name,
          data: base64,
          type: file.type
        })
      });
      
      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`HTTP ${response.status}: ${errorText}`);
      }
      
      const result = await response.json();
      
      if (result.success) {
        // 使用与结果页完全一致的稳定编号逻辑
        const steelsWithStableIds = generateDisplayIds(result.designSteels || []);
        
        // 清除之前的优化结果（因为上传了新数据）
        clearOptimizationData();
        
        // 保存到Context
        setDesignSteels(steelsWithStableIds);
        
        // 智能解析成功消息
        message.success(`${result.message} - 已应用稳定编号体系`);
        
        // 显示智能解析报告
        if (result.analysisReport) {
          const report = result.analysisReport;
          
          // 构建详细报告内容
          let reportContent = [];
          
          // 🆕 表头发现信息
          if (report.headerDiscovery) {
            reportContent.push("🎯 表头发现：");
            reportContent.push(`  • ${report.headerDiscovery.message}`);
            if (report.headerDiscovery.searchScore) {
              reportContent.push(`  • 识别置信度: ${report.headerDiscovery.searchScore}分`);
            }
          }
          
          // 字段识别情况
          if (Object.keys(report.fieldMapping).length > 0) {
            reportContent.push("📊 字段识别情况：");
            Object.entries(report.fieldMapping).forEach(([field, column]) => {
              const confidence = report.confidence[field] || 0;
              reportContent.push(`  • ${field}: "${column}" (置信度: ${confidence}%)`);
            });
          }
          
          // 数据清洗报告
          if (report.cleaningReport && report.cleaningReport.length > 0) {
            reportContent.push("🔧 数据清洗：");
            report.cleaningReport.forEach((action: string) => {
              reportContent.push(`  • ${action}`);
            });
          }
          
          // 未识别的列
          if (report.unidentifiedColumns && report.unidentifiedColumns.length > 0) {
            reportContent.push("⚠️ 未识别的列（已忽略）：");
            reportContent.push(`  • ${report.unidentifiedColumns.join(', ')}`);
          }
          
          // 显示详细报告
          if (reportContent.length > 0) {
            console.log('=== 智能解析报告 ===');
            console.log(reportContent.join('\n'));
            
            // 显示用户友好的解析摘要
            const summaryParts = [];
            if (report.headerDiscovery && report.headerDiscovery.foundAtRow > 1) {
              summaryParts.push(`智能发现表头在第${report.headerDiscovery.foundAtRow}行`);
            }
            if (report.dataStats.validRows > 0) {
              summaryParts.push(`成功解析 ${report.dataStats.validRows} 条数据`);
            }
            if (report.dataStats.skippedRows > 0) {
              summaryParts.push(`跳过 ${report.dataStats.skippedRows} 条无效数据`);
            }
            if (Object.keys(report.fieldMapping).length > 0) {
              summaryParts.push(`识别 ${Object.keys(report.fieldMapping).length} 个字段`);
            }
            
            if (summaryParts.length > 0) {
              message.info(
                `智能解析完成：${summaryParts.join('，')}。查看控制台了解详细信息。`,
                10 // 延长显示时间，因为信息更丰富
              );
            }
            
            // 🆕 特别提示表头发现功能
            if (report.headerDiscovery && report.headerDiscovery.foundAtRow > 1) {
              message.success(
                `✨ 智能功能：自动跳过了前${report.headerDiscovery.foundAtRow - 1}行，在第${report.headerDiscovery.foundAtRow}行找到表头！`,
                8
              );
            }
          }
        }
        
        // 显示传统调试信息（保持兼容性）
        if (result.debugInfo) {
          console.log('=== 解析统计信息 ===');
          console.log('原始行数:', result.debugInfo.原始行数);
          console.log('有效数据:', result.debugInfo.有效数据);
          console.log('跳过行数:', result.debugInfo.跳过行数);
          console.log('字段识别数:', result.debugInfo.字段识别);
          console.log('版本信息:', result.debugInfo.版本信息);
          
          // 兼容旧版本的截面面积统计提示
          if (result.debugInfo.截面面积统计?.无截面面积 > 0) {
            message.warning(
              `注意：${result.debugInfo.截面面积统计.无截面面积} 条数据的截面面积为0，已设为默认值1000mm²！`,
              6
            );
          }
        }
      } else {
        throw new Error(result.error || '上传失败');
      }
      
      console.log('=== 设计钢材文件上传完成 ===');
    } catch (error: any) {
      console.error('=== 设计钢材上传错误 ===', error);
      message.error(`上传失败: ${error.message}`);
    } finally {
      setUploading(false);
    }
    return false;
  };

  // 保存设计钢材
  const handleSaveDesignSteel = (values: any) => {
    const steel: DesignSteel = {
      id: editingDesignSteel?.id || generateId(),
      length: values.length,
      quantity: values.quantity,
      crossSection: values.crossSection,
      // 注意：displayId 由 useMemo 动态生成，此处无需手动赋值
      componentNumber: values.componentNumber,
      specification: values.specification,
      partNumber: values.partNumber
    };

    if (editingDesignSteel) {
      // 更新时，先找到旧数据并替换，然后让 useMemo 重新计算编号
      const updatedSteels = designSteels.map(s => s.id === editingDesignSteel.id ? steel : s);
      setDesignSteels(updatedSteels);
      message.success('设计钢材更新成功');
    } else {
      // 添加时，直接加入列表，让 useMemo 重新计算编号
      setDesignSteels([...designSteels, steel]);
      message.success('设计钢材添加成功');
    }

    setShowDesignModal(false);
    setEditingDesignSteel(null);
    form.resetFields();
  };

  // 删除设计钢材
  const handleDeleteDesignSteel = (steel: DesignSteel) => {
    removeDesignSteel(steel.id);
    message.success('设计钢材删除成功');
  };

  // 编辑设计钢材
  const handleEditDesignSteel = (steel: DesignSteel) => {
    setEditingDesignSteel(steel);
    form.setFieldsValue({
      length: steel.length,
      quantity: steel.quantity,
      crossSection: steel.crossSection,
      componentNumber: steel.componentNumber,
      specification: steel.specification,
      partNumber: steel.partNumber
    });
    setShowDesignModal(true);
  };

  // ==================== 模数钢材管理 ====================
  
  // 保存模数钢材
  const handleSaveModuleSteel = (values: any) => {
    const steel: ModuleSteel = {
      id: editingModuleSteel?.id || generateId(),
      name: values.name,
      length: values.length
    };

    if (editingModuleSteel) {
      updateModuleSteel(editingModuleSteel.id, steel);
      message.success('模数钢材更新成功');
    } else {
      addModuleSteel(steel);
      message.success('模数钢材添加成功');
    }

    setShowModuleModal(false);
    setEditingModuleSteel(null);
    moduleForm.resetFields();
  };

  // 删除模数钢材
  const handleDeleteModuleSteel = (steel: ModuleSteel) => {
    removeModuleSteel(steel.id);
    message.success('模数钢材删除成功');
  };

  // 编辑模数钢材
  const handleEditModuleSteel = (steel: ModuleSteel) => {
    setEditingModuleSteel(steel);
    moduleForm.setFieldsValue({
      name: steel.name,
      length: steel.length
    });
    setShowModuleModal(true);
  };

  // ==================== 优化执行 ====================
  
  const handleStartOptimization = async () => {
    if (designSteels.length === 0) {
      message.error('请先添加或上传设计钢材清单');
      return;
    }

    if (moduleSteels.length === 0) {
      message.error('请先添加模数钢材');
      return;
    }

    // 验证约束条件
    const constraintErrors = validateConstraints();
    if (constraintErrors.length > 0) {
      message.error(`约束条件错误: ${constraintErrors[0]}`);
      return;
    }

    // 验证焊接约束
    const weldingValidation = validateWeldingConstraint();
    if (!weldingValidation.isValid) {
      Modal.confirm({
        title: '焊接约束冲突',
        content: (
          <div>
            <p>{weldingValidation.message}</p>
            <p>是否继续优化？</p>
          </div>
        ),
        okText: '继续优化',
        cancelText: '取消',
        onOk() {
          executeOptimization();
        },
      });
      return;
    }

    executeOptimization();
  };

  const executeOptimization = async () => {
    try {
      console.log('=== 开始优化执行 ===');
      await startOptimization();
    } catch (error: any) {
      console.error('优化失败:', error);
      message.error('优化任务提交失败，请重试');
    }
  };

  // ==================== 约束条件管理 ====================
  
  const handleConstraintChange = (field: string, value: any) => {
    setConstraints({
      ...constraints,
      [field]: value
    });
  };

  const resetConstraints = () => {
    Modal.confirm({
      title: '重置约束条件',
      content: '确定要重置所有约束条件为默认值吗？',
      okText: '确定',
      cancelText: '取消',
      onOk() {
        // 🔧 修复：使用统一的默认约束配置，消除硬编码
        setConstraints({ ...DEFAULT_CONSTRAINTS });
        message.success('约束条件已重置');
      },
    });
  };

  const getConstraintDescription = (field: string) => {
    const descriptions = {
      wasteThreshold: '当余料长度小于此值时，将被视为废料无法再次利用',
      targetLossRate: '算法优化时的目标损耗率，作为参考值（不是强制要求）',
      timeLimit: '算法计算的最大允许时间，超时后返回当前最优解',
      maxWeldingSegments: '切割过程中允许的最大焊接次数，0次表示不允许焊接（V3新增功能）'
    };
    return descriptions[field as keyof typeof descriptions] || '';
  };

  const validateConstraints = () => {
    const errors: string[] = [];
    
    if (constraints.wasteThreshold < 0 || constraints.wasteThreshold > 1000) {
      errors.push('废料阈值必须在0-1000mm之间');
    }
    
    if (constraints.targetLossRate < 0 || constraints.targetLossRate > 50) {
      errors.push('目标损耗率必须在0-50%之间');
    }
    
    if (constraints.timeLimit < 1 || constraints.timeLimit > 300) {
      errors.push('计算时间限制必须在1-300秒之间');
    }
    
    if (constraints.maxWeldingSegments < 0 || constraints.maxWeldingSegments > 9) {
      errors.push('最大焊接次数必须在0-9次之间');
    }
    
    return errors;
  };

  const validateWeldingConstraint = () => {
    if (designSteels.length === 0 || moduleSteels.length === 0) {
      return { isValid: true, message: '' };
    }

    const maxModuleLength = Math.max(...moduleSteels.map(m => m.length));
    const conflictSteels = designSteels.filter(d => d.length > maxModuleLength);
    
    if (conflictSteels.length > 0 && constraints.maxWeldingSegments === 0) {
      const maxDesignLength = Math.max(...conflictSteels.map(s => s.length));
      const requiredTimes = Math.ceil(maxDesignLength / maxModuleLength) - 1;
      
      return {
        isValid: false,
        message: `有 ${conflictSteels.length} 根设计钢材长度超过最长模数钢材(${maxModuleLength}mm)，建议将最大焊接次数调整为 ${requiredTimes} 次以上`
      };
    }

    return { isValid: true, message: '' };
  };

  // ==================== 表格列定义 ====================
  
  const columns = [
    {
      title: '分组编号',
      dataIndex: 'displayId',
      key: 'displayId',
      render: (value: string) => value || '-',
      width: 100,
    },
    {
      title: '构件编号',
      dataIndex: 'componentNumber',
      key: 'componentNumber',
      render: (value: string) => value || '-',
    },
    {
      title: '规格',
      dataIndex: 'specification',
      key: 'specification',
      render: (value: string) => value || '-',
    },
    {
      title: '长度 (mm)',
      dataIndex: 'length',
      key: 'length',
      render: (value: number) => value?.toLocaleString() || '-',
    },
    {
      title: '数量',
      dataIndex: 'quantity',
      key: 'quantity',
      render: (value: number) => value?.toLocaleString() || '-',
    },
    {
      title: '截面积 (mm²)',
      dataIndex: 'crossSection',
      key: 'crossSection',
      render: (value: number) => value?.toLocaleString() || '-',
    },
    {
      title: '部件编号',
      dataIndex: 'partNumber',
      key: 'partNumber',
      render: (value: string) => value || '-',
    },
    {
      title: '操作',
      key: 'action',
      width: 120,
      render: (_: any, record: DesignSteel) => (
        <Space size="small">
          <Button 
            type="text" 
            icon={<EditOutlined />}
            onClick={() => handleEditDesignSteel(record)}
            size="small"
          >
            编辑
          </Button>
          <Popconfirm
            title="确定删除这条设计钢材吗？"
            onConfirm={() => handleDeleteDesignSteel(record)}
            okText="确定"
            cancelText="取消"
          >
            <Button
              type="text"
              danger
              icon={<DeleteOutlined />}
              size="small"
            >
              删除
            </Button>
          </Popconfirm>
        </Space>
      ),
    },
  ];

  const moduleColumns = [
    {
      title: '名称',
      dataIndex: 'name',
      key: 'name',
    },
    {
      title: '长度 (mm)',
      dataIndex: 'length',
      key: 'length',
      sorter: (a: ModuleSteel, b: ModuleSteel) => a.length - b.length,
    },
    {
      title: '操作',
      key: 'action',
      render: (_: any, record: ModuleSteel) => (
        <Space size="small">
          <Button
            type="text"
            icon={<EditOutlined />}
            onClick={() => handleEditModuleSteel(record)}
            size="small"
          >
            编辑
          </Button>
          <Popconfirm
            title="确定删除这个模数钢材吗？"
            onConfirm={() => handleDeleteModuleSteel(record)}
            okText="确定"
            cancelText="取消"
          >
          <Button 
            type="text" 
            danger 
            icon={<DeleteOutlined />}
              size="small"
          >
            删除
          </Button>
          </Popconfirm>
        </Space>
      ),
    },
  ];

  return (
    <PageContainer>
      <motion.div
        initial={{ opacity: 0, y: 20 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ duration: 0.5 }}
      >
        <Title level={2}>钢材优化配置</Title>
        <Text type="secondary">配置优化参数，管理钢材清单，开始智能优化</Text>
        
        <Divider />

        {/* 步骤1: 设计钢材管理 */}
        <StepCard 
          title={
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <Space>
                <FileExcelOutlined />
                设计钢材管理
                <Button
                  type="text"
                  icon={designCollapsed ? <RightOutlined /> : <DownOutlined />}
                  onClick={() => setDesignCollapsed(!designCollapsed)}
                  size="small"
                >
                  {designCollapsed ? '展开' : '折叠'}
                </Button>
            </Space>
              <Space>
                <Upload 
                  beforeUpload={handleDesignSteelUpload}
                  accept=".xlsx,.xls,.csv"
                  showUploadList={false}
                >
                  <Button icon={<UploadOutlined />} loading={uploading}>
                    上传Excel文件
                  </Button>
                </Upload>
                <Button
                  type="primary"
                  icon={<PlusOutlined />}
                  onClick={() => setShowDesignModal(true)}
                >
                  手动添加
                </Button>
              </Space>
                    </div>
          }
        >
          <div style={{ marginBottom: 12, display: 'flex', alignItems: 'center', gap: 12 }}>
            <Button
              danger
              disabled={selectedRowKeys.length === 0}
              onClick={handleBatchDelete}
            >
              批量删除{selectedRowKeys.length > 0 ? `（已选${selectedRowKeys.length}项）` : ''}
            </Button>
          </div>
          <Collapse activeKey={designCollapsed ? [] : ['1']} ghost>
            <Panel header="" key="1" showArrow={false}>
              <Table
                rowSelection={rowSelection}
                columns={columns}
                dataSource={designSteelsForDisplay}
                rowKey={(row) => row.id}
                pagination={{
                  pageSize: 10,
                  showSizeChanger: true,
                  showQuickJumper: true,
                  showTotal: (total) => `共 ${total} 条记录`,
                }}
                size="small"
                scroll={{ x: 900 }}
              />
            </Panel>
          </Collapse>
        </StepCard>

        {/* 步骤2: 模数钢材管理 */}
        <StepCard 
          title={
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <Space>
              <SettingOutlined />
                模数钢材管理
            </Space>
              <Button
              type="primary"
              icon={<PlusOutlined />}
                onClick={() => setShowModuleModal(true)}
            >
                添加模数钢材
              </Button>
            </div>
          }
        >
          <Table
            columns={moduleColumns}
            dataSource={moduleSteels}
            rowKey="id"
            pagination={false}
            size="small"
          />
        </StepCard>

        {/* 步骤3: 约束条件设置 */}
        <StepCard 
          title={
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <Space>
              <SettingOutlined />
                约束条件设置
            </Space>
              <Button
                type="primary"
                icon={<ReloadOutlined />}
                onClick={resetConstraints}
              >
                重置默认值
              </Button>
            </div>
          }
        >
          {/* 焊接约束冲突检查 */}
          {(() => {
            const validation = validateWeldingConstraint();
            if (!validation.isValid) {
              return (
                <Alert
                  message="焊接约束冲突"
                  description={validation.message}
                  type="warning"
                  showIcon
                  style={{ marginBottom: 16 }}
                />
              );
            }
            return null;
          })()}
            <Row gutter={[16, 16]}>
            <Col xs={24} sm={12} md={6}>
              <Card size="small" title="废料阈值 (mm)">
                <InputNumber
                  value={constraints.wasteThreshold}
                  onChange={(value) => handleConstraintChange('wasteThreshold', value || DEFAULT_CONSTRAINTS.wasteThreshold)}
                  min={0}
                  max={1000}
                  style={{ width: '100%' }}
                  placeholder={DEFAULT_CONSTRAINTS.wasteThreshold.toString()}
                />
                <Text type="secondary" style={{ fontSize: '12px', display: 'block', marginTop: '4px' }}>
                  {getConstraintDescription('wasteThreshold')}
                </Text>
              </Card>
            </Col>

            <Col xs={24} sm={12} md={6}>
              <Card size="small" title="目标损耗率 (%)">
                <InputNumber
                  value={constraints.targetLossRate}
                  onChange={(value) => handleConstraintChange('targetLossRate', value || DEFAULT_CONSTRAINTS.targetLossRate)}
                  min={0}
                  max={50}
                  step={0.1}
                  style={{ width: '100%' }}
                  placeholder={DEFAULT_CONSTRAINTS.targetLossRate.toString()}
                />
                <Text type="secondary" style={{ fontSize: '12px', display: 'block', marginTop: '4px' }}>
                  {getConstraintDescription('targetLossRate')}
                </Text>
              </Card>
              </Col>
              
            <Col xs={24} sm={12} md={6}>
              <Card size="small" title="计算时间限制 (秒)">
                <InputNumber
                  value={constraints.timeLimit}
                  onChange={(value) => handleConstraintChange('timeLimit', value || DEFAULT_CONSTRAINTS.timeLimit)}
                  min={1}
                  max={300}
                  style={{ width: '100%' }}
                  placeholder={DEFAULT_CONSTRAINTS.timeLimit.toString()}
                />
                <Text type="secondary" style={{ fontSize: '12px', display: 'block', marginTop: '4px' }}>
                  {getConstraintDescription('timeLimit')}
                </Text>
              </Card>
              </Col>
              
            <Col xs={24} sm={12} md={6}>
              <Card size="small" title={
                <span>
                  最大焊接次数 (次)
                  <Tag color="orange" style={{ marginLeft: 8 }}>V3新增</Tag>
                </span>
              }>
                <InputNumber
                  value={constraints.maxWeldingSegments}
                  onChange={(value) => handleConstraintChange('maxWeldingSegments', value || DEFAULT_CONSTRAINTS.maxWeldingSegments)}
                  min={0}
                  max={9}
                  style={{ width: '100%' }}
                  placeholder={DEFAULT_CONSTRAINTS.maxWeldingSegments.toString()}
                />
                <Text type="secondary" style={{ fontSize: '12px', display: 'block', marginTop: '4px' }}>
                  {getConstraintDescription('maxWeldingSegments')}
                </Text>
              </Card>
              </Col>
            </Row>
        </StepCard>

        {/* 步骤4: 开始优化 */}
        <StepCard 
          title={
            <Space>
              <PlayCircleOutlined />
              开始优化
            </Space>
          }
        >
          <div style={{ textAlign: 'center', padding: '40px 0' }}>
            {isOptimizing ? (
              <div>
                  <Progress 
                  type="circle" 
                    percent={Math.round(progress)} 
                    status="active"
                  style={{ marginBottom: 16 }}
                  />
                <div>
                  <Text>正在进行智能优化计算...</Text>
                </div>
              </div>
            ) : (
              <div>
                <ActionButton
                  type="primary"
                  size="large"
                  icon={<PlayCircleOutlined />}
                  onClick={handleStartOptimization}
                  disabled={designSteels.length === 0 || moduleSteels.length === 0}
                  style={{ 
                    height: '60px', 
                    fontSize: '18px', 
                    padding: '0 40px',
                    background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
                    border: 'none'
                  }}
                >
                  开始智能优化
                </ActionButton>
                <div style={{ marginTop: 16 }}>
                  <Text type="secondary">
                    已配置 {designSteels.length} 条设计钢材，{moduleSteels.length} 种模数钢材
                  </Text>
                </div>
              </div>
            )}
            
            {error && (
              <Alert
                message="优化错误"
                description={error}
                type="error"
                showIcon
                style={{ marginTop: 16, textAlign: 'left' }}
              />
            )}
          </div>
        </StepCard>

        {/* 设计钢材添加/编辑模态框 */}
      <Modal
          title={editingDesignSteel ? "编辑设计钢材" : "添加设计钢材"}
          open={showDesignModal}
          onCancel={() => {
            setShowDesignModal(false);
            setEditingDesignSteel(null);
            form.resetFields();
          }}
          footer={null}
        width={600}
      >
          <Form
            form={form}
            layout="vertical"
            onFinish={handleSaveDesignSteel}
          >
            <Row gutter={16}>
              <Col span={12}>
                <Form.Item
                  label="长度 (mm)"
                  name="length"
                  rules={[
                    { required: true, message: '请输入钢材长度' },
                    { type: 'number', min: 1, max: 50000, message: '长度必须在1-50000mm之间' }
                  ]}
                >
                  <InputNumber style={{ width: '100%' }} placeholder="如: 3000" />
                </Form.Item>
              </Col>
              <Col span={12}>
                <Form.Item
                  label="数量"
                  name="quantity"
                  rules={[
                    { required: true, message: '请输入数量' },
                    { type: 'number', min: 1, max: 10000, message: '数量必须在1-10000之间' }
                  ]}
                >
                  <InputNumber style={{ width: '100%' }} placeholder="如: 10" />
                </Form.Item>
              </Col>
            </Row>
            
            <Row gutter={16}>
              <Col span={12}>
                <Form.Item
                  label="截面面积 (mm²)"
                  name="crossSection"
                  rules={[
                    { required: true, message: '请输入截面面积' },
                    { type: 'number', min: 1, max: 100000, message: '截面面积必须在1-100000mm²之间' }
                  ]}
                >
                  <InputNumber style={{ width: '100%' }} placeholder="如: 2000" />
                </Form.Item>
              </Col>
              <Col span={12}>
                <Form.Item
                  label="构件编号"
                  name="componentNumber"
                >
                  <Input placeholder="如: GJ-001" />
                </Form.Item>
              </Col>
            </Row>
            
            <Row gutter={16}>
              <Col span={12}>
                <Form.Item
                  label="规格"
                  name="specification"
                >
                  <Input placeholder="如: HRB400" />
                </Form.Item>
              </Col>
              <Col span={12}>
                <Form.Item
                  label="零件号"
                  name="partNumber"
                >
                  <Input placeholder="如: P-001" />
                </Form.Item>
              </Col>
            </Row>

            <Form.Item style={{ marginBottom: 0, textAlign: 'right' }}>
              <Space>
                <Button onClick={() => {
                  setShowDesignModal(false);
                  setEditingDesignSteel(null);
                  form.resetFields();
                }}>
                  取消
                </Button>
                <Button type="primary" htmlType="submit">
                  {editingDesignSteel ? '更新' : '添加'}
                </Button>
              </Space>
            </Form.Item>
          </Form>
        </Modal>

        {/* 模数钢材添加/编辑模态框 */}
        <Modal
          title={editingModuleSteel ? "编辑模数钢材" : "添加模数钢材"}
          open={showModuleModal}
          onCancel={() => {
            setShowModuleModal(false);
            setEditingModuleSteel(null);
            moduleForm.resetFields();
          }}
          footer={null}
        >
          <Form
            form={moduleForm}
            layout="vertical"
            onFinish={handleSaveModuleSteel}
          >
            <Form.Item
              label="名称"
              name="name"
              rules={[{ required: true, message: '请输入模数钢材名称' }]}
            >
              <Input placeholder="如: 12米标准钢材" />
          </Form.Item>
          
            <Form.Item
              label="长度 (mm)"
              name="length"
              rules={[
                { required: true, message: '请输入长度' },
                { type: 'number', min: 1000, max: 50000, message: '长度必须在1000-50000mm之间' }
              ]}
            >
              <InputNumber style={{ width: '100%' }} placeholder="如: 12000" />
          </Form.Item>
          
            <Form.Item style={{ marginBottom: 0, textAlign: 'right' }}>
              <Space>
                <Button onClick={() => {
                  setShowModuleModal(false);
                  setEditingModuleSteel(null);
                  moduleForm.resetFields();
                }}>
                  取消
                </Button>
                <Button type="primary" htmlType="submit">
                  {editingModuleSteel ? '更新' : '添加'}
                </Button>
              </Space>
          </Form.Item>
        </Form>
      </Modal>
      </motion.div>
    </PageContainer>
  );
};

export default OptimizationPage;