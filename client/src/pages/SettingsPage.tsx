import React, { useState } from 'react';
import { 
  Card, 
  Form, 
  Input, 
  Select, 
  Switch, 
  Button, 
  Row, 
  Col, 
  Divider, 
  message, 
  Tabs, 
  InputNumber, 
  Typography, 
  Space, 
  Alert,
  Table,
  Modal,
  Tag,
  Tooltip
} from 'antd';
import { 
  SettingOutlined, 
  SaveOutlined, 
  ReloadOutlined, 
  ExclamationCircleOutlined,
  InfoCircleOutlined,
  PlusOutlined,
  EditOutlined,
  DeleteOutlined
} from '@ant-design/icons';
import styled from 'styled-components';
import { motion } from 'framer-motion';

const { Title, Text, Paragraph } = Typography;
const { Option } = Select;
const { TextArea } = Input;
const { TabPane } = Tabs;

const PageContainer = styled.div`
  height: 100%;
  overflow: auto;
  padding: 0;
`;

const SettingsCard = styled(Card)`
  margin-bottom: 24px;
  border-radius: 12px;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.06);
`;

const SectionTitle = styled.div`
  display: flex;
  align-items: center;
  margin-bottom: 16px;
  
  .anticon {
    margin-right: 8px;
    color: ${props => props.theme?.colors?.primary || '#1677ff'};
  }
`;

const ActionButton = styled(Button)`
  border-radius: 8px;
  height: 40px;
  font-weight: 500;
`;

const SettingsPage: React.FC = () => {
  const [form] = Form.useForm();
  const [constraintForm] = Form.useForm();
  const [loading, setLoading] = useState(false);
  const [activeTab, setActiveTab] = useState('system');
  const [showConstraintModal, setShowConstraintModal] = useState(false);
  const [editingConstraint, setEditingConstraint] = useState<any>(null);
  
  // 系统约束条件数据
  const [systemConstraints, setSystemConstraints] = useState([
    {
      id: 1,
      name: '最小焊接长度',
      type: 'welding_length_min',
      value: 100,
      unit: 'mm',
      description: '钢材焊接时的最小长度限制',
      enabled: true,
      category: 'welding'
    },
    {
      id: 2,
      name: '最大焊接长度',
      type: 'welding_length_max',
      value: 6000,
      unit: 'mm',
      description: '钢材焊接时的最大长度限制',
      enabled: true,
      category: 'welding'
    },
    {
      id: 3,
      name: '最大损耗率',
      type: 'max_loss_rate',
      value: 5,
      unit: '%',
      description: '钢材加工时允许的最大损耗率',
      enabled: true,
      category: 'optimization'
    },
    {
      id: 4,
      name: '最小利用率',
      type: 'min_utilization',
      value: 85,
      unit: '%',
      description: '钢材利用的最小利用率要求',
      enabled: true,
      category: 'optimization'
    }
  ]);

  // 保存系统设置
  const handleSaveSettings = async () => {
    try {
      setLoading(true);
      const values = await form.validateFields();
      
      // 模拟保存到后端
      setTimeout(() => {
        console.log('保存的设置:', values); // 使用values变量
        message.success('设置已保存');
        setLoading(false);
      }, 1000);
      
    } catch (error) {
      message.error('请检查输入项');
      setLoading(false);
    }
  };

  // 重置设置
  const handleResetSettings = () => {
    Modal.confirm({
      title: '确认重置',
      icon: <ExclamationCircleOutlined />,
      content: '是否要重置所有设置为默认值？此操作不可恢复。',
      okText: '确认重置',
      cancelText: '取消',
      onOk() {
        form.resetFields();
        message.success('设置已重置');
      },
    });
  };

  // 添加约束条件
  const addConstraint = () => {
    setEditingConstraint(null);
    constraintForm.resetFields();
    setShowConstraintModal(true);
  };

  // 编辑约束条件
  const editConstraint = (constraint: any) => {
    setEditingConstraint(constraint);
    constraintForm.setFieldsValue(constraint);
    setShowConstraintModal(true);
  };

  // 删除约束条件
  const deleteConstraint = (id: number) => {
    Modal.confirm({
      title: '确认删除',
      icon: <ExclamationCircleOutlined />,
      content: '确定要删除这个约束条件吗？',
      okText: '确认删除',
      cancelText: '取消',
      okType: 'danger',
      onOk() {
        setSystemConstraints(prev => prev.filter(c => c.id !== id));
        message.success('约束条件已删除');
      },
    });
  };

  // 切换约束条件状态
  const toggleConstraint = (id: number) => {
    setSystemConstraints(prev => prev.map(c => 
      c.id === id ? { ...c, enabled: !c.enabled } : c
    ));
  };

  // 保存约束条件
  const handleSaveConstraint = async () => {
    try {
      const values = await constraintForm.validateFields();
      
      if (editingConstraint) {
        // 更新现有约束
        setSystemConstraints(prev => prev.map(c => 
          c.id === editingConstraint.id ? { ...c, ...values } : c
        ));
        message.success('约束条件已更新');
      } else {
        // 添加新约束
        const newConstraint = {
          id: Date.now(),
          ...values,
          enabled: true
        };
        setSystemConstraints(prev => [...prev, newConstraint]);
        message.success('约束条件已添加');
      }
      
      setShowConstraintModal(false);
    } catch (error) {
      message.error('请检查输入项');
    }
  };

  // 约束条件表格列配置
  const constraintColumns = [
    {
      title: '约束名称',
      dataIndex: 'name',
      key: 'name',
      render: (text: string, record: any) => (
        <Space direction="vertical" size={0}>
          <Text strong>{text}</Text>
          <Text type="secondary" style={{ fontSize: 12 }}>
            {record.description}
          </Text>
        </Space>
      ),
    },
    {
      title: '类型',
      dataIndex: 'type',
      key: 'type',
      render: (type: string) => (
        <Tag color="blue">{type}</Tag>
      ),
    },
    {
      title: '值',
      key: 'value',
      render: (_: any, record: any) => (
        <Text code>{record.value} {record.unit}</Text>
      ),
    },
    {
      title: '分类',
      dataIndex: 'category',
      key: 'category',
      render: (category: string) => {
        const categoryMap: Record<string, { color: string; text: string }> = {
          'welding': { color: 'orange', text: '焊接' },
          'optimization': { color: 'green', text: '优化' },
          'material': { color: 'purple', text: '材料' },
          'process': { color: 'cyan', text: '工艺' }
        };
        const cat = categoryMap[category] || { color: 'default', text: category };
        return <Tag color={cat.color}>{cat.text}</Tag>;
      },
    },
    {
      title: '状态',
      dataIndex: 'enabled',
      key: 'enabled',
      render: (enabled: boolean, record: any) => (
        <Switch 
          checked={enabled} 
          onChange={() => toggleConstraint(record.id)}
          size="small"
        />
      ),
    },
    {
      title: '操作',
      key: 'actions',
      render: (_: any, record: any) => (
        <Space>
          <Tooltip title="编辑">
            <Button 
              type="text" 
              icon={<EditOutlined />} 
              size="small"
              onClick={() => editConstraint(record)}
            />
          </Tooltip>
          <Tooltip title="删除">
            <Button 
              type="text" 
              danger 
              icon={<DeleteOutlined />} 
              size="small"
              onClick={() => deleteConstraint(record.id)}
            />
          </Tooltip>
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
        <Title level={2}>系统设置</Title>
        <Text type="secondary">配置系统参数、约束条件和优化策略</Text>
        
        <Divider />

        <Tabs activeKey={activeTab} onChange={setActiveTab} type="card">
          {/* 系统参数设置 */}
          <TabPane tab="系统参数" key="system">
            <Form form={form} layout="vertical" initialValues={{
              defaultStrategy: 'balanced',
              maxIterations: 1000,
              convergenceTolerance: 0.001,
              maxOptimizationTime: 300,
              enableParallelProcessing: true,
              logLevel: 'info',
              autoSaveResults: true,
              resultRetentionDays: 30
            }}>
              <Row gutter={[24, 24]}>
                <Col xs={24} lg={12}>
                  <SettingsCard title="优化引擎设置">
                    <SectionTitle>
                      <SettingOutlined />
                      <Text strong>核心参数</Text>
                    </SectionTitle>
                    
                    <Form.Item 
                      label="默认优化策略" 
                      name="defaultStrategy"
                      tooltip="系统默认使用的优化策略"
                    >
                      <Select>
                        <Option value="cost">成本优先</Option>
                        <Option value="balanced">平衡优化</Option>
                        <Option value="speed">速度优先</Option>
                        <Option value="quality">质量优先</Option>
                      </Select>
                    </Form.Item>
                    
                    <Form.Item 
                      label="最大迭代次数" 
                      name="maxIterations"
                      tooltip="优化算法的最大迭代次数"
                    >
                      <InputNumber min={100} max={50000} style={{ width: '100%' }} />
                    </Form.Item>
                    
                    <Form.Item 
                      label="收敛精度" 
                      name="convergenceTolerance"
                      tooltip="算法收敛的精度要求"
                    >
                      <InputNumber min={0.0001} max={0.1} step={0.0001} style={{ width: '100%' }} />
                    </Form.Item>
                    
                    <Form.Item 
                      label="最大优化时间 (秒)" 
                      name="maxOptimizationTime"
                      tooltip="单次优化任务的最大执行时间"
                    >
                      <InputNumber min={60} max={3600} style={{ width: '100%' }} />
                    </Form.Item>
                  </SettingsCard>
                </Col>
                
                <Col xs={24} lg={12}>
                  <SettingsCard title="系统行为设置">
                    <SectionTitle>
                      <InfoCircleOutlined />
                      <Text strong>运行参数</Text>
                    </SectionTitle>
                    
                    <Form.Item 
                      label="启用并行处理" 
                      name="enableParallelProcessing"
                      valuePropName="checked"
                    >
                      <Switch />
                    </Form.Item>
                    
                    <Form.Item 
                      label="日志级别" 
                      name="logLevel"
                      tooltip="系统日志的详细程度"
                    >
                      <Select>
                        <Option value="error">错误</Option>
                        <Option value="warn">警告</Option>
                        <Option value="info">信息</Option>
                        <Option value="debug">调试</Option>
                      </Select>
                    </Form.Item>
                    
                    <Form.Item 
                      label="自动保存结果" 
                      name="autoSaveResults"
                      valuePropName="checked"
                    >
                      <Switch />
                    </Form.Item>
                    
                    <Form.Item 
                      label="结果保留天数" 
                      name="resultRetentionDays"
                      tooltip="优化结果在系统中保留的天数"
                    >
                      <InputNumber min={1} max={365} style={{ width: '100%' }} />
                    </Form.Item>
                  </SettingsCard>
                </Col>
              </Row>
              
              <Row gutter={[24, 24]}>
                <Col span={24}>
                  <SettingsCard title="高级配置">
                    <Form.Item 
                      label="API 配置" 
                      name="apiConfig"
                      tooltip="后端API的配置信息"
                    >
                      <TextArea 
                        rows={4} 
                        placeholder="请输入JSON格式的API配置"
                        defaultValue='{\n  "baseUrl": "http://localhost:5000/api",\n  "timeout": 30000,\n  "retries": 3\n}'
                      />
                    </Form.Item>
                  </SettingsCard>
                </Col>
              </Row>
            </Form>
          </TabPane>

          {/* 约束条件设置 */}
          <TabPane tab="约束条件" key="constraints">
            <SettingsCard 
              title="系统约束条件"
              extra={
                <ActionButton
                  type="primary"
                  icon={<PlusOutlined />}
                  onClick={addConstraint}
                >
                  添加约束
                </ActionButton>
              }
            >
              <Alert
                message="约束条件说明"
                description="这些约束条件将作为系统默认值应用到所有优化任务中。您可以在具体的优化任务中覆盖这些设置。"
                type="info"
                showIcon
                style={{ marginBottom: 16 }}
              />
              
              <Table
                columns={constraintColumns}
                dataSource={systemConstraints}
                rowKey="id"
                pagination={false}
                size="small"
              />
            </SettingsCard>
          </TabPane>

          {/* 数据管理 */}
          <TabPane tab="数据管理" key="data">
            <Row gutter={[24, 24]}>
              <Col xs={24} lg={12}>
                <SettingsCard title="数据清理">
                  <Paragraph>
                    定期清理过期的优化结果和临时文件，保持系统运行效率。
                  </Paragraph>
                  
                  <Space direction="vertical" style={{ width: '100%' }}>
                    <ActionButton block>
                      清理过期结果
                    </ActionButton>
                    <ActionButton block>
                      清理临时文件
                    </ActionButton>
                    <ActionButton block danger>
                      清理所有历史数据
                    </ActionButton>
                  </Space>
                </SettingsCard>
              </Col>
              
              <Col xs={24} lg={12}>
                <SettingsCard title="数据备份">
                  <Paragraph>
                    备份系统配置和重要数据，确保数据安全。
                  </Paragraph>
                  
                  <Space direction="vertical" style={{ width: '100%' }}>
                    <ActionButton block type="primary">
                      导出系统配置
                    </ActionButton>
                    <ActionButton block>
                      导入系统配置
                    </ActionButton>
                    <ActionButton block>
                      备份历史数据
                    </ActionButton>
                  </Space>
                </SettingsCard>
              </Col>
            </Row>
          </TabPane>
        </Tabs>

        {/* 操作按钮 */}
        <Card style={{ marginTop: 24, textAlign: 'center' }}>
          <Space size="large">
            <ActionButton
              type="primary"
              icon={<SaveOutlined />}
              loading={loading}
              onClick={handleSaveSettings}
              size="large"
            >
              保存设置
            </ActionButton>
            
            <ActionButton
              icon={<ReloadOutlined />}
              onClick={handleResetSettings}
              size="large"
            >
              重置设置
            </ActionButton>
          </Space>
        </Card>
      </motion.div>

      {/* 约束条件编辑弹窗 */}
      <Modal
        title={editingConstraint ? "编辑约束条件" : "添加约束条件"}
        open={showConstraintModal}
        onCancel={() => setShowConstraintModal(false)}
        onOk={handleSaveConstraint}
        width={600}
        destroyOnClose
      >
        <Form form={constraintForm} layout="vertical">
          <Form.Item 
            label="约束名称" 
            name="name" 
            rules={[{ required: true, message: '请输入约束名称' }]}
          >
            <Input placeholder="输入约束名称" />
          </Form.Item>
          
          <Form.Item 
            label="约束类型" 
            name="type" 
            rules={[{ required: true, message: '请选择约束类型' }]}
          >
            <Select placeholder="选择约束类型">
              <Option value="welding_length_min">最小焊接长度</Option>
              <Option value="welding_length_max">最大焊接长度</Option>
              <Option value="max_loss_rate">最大损耗率</Option>
              <Option value="min_utilization">最小利用率</Option>
              <Option value="max_quantity">最大数量</Option>
              <Option value="min_quantity">最小数量</Option>
            </Select>
          </Form.Item>
          
          <Row gutter={16}>
            <Col span={16}>
              <Form.Item 
                label="约束值" 
                name="value" 
                rules={[{ required: true, message: '请输入约束值' }]}
              >
                <InputNumber style={{ width: '100%' }} placeholder="输入数值" />
              </Form.Item>
            </Col>
            
            <Col span={8}>
              <Form.Item 
                label="单位" 
                name="unit" 
                rules={[{ required: true, message: '请输入单位' }]}
              >
                <Select placeholder="选择单位">
                  <Option value="mm">毫米 (mm)</Option>
                  <Option value="cm">厘米 (cm)</Option>
                  <Option value="m">米 (m)</Option>
                  <Option value="%">百分比 (%)</Option>
                  <Option value="个">个</Option>
                  <Option value="kg">千克 (kg)</Option>
                </Select>
              </Form.Item>
            </Col>
          </Row>
          
          <Form.Item 
            label="分类" 
            name="category" 
            rules={[{ required: true, message: '请选择分类' }]}
          >
            <Select placeholder="选择约束分类">
              <Option value="welding">焊接</Option>
              <Option value="optimization">优化</Option>
              <Option value="material">材料</Option>
              <Option value="process">工艺</Option>
            </Select>
          </Form.Item>
          
          <Form.Item label="描述" name="description">
            <TextArea 
              rows={3} 
              placeholder="输入约束条件的详细描述"
            />
          </Form.Item>
        </Form>
      </Modal>
    </PageContainer>
  );
};

export default SettingsPage; 