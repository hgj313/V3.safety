import React from 'react';
import { Card, Row, Col, Statistic, Button, List, Timeline, Space } from 'antd';
import { 
  BarChartOutlined, 
  DollarOutlined, 
  ThunderboltOutlined, 
  CheckCircleOutlined,
  PlayCircleOutlined,
  FileTextOutlined
} from '@ant-design/icons';
import styled from 'styled-components';
import { motion } from 'framer-motion';
import { useNavigate } from 'react-router-dom';

const PageContainer = styled.div`
  height: 100%;
  overflow: auto;
  background-image: url('https://www.transparenttextures.com/patterns/brushed-metal.png');
  background-color: ${props => props.theme?.colors?.background || '#ffffff'};
`;

const WelcomeCard = styled(Card)`
  margin-bottom: 24px;
  background: linear-gradient(135deg, ${props => (props.theme?.colors?.primary || '#1677ff') + '20'}, ${props => (props.theme?.colors?.secondary || '#722ed1') + '20'});
  border: 1px solid ${props => props.theme?.colors?.border || '#e8e8e8'};
`;

const WelcomeTitle = styled.h1`
  font-size: 32px;
  font-weight: 600;
  margin: 0 0 8px 0;
  color: ${props => props.theme?.colors?.text?.primary || '#262626'};
  background: -webkit-linear-gradient(45deg, ${props => props.theme?.colors.primary}, ${props => props.theme?.colors?.secondary});
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
`;

const WelcomeSubtitle = styled.p`
  font-size: 16px;
  color: ${props => props.theme?.colors?.text?.secondary || '#8c8c8c'};
  margin: 0 0 24px 0;
`;

const QuickActionButton = styled(Button)`
  height: 48px;
  border-radius: 8px;
  font-weight: 500;
`;

const StatsCard = styled(Card)`
  .ant-card-body {
    padding: 24px;
  }
`;

const ActivityCard = styled(Card)`
  height: 400px;
  
  .ant-card-body {
    height: calc(100% - 57px);
    overflow: auto;
  }
`;

const HomePage: React.FC = () => {
  const navigate = useNavigate();

  // 模拟数据
  const stats = [
    {
      title: '总优化次数',
      value: '-----',
      icon: <BarChartOutlined />,
      color: '#1677ff',
      prefix: undefined,
      suffix: undefined,
      isPlaceholder: false,
    },
    {
      title: '节省成本',
      value: '待开发',
      icon: <DollarOutlined />,
      color: '#52c41a',
      prefix: undefined,
      suffix: undefined,
      isPlaceholder: true,
    },
    {
      title: '优化效率',
      value: '----',
      suffix: '%',
      icon: <ThunderboltOutlined />,
      color: '#faad14',
      prefix: undefined,
      isPlaceholder: false,
    },
    {
      title: '成功率',
      value: '----',
      suffix: '%',
      icon: <CheckCircleOutlined />,
      color: '#722ed1',
      prefix: undefined,
      isPlaceholder: false,
    },
  ];

  const recentActivities = [
    {
      title: '完成西耳墙项目优化',
      description: '节省成本 ¥......',
      time: '----',
      type: 'success',
    },
    {
      title: '开始新的优化任务',
      description: '钢材清单已上传',
      time: '----',
      type: 'processing',
    },
    {
      title: '导出优化报告',
      description: 'PDF 报告已生成',
      time: '----',
      type: 'default',
    },
    {
      title: '系统更新完成',
      description: '优化算法 V3.0 已上线',
      time: '----',
      type: 'success',
    },
  ];

  const quickActions = [
    {
      key: 'optimize',
      title: '开始优化',
      description: '上传钢材清单，开始智能优化',
      icon: <PlayCircleOutlined />,
      color: '#1677ff',
      path: '/optimization',
    },
    {
      key: 'results',
      title: '查看结果',
      description: '查看最新的优化结果',
      icon: <BarChartOutlined />,
      color: '#52c41a',
      path: '/results',
    },
    {
      key: 'history',
      title: '历史记录',
      description: '查看过往的优化历史',
      icon: <FileTextOutlined />,
      color: '#faad14',
      path: '/history',
    },
  ];

  return (
    <PageContainer>
      <motion.div
        initial={{ opacity: 0, y: 20 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ duration: 0.5 }}
      >
        <WelcomeCard>
          <Row align="middle" justify="space-between">
            <Col>
              <WelcomeTitle> GSE智能钢材采购系统 </WelcomeTitle>
              <WelcomeSubtitle>
                智能化钢材采购方案，最大化利用率，最小化成本浪费
              </WelcomeSubtitle>
            </Col>
            <Col>
              <Space>
                {quickActions.map((action, index) => (
                  <motion.div
                    key={action.key}
                    initial={{ opacity: 0, x: 20 }}
                    animate={{ opacity: 1, x: 0 }}
                    transition={{ duration: 0.3, delay: index * 0.1 }}
                  >
                    <QuickActionButton
                      type="primary"
                      icon={action.icon}
                      onClick={() => navigate(action.path)}
                      style={{ backgroundColor: action.color, borderColor: action.color }}
                    >
                      {action.title}
                    </QuickActionButton>
                  </motion.div>
                ))}
              </Space>
            </Col>
          </Row>
        </WelcomeCard>
      </motion.div>

      <Row gutter={[24, 24]}>
        {/* 统计数据 */}
        {stats.map((stat, index) => (
          <Col xs={24} sm={12} lg={6} key={index}>
            <motion.div
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ duration: 0.5, delay: index * 0.1 }}
            >
              <StatsCard>
                <Statistic
                  title={stat.title}
                  value={stat.isPlaceholder ? undefined : stat.value}
                  prefix={stat.prefix}
                  suffix={stat.suffix}
                  valueStyle={{ 
                    color: stat.color,
                    fontSize: stat.isPlaceholder ? '14px' : undefined,
                    fontStyle: stat.isPlaceholder ? 'italic' : undefined
                  }}
                  formatter={stat.isPlaceholder ? () => (
                    <span style={{ color: '#8c8c8c' }}>
                      {stat.value}
                      <br />
                      <small style={{ fontSize: '12px' }}>功能开发中...</small>
                    </span>
                  ) : undefined}
                />
              </StatsCard>
            </motion.div>
          </Col>
        ))}
      </Row>

      <Row gutter={[24, 24]} style={{ marginTop: 24 }}>
        {/* 最近活动 */}
        <Col xs={24} lg={12}>
          <motion.div
            initial={{ opacity: 0, x: -20 }}
            animate={{ opacity: 1, x: 0 }}
            transition={{ duration: 0.5, delay: 0.4 }}
          >
            <ActivityCard title="最近活动" bordered={false}>
              <Timeline
                items={recentActivities.map((activity, index) => ({
                  children: (
                    <div>
                      <div style={{ fontWeight: 500, marginBottom: 4 }}>
                        {activity.title}
                      </div>
                      <div style={{ 
                        color: 'rgba(0, 0, 0, 0.45)', 
                        fontSize: 14,
                        marginBottom: 4 
                      }}>
                        {activity.description}
                      </div>
                      <div style={{ 
                        color: 'rgba(0, 0, 0, 0.25)', 
                        fontSize: 12 
                      }}>
                        {activity.time}
                      </div>
                    </div>
                  ),
                  color: activity.type === 'success' ? 'green' : 
                         activity.type === 'processing' ? 'blue' : 'gray',
                }))}
              />
            </ActivityCard>
          </motion.div>
        </Col>

        {/* 快速操作 */}
        <Col xs={24} lg={12}>
          <motion.div
            initial={{ opacity: 0, x: 20 }}
            animate={{ opacity: 1, x: 0 }}
            transition={{ duration: 0.5, delay: 0.5 }}
          >
            <Card title="快速操作" bordered={false}>
              <List
                dataSource={quickActions}
                renderItem={(item) => (
                  <List.Item
                    actions={[
                      <Button 
                        type="link" 
                        onClick={() => navigate(item.path)}
                      >
                        开始
                      </Button>
                    ]}
                  >
                    <List.Item.Meta
                      avatar={<div style={{ color: item.color, fontSize: 24 }}>{item.icon}</div>}
                      title={item.title}
                      description={item.description}
                    />
                  </List.Item>
                )}
              />
            </Card>
          </motion.div>
        </Col>
      </Row>
    </PageContainer>
  );
};

export default HomePage; 