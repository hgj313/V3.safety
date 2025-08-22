import React from 'react';
import { Card, Table, Tag, Alert, Typography } from 'antd';
import { CheckCircleOutlined, ExclamationCircleOutlined } from '@ant-design/icons';
import { RequirementValidationItem } from '../../hooks/useOptimizationResults';
import { formatNumber } from '../../utils/steelUtils';

const { Text } = Typography;

interface RequirementsValidationProps {
  requirementValidation: RequirementValidationItem[];
  isAllRequirementsSatisfied: boolean;
}

const RequirementsValidation: React.FC<RequirementsValidationProps> = ({
  requirementValidation,
  isAllRequirementsSatisfied
}) => {
  // 表格列定义
  const requirementColumns = [
    {
      title: '编号',
      dataIndex: 'id',
      key: 'id',
      render: (id: string) => <Text strong>{id}</Text>
    },
    {
      title: '规格',
      dataIndex: 'specification',
      key: 'specification',
      render: (spec: string) => (
        <Tag color="blue" style={{ fontSize: '13px', padding: '4px 8px' }}>
          {spec}
        </Tag>
      )
    },
    {
      title: '长度(mm)',
      dataIndex: 'length',
      key: 'length',
      render: (length: number) => formatNumber(length, 0)
    },
    {
      title: '需求数量',
      dataIndex: 'quantity',
      key: 'quantity'
    },
    {
      title: '生产数量',
      dataIndex: 'produced',
      key: 'produced',
      render: (produced: number, record: RequirementValidationItem) => (
        <Text type={record.satisfied ? 'success' : 'danger'}>
          {produced}
        </Text>
      )
    },
    {
      title: '满足状态',
      dataIndex: 'satisfied',
      key: 'satisfied',
      render: (satisfied: boolean) => (
        <Tag color={satisfied ? 'green' : 'red'} icon={satisfied ? <CheckCircleOutlined /> : <ExclamationCircleOutlined />}>
          {satisfied ? '满足' : '不满足'}
        </Tag>
      )
    },
    {
      title: '差值',
      dataIndex: 'difference',
      key: 'difference',
      render: (diff: number) => (
        <Text type={diff === 0 ? 'secondary' : diff > 0 ? 'success' : 'danger'}>
          {diff > 0 ? `+${diff}` : diff}
        </Text>
      )
    }
  ];

  return (
    <>
      <Card title="需求验证表" size="small">
        <Alert
          type="info"
          message="注意事项："
          description="如果需求未全部满足，请检查是否在可焊接次数=1时，存在钢材长度超过最大模数钢材的长度。"
          style={{ marginBottom: 16 }}
          showIcon
        />
        <Table
          columns={requirementColumns}
          dataSource={requirementValidation}
          rowKey="key"
          pagination={{ pageSize: 10 }}
          size="small"
          summary={(pageData) => {
            const totalRequired = pageData.reduce((sum, record) => sum + record.quantity, 0);
            const totalProduced = pageData.reduce((sum, record) => sum + record.produced, 0);
            const unsatisfiedCount = pageData.filter(record => !record.satisfied).length;
            
            return (
              <Table.Summary fixed>
                <Table.Summary.Row>
                  <Table.Summary.Cell index={0}>
                    <Text strong>统计</Text>
                  </Table.Summary.Cell>
                  <Table.Summary.Cell index={1}>-</Table.Summary.Cell>
                  <Table.Summary.Cell index={2}>-</Table.Summary.Cell>
                  <Table.Summary.Cell index={3}>
                    <Text strong>{totalRequired}</Text>
                  </Table.Summary.Cell>
                  <Table.Summary.Cell index={4}>
                    <Text strong style={{ color: '#1890ff' }}>{totalProduced}</Text>
                  </Table.Summary.Cell>
                  <Table.Summary.Cell index={5}>
                    <Tag color={unsatisfiedCount === 0 ? 'green' : 'red'}>
                      {unsatisfiedCount === 0 ? '全部满足' : `${unsatisfiedCount}项未满足`}
                    </Tag>
                  </Table.Summary.Cell>
                  <Table.Summary.Cell index={6}>
                    <Text strong style={{ 
                      color: (totalProduced - totalRequired) === 0 ? '#52c41a' : 
                            (totalProduced - totalRequired) > 0 ? '#1890ff' : '#ff4d4f'
                    }}>
                      {totalProduced - totalRequired > 0 ? 
                        `+${totalProduced - totalRequired}` : 
                        totalProduced - totalRequired}
                    </Text>
                  </Table.Summary.Cell>
                </Table.Summary.Row>
              </Table.Summary>
            );
          }}
        />
      </Card>
      
      <div style={{ marginTop: 16 }}>
        <Alert
          type={isAllRequirementsSatisfied ? 'success' : 'warning'}
          message={isAllRequirementsSatisfied ? '需求验证通过' : '需求验证异常'}
          description={
            isAllRequirementsSatisfied ? (
              <ul style={{ margin: 0, paddingLeft: 20 }}>
                <li>所有设计钢材需求已完全满足</li>
                <li>V3规格化设计确保精确匹配</li>
                <li>生产计划可直接执行</li>
              </ul>
            ) : (
              <ul style={{ margin: 0, paddingLeft: 20 }}>
                <li style={{ color: '#ff4d4f' }}>部分设计钢材需求未满足</li>
                <li>建议检查规格匹配情况</li>
                <li>或调整优化参数重新计算</li>
              </ul>
            )
          }
        />
      </div>
    </>
  );
};

export default React.memo(RequirementsValidation); 