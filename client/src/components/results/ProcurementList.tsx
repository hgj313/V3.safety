import React from 'react';
import { Card, Table, Alert, Typography, Space } from 'antd';
import { ModuleUsageItem } from '../../hooks/useOptimizationResults';
import { formatNumber } from '../../utils/steelUtils';

const { Text } = Typography;

interface ProcurementListProps {
  moduleUsageStats: {
    sortedStats: ModuleUsageItem[];
    groupKeyTotals: Record<string, { count: number; totalLength: number }>;
    grandTotal: { count: number; totalLength: number };
  };
}

const ProcurementList: React.FC<ProcurementListProps> = ({
  moduleUsageStats
}) => {
  // 表格列定义
  const moduleStatsColumns = [
    {
      title: '规格',
      dataIndex: 'specification',
      key: 'specification',
      render: (spec: string, record: ModuleUsageItem) => (
        <Text strong={record.isSubtotal || record.isTotal}>
          {record.isSubtotal || record.isTotal ? (
            <span style={{ fontWeight: 'bold' }}>{spec}</span>
          ) : (
            <span style={{ 
              backgroundColor: '#f0f2f5', 
              padding: '2px 6px', 
              borderRadius: '4px',
              fontSize: '12px'
            }}>
              {spec}
            </span>
          )}
        </Text>
      )
    },
    {
      title: '长度(mm)',
      dataIndex: 'length',
      key: 'length',
      render: (length: number | string) => 
        typeof length === 'number' ? formatNumber(length, 0) : length
    },
    {
      title: '数量(根)',
      dataIndex: 'count',
      key: 'count',
      render: (count: number, record: ModuleUsageItem) => (
        <Text strong={record.isSubtotal || record.isTotal}>
          {count}
        </Text>
      )
    },
    {
      title: '总长度(mm)',
      dataIndex: 'totalLength',
      key: 'totalLength',
      render: (totalLength: number, record: ModuleUsageItem) => (
        <Text strong={record.isSubtotal || record.isTotal}>
          {formatNumber(totalLength, 0)}
        </Text>
      )
    }
  ];

  // 构建表格数据
  const tableData: ModuleUsageItem[] = [
    // 详细数据
    ...moduleUsageStats.sortedStats,
    // 小计数据
    ...Object.entries(moduleUsageStats.groupKeyTotals).map(([groupKey, totals]) => {
      const [specification, crossSectionStr] = groupKey.split('_');
      return {
        key: `subtotal-${groupKey}`,
        specification: `${specification}(${crossSectionStr}mm²) 小计`,
        length: '-',
        count: totals.count,
        totalLength: totals.totalLength,
        isSubtotal: true
      } as ModuleUsageItem;
    }),
    // 总计数据
    {
      key: 'total',
      specification: '总计',
      length: '-',
      count: moduleUsageStats.grandTotal.count,
      totalLength: moduleUsageStats.grandTotal.totalLength,
      isTotal: true
    }
  ];

  return (
    <Card title="采购统计" size="small">
      <Alert
        type="info"
        message="采购指导"
        description="采购人员可直接按规格名称采购对应数量的钢材。"
        style={{ marginBottom: 16 }}
        showIcon
      />
      <Table
        columns={moduleStatsColumns}
        dataSource={tableData}
        rowKey="key"
        pagination={false}
        size="small"
        rowClassName={(record: ModuleUsageItem) => {
          if (record.isTotal) return 'module-stats-total-row';
          if (record.isSubtotal) return 'module-stats-subtotal-row';
          return '';
        }}
        title={() => (
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <Text strong>采购清单</Text>
            <Space>
              <Text type="secondary">总计: {moduleUsageStats.grandTotal.count}根</Text>
              <Text type="secondary">总长: {formatNumber(moduleUsageStats.grandTotal.totalLength, 0)}mm</Text>
            </Space>
          </div>
        )}
      />
      
      <div style={{ marginTop: 16 }}>
        <Alert
          type="success"
          message="采购优势"
          description={
            <ul style={{ margin: 0, paddingLeft: 20 }}>
              <li><strong>🎉直接按规格采购：</strong>直接得出各个规格的采购数量，无需计算</li>
              <li><strong>📝后续计划更新：</strong>增加采购重量和单价，便于直观查看成本</li>
            </ul>
          }
          showIcon
        />
      </div>

      <style>{`
        .module-stats-total-row {
          background-color: #f0f2f5 !important;
          font-weight: bold !important;
        }
        .module-stats-subtotal-row {
          background-color: #fafafa !important;
          font-weight: 500 !important;
        }
      `}</style>
    </Card>
  );
};

export default React.memo(ProcurementList); 