import React from 'react';
import { Alert, Collapse, Table, Tag, Typography } from 'antd';
import { formatNumber } from '../../utils/steelUtils';

const { Text } = Typography;
const { Panel } = Collapse;

interface CuttingPlansTableProps {
  regroupedResults: Record<string, any>;
  designIdToDisplayIdMap: Map<string, string>;
}

const CuttingPlansTable: React.FC<CuttingPlansTableProps> = ({
  regroupedResults,
  designIdToDisplayIdMap
}) => {

  // 为每个规格组合生成一个唯一的余料ID映射
  const remainderDisplayIdMaps = React.useMemo(() => {
    const maps: Record<string, Map<string, string>> = {};
    Object.keys(regroupedResults).forEach(spec => {
      maps[spec] = new Map<string, string>();
      let counter = 1;
      regroupedResults[spec].cuttingPlans?.forEach((plan: any) => {
        plan.newRemainders?.forEach((remainder: any) => {
          if (!maps[spec].has(remainder.id)) {
            maps[spec].set(remainder.id, `a${counter++}`);
          }
        });
      });
    });
    return maps;
  }, [regroupedResults]);

  // 表格列定义
  const getCuttingColumns = (specification: string) => [
    {
      title: '原料',
      dataIndex: 'sourceType',
      key: 'sourceType',
      render: (type: string, record: any) => {
        const style = {
          padding: '2px 8px',
          borderRadius: '4px',
          border: '1px solid'
        };

        if (type === 'module') {
          return (
            <span style={{
              ...style,
              backgroundColor: '#E6F7FF',
              borderColor: '#91D5FF',
              color: '#096DD9'
            }}>
              模数钢材
            </span>
          );
        }
        if (type === 'remainder') {
          const remainderDisplayIdMap = remainderDisplayIdMaps[specification];
          const usedRemainderIds = record.usedRemainders?.map((r: any) => 
            remainderDisplayIdMap.get(r.id) || r.id
          ).join(' + ') || '余料';

          return (
            <span style={{
              ...style,
              backgroundColor: '#F6FFED',
              borderColor: '#B7EB8F',
              color: '#389E0D'
            }} title="该余料被成功再利用">
              余料 {usedRemainderIds}
            </span>
          );
        }
        return <Tag>{type}</Tag>;
      }
    },
    {
      title: '原料长度 (mm)',
      dataIndex: 'sourceLength',
      key: 'sourceLength',
      render: (text: number) => formatNumber(text, 0),
    },
    {
      title: '切割详情',
      dataIndex: 'cuts',
      key: 'cuts',
      render: (cuts: any[]) => (
        <div>
          {cuts?.map((cut: any, index: number) => (
            <div key={index}>
              <Text>{`${cut.displayId || designIdToDisplayIdMap.get(cut.designId) || cut.designId}: ${formatNumber(cut.length, 0)}mm × ${cut.quantity}件`}</Text>
            </div>
          ))}
        </div>
      )
    },
    {
      title: '新余料 (mm)',
      dataIndex: 'newRemainders',
      key: 'newRemainders',
      render: (newRemainders: any[], record: any) => {
        if (!newRemainders || newRemainders.length === 0) {
          return <Text type="secondary">0</Text>;
        }
        
        // 过滤出非废料的余料
        const displayRemainders = newRemainders.filter(r => r.type !== 'waste');

        if (displayRemainders.length === 0) {
          return <Text type="secondary">0</Text>;
        }

        return (
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
            {displayRemainders.map((r, idx) => {
              const displayId = remainderDisplayIdMaps[specification]?.get(r.id) || r.id;
              // 最终真余料：红色背景
              if (r.type === 'real') {
                return (
                  <span
                    key={idx}
            style={{ 
                      background: '#fff2f0',
                      border: '1px solid #ffccc7',
                      color: '#cf1322',
                      fontWeight: 'bold',
                      borderRadius: 4,
                      padding: '0 8px',
                      display: 'inline-block',
                    }}
                    title={`真余料: ${r.id}`}
                  >
                    {`${displayId}: ${formatNumber(r.length, 0)}`}
                  </span>
                );
              }
              // 其他所有中间状态（伪余料、待定）都只显示长度，无特殊背景
              else {
                return (
                  <span key={idx} title={`中间余料: ${r.id}`}>
                    {`${displayId}: ${formatNumber(r.length, 0)}`}
                  </span>
                );
              }
            })}
          </div>
        );
      },
    },
    {
      title: '废料 (mm)',
      dataIndex: 'waste',
      key: 'waste',
      render: (waste: number, record: any) => {
        // 汇总所有废料：plan.waste + newRemainders中标记为waste的
        let totalWaste = waste || 0;
        if (record.newRemainders) {
          const wasteFromRemainders = record.newRemainders
            .filter((r: any) => r.type === 'waste')
            .reduce((sum: number, r: any) => sum + (r.length || 0), 0);
          totalWaste += wasteFromRemainders;
        }
        
        if (totalWaste <= 0) {
          return <Text type="secondary">0</Text>;
        }

        return (
          <Text style={{ color: '#D46B08', fontWeight: 'bold' }}>
            {formatNumber(totalWaste, 0)}
        </Text>
        );
      },
    },
  ];

  return (
    <>
      <Alert
        type="info"
        message="余料概念说明"
        description={
          <div style={{ lineHeight: 1.8 }}>
            • 伪余料：还在加工流程里，下一刀可能就用掉<br/>
            • 真余料：最终剩下的合格边料，可以收回仓库下次用（纳入单次项目损耗）<br/>
            • 废料：太短的边角料，只能当废料处理
          </div>
        }
        style={{ marginBottom: 16 }}
        showIcon
      />
      
      {Object.entries(regroupedResults).map(([specification, solution]: [string, any]) => {
        console.log('检查 solution:', { specification, solution });
        // 分组损耗率直接使用后端统计的 totalMaterial，避免前端重复计算
        const groupLossRate = solution.lossRate ?? 0;
          
        return (
          <Collapse key={specification} style={{ marginBottom: 16 }}>
            <Panel
              header={
                <div style={{ display: 'flex', alignItems: 'center', gap: 16 }}>
                  <Tag color="blue" style={{ fontSize: '14px', padding: '4px 8px' }}>
                    {specification}
                  </Tag>
                  <Text>
                    真实损耗率: {groupLossRate.toFixed(2)}%
                  </Text>
                </div>
              }
              key={specification}
            >
              <Table
                columns={getCuttingColumns(specification)}
                dataSource={solution.cuttingPlans?.map((plan: any, index: number) => ({
                  ...plan,
                  key: `${specification}-${index}`,
                  cuts: plan.cuts?.map((cut: any) => ({
                    ...cut,
                    displayId: cut.displayId || designIdToDisplayIdMap.get(cut.designId) || cut.designId
                  })) || []
                })) || []}
                rowKey="key"
                pagination={false}
                size="small"
                title={() => (
                  <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                    <Text strong>切割计划明细</Text>
                  </div>
                )}
              />
            </Panel>
          </Collapse>
        );
      })}
    </>
  );
};

export default React.memo(CuttingPlansTable); 