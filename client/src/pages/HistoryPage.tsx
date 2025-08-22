import React from 'react';
import { Card } from 'antd';
import styled from 'styled-components';

const PageContainer = styled.div`
  height: 100%;
  overflow: auto;
`;

const HistoryPage: React.FC = () => {
  return (
    <PageContainer>
      <Card title="历史记录" bordered={false}>
        <p>历史记录页面开发中...</p>
      </Card>
    </PageContainer>
  );
};

export default HistoryPage; 