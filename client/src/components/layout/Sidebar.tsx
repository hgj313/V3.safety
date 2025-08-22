import React from 'react';
import { Layout, Menu, Button, message } from 'antd';
import { 
  HomeOutlined, 
  BarChartOutlined, 
  TableOutlined, 
  HistoryOutlined, 
  SettingOutlined,
  MenuFoldOutlined,
  MenuUnfoldOutlined,
  BlockOutlined
} from '@ant-design/icons';
import { useLocation, useNavigate } from 'react-router-dom';
import styled from 'styled-components';
import { motion } from 'framer-motion';

const { Sider } = Layout;

interface SidebarProps {
  collapsed: boolean;
  onCollapse: (collapsed: boolean) => void;
}

const StyledSider = styled(Sider)`
  .ant-layout-sider-trigger {
    background: ${props => props.theme?.colors?.surface || '#fafafa'};
    border-top: 1px solid ${props => props.theme?.colors?.border || '#e8e8e8'};
    color: ${props => props.theme?.colors?.text?.primary || '#262626'};
    
    &:hover {
      background: ${props => (props.theme?.colors?.primary || '#1677ff') + '20'};
    }
  }

  &.ant-layout-sider-collapsed {
    .stone-logo-menu-item {
      margin: 16px auto !important;
      width: 48px;
      height: 48px;
      padding: 0 !important;
      border-radius: 12px;
      display: flex;
      justify-content: center;
      align-items: center;
      &::after {
        display: none;
      }
      &.ant-menu-item-selected,
      &:hover {
        background: #f0f2f5 !important;
      }
    }

    .logo-icon-div {
      width: 32px;
      height: 32px;
      background-image: url(/gse-stone-logo.png);
      background-size: contain;
      background-repeat: no-repeat;
      background-position: center;
    }
  }
`;

const LogoContainer = styled(motion.div)`
  height: 64px;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 0 16px; /* Adjusted padding */
  border-bottom: 1px solid ${props => props.theme?.colors?.border || '#e8e8e8'};
  background: ${props => props.theme?.colors?.surface || '#fafafa'};
  cursor: pointer;
`;

const LogoText = styled.h1`
  color: ${props => props.theme?.colors?.primary || '#1677ff'};
  font-size: 18px;
  font-weight: 600;
  margin: 0;
  white-space: nowrap;
  display: flex;
  align-items: center;
`;

const GradientText = styled.span`
  background: linear-gradient(90deg, #006400, #32CD32);
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
  background-clip: text;
  text-fill-color: transparent;
  position: relative;
  top: -2px;
`;



const CollapseButton = styled(Button)`
  position: absolute;
  top: 16px;
  right: -16px;
  z-index: 1000;
  width: 32px;
  height: 32px;
  border-radius: 50%;
  display: flex;
  align-items: center;
  justify-content: center;
  background: ${props => props.theme?.colors?.surface || '#fafafa'};
  border: 1px solid ${props => props.theme?.colors?.border || '#e8e8e8'};
  color: ${props => props.theme?.colors?.text?.primary || '#262626'};
  
  &:hover {
    background: ${props => (props.theme?.colors?.primary || '#1677ff') + '20'};
    border-color: ${props => props.theme?.colors?.primary || '#1677ff'};
  }
`;

const Sidebar: React.FC<SidebarProps> = ({ collapsed, onCollapse }) => {
  const location = useLocation();
  const navigate = useNavigate();

  const menuItems = [
    { key: '/', icon: <HomeOutlined />, label: '首页' },
    { key: '/optimization', icon: <BarChartOutlined />, label: '优化配置' },
    { key: '/results', icon: <TableOutlined />, label: '结果查看' },
    { key: '/history', icon: <HistoryOutlined />, label: '历史记录' },
    { key: '/settings', icon: <SettingOutlined />, label: '系统设置' },
  ];

  const handleMenuClick = ({ key }: { key: string }) => {
    navigate(key);
  };

  const handleStoneSystemClick = () => {
    message.info('开发中，敬请期待！', 2);
  };

  const systemMenuItems = [
    {
      key: 'stone-system',
      icon: <BlockOutlined />,
      label: '石材采购估算系统',
      onClick: handleStoneSystemClick,
    }
  ];

  return (
    <StyledSider
      trigger={null}
      collapsible
      collapsed={collapsed}
      width={280} /* Adjusted width */
      collapsedWidth={80}
      theme="light"
    >
      <LogoContainer
        initial={{ opacity: 0 }}
        animate={{ opacity: 1 }}
        transition={{ duration: 0.3 }}
        onClick={() => onCollapse(!collapsed)}
      >
        {collapsed ? (
          <LogoText>
            <GradientText>GSE</GradientText>
          </LogoText>
        ) : (
          <LogoText>
            <span style={{ fontWeight: 700, fontSize: '1.4em', marginRight: '8px', color: '#006400' }}>GSE</span>
            <GradientText>智能采购系统 V3</GradientText>
          </LogoText>
        )}
      </LogoContainer>

      <CollapseButton
        type="text"
        icon={collapsed ? <MenuUnfoldOutlined /> : <MenuFoldOutlined />}
        onClick={() => onCollapse(!collapsed)}
        size="small"
      />

      {collapsed ? (
        <Menu 
          theme="light" 
          mode="inline" 
          items={systemMenuItems}
          style={{ marginTop: 16, background: 'transparent', border: 'none' }}
        />
      ) : (
        <Menu
          mode="inline"
          selectedKeys={[location.pathname]}
          items={menuItems}
          onClick={handleMenuClick}
          style={{
            border: 'none',
            background: 'transparent',
            marginTop: 16,
          }}
        />
      )}
    </StyledSider>
  );
};

export default Sidebar; 