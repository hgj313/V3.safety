import React, { useState } from 'react';
import { Layout, Switch, Button, Space, Tooltip, Badge, Dropdown, List, Typography, Empty } from 'antd';
import { 
  SunOutlined, 
  MoonOutlined, 
  BellOutlined,
  FullscreenOutlined,
  FullscreenExitOutlined,
  CheckCircleOutlined,
  ExclamationCircleOutlined,
  InfoCircleOutlined
} from '@ant-design/icons';
import styled from 'styled-components';
import { motion } from 'framer-motion';

const { Header: AntHeader } = Layout;
const { Text } = Typography;

interface HeaderProps {
  isDarkMode: boolean;
  onThemeToggle: () => void;
  sidebarCollapsed: boolean;
  onSidebarToggle: () => void;
}

// 模拟通知数据
interface NotificationItem {
  id: string;
  type: 'success' | 'warning' | 'info';
  title: string;
  content: string;
  time: string;
  read: boolean;
}

const mockNotifications: NotificationItem[] = [
  {
    id: '1',
    type: 'success',
    title: '优化完成',
    content: '西耳墙项目优化已完成，节省成本 ¥12,300',
    time: '2分钟前',
    read: false
  },
  {
    id: '2',
    type: 'info',
    title: '系统更新',
    content: '钢材采购优化系统已更新至 V3.0',
    time: '1小时前',
    read: false
  },
  {
    id: '3',
    type: 'warning',
    title: '数据提醒',
    content: '请及时更新钢材价格数据以确保优化准确性',
    time: '3小时前',
    read: true
  }
];

const StyledHeader = styled(AntHeader)`
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 0 24px;
  height: 64px;
  line-height: 64px;
`;

const LeftSection = styled.div`
  display: flex;
  align-items: center;
  gap: 16px;
`;

const RightSection = styled.div`
  display: flex;
  align-items: center;
  gap: 16px;
`;

const ThemeSwitch = styled(Switch)`
  .ant-switch-handle {
    display: flex;
    align-items: center;
    justify-content: center;
    
    &::before {
      display: none;
    }
  }
`;

const ActionButton = styled(Button)`
  border: none;
  background: transparent;
  color: ${props => props.theme?.colors?.text?.primary || '#262626'};
  display: flex;
  align-items: center;
  justify-content: center;
  
  &:hover {
    background: ${props => (props.theme?.colors?.primary || '#1677ff') + '20'};
    color: ${props => props.theme?.colors?.primary || '#1677ff'};
  }
`;

const NotificationDropdown = styled.div`
  width: 350px;
  max-height: 400px;
  overflow-y: auto;
`;

const NotificationHeader = styled.div`
  padding: 12px 16px;
  border-bottom: 1px solid #f0f0f0;
  display: flex;
  justify-content: space-between;
  align-items: center;
`;

const Header: React.FC<HeaderProps> = ({ 
  isDarkMode, 
  onThemeToggle, 
  sidebarCollapsed, 
  onSidebarToggle 
}) => {
  const [isFullscreen, setIsFullscreen] = React.useState(false);
  const [notifications, setNotifications] = useState<NotificationItem[]>(mockNotifications);

  const toggleFullscreen = () => {
    if (!document.fullscreenElement) {
      document.documentElement.requestFullscreen();
      setIsFullscreen(true);
    } else {
      document.exitFullscreen();
      setIsFullscreen(false);
    }
  };

  const unreadCount = notifications.filter(n => !n.read).length;

  const handleMarkAllRead = () => {
    setNotifications(prev => prev.map(n => ({ ...n, read: true })));
  };

  const handleNotificationClick = (id: string) => {
    setNotifications(prev => 
      prev.map(n => n.id === id ? { ...n, read: true } : n)
    );
  };

  const getNotificationIcon = (type: string) => {
    switch (type) {
      case 'success':
        return <CheckCircleOutlined style={{ color: '#52c41a' }} />;
      case 'warning':
        return <ExclamationCircleOutlined style={{ color: '#faad14' }} />;
      case 'info':
      default:
        return <InfoCircleOutlined style={{ color: '#1677ff' }} />;
    }
  };

  const notificationMenu = (
    <NotificationDropdown>
      <NotificationHeader>
        <Text strong>通知中心</Text>
        {unreadCount > 0 && (
          <Button type="link" size="small" onClick={handleMarkAllRead}>
            全部已读
          </Button>
        )}
      </NotificationHeader>
      
      {notifications.length > 0 ? (
        <List
          dataSource={notifications}
          renderItem={(item) => (
            <List.Item
              style={{ 
                padding: '12px 16px',
                cursor: 'pointer',
                backgroundColor: item.read ? 'transparent' : '#f6ffed',
                borderLeft: item.read ? 'none' : '3px solid #52c41a'
              }}
              onClick={() => handleNotificationClick(item.id)}
            >
              <List.Item.Meta
                avatar={getNotificationIcon(item.type)}
                title={
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <Text strong={!item.read}>{item.title}</Text>
                    <Text type="secondary" style={{ fontSize: '12px' }}>{item.time}</Text>
                  </div>
                }
                description={
                  <Text style={{ fontSize: '13px', color: '#666' }}>
                    {item.content}
                  </Text>
                }
              />
            </List.Item>
          )}
        />
      ) : (
        <div style={{ padding: '40px 16px', textAlign: 'center' }}>
          <Empty 
            image={Empty.PRESENTED_IMAGE_SIMPLE}
            description="暂无通知"
          />
        </div>
      )}
    </NotificationDropdown>
  );

  React.useEffect(() => {
    const handleFullscreenChange = () => {
      setIsFullscreen(!!document.fullscreenElement);
    };

    document.addEventListener('fullscreenchange', handleFullscreenChange);
    return () => document.removeEventListener('fullscreenchange', handleFullscreenChange);
  }, []);

  return (
    <StyledHeader>
      <LeftSection>
        <motion.div
          initial={{ opacity: 0, x: -20 }}
          animate={{ opacity: 1, x: 0 }}
          transition={{ duration: 0.3 }}
        >
          <h2 style={{ 
            margin: 0, 
            color: 'inherit',
            fontSize: '16px',
            fontWeight: 600 
          }}>
            钢材采购优化系统 V3.0
          </h2>
        </motion.div>
      </LeftSection>

      <RightSection>
        <Space size="middle">
          {/* 主题切换 */}
          <Tooltip title={isDarkMode ? '切换到浅色模式' : '切换到深色模式'}>
            <ThemeSwitch
              checked={isDarkMode}
              onChange={onThemeToggle}
              checkedChildren={<MoonOutlined />}
              unCheckedChildren={<SunOutlined />}
              size="default"
            />
          </Tooltip>

          {/* 全屏切换 */}
          <Tooltip title={isFullscreen ? '退出全屏' : '进入全屏'}>
            <ActionButton
              icon={isFullscreen ? <FullscreenExitOutlined /> : <FullscreenOutlined />}
              onClick={toggleFullscreen}
              size="large"
            />
          </Tooltip>

          {/* 通知 */}
          <Dropdown
            overlay={notificationMenu}
            trigger={['click']}
            placement="bottomRight"
            arrow
          >
            <Tooltip title="通知">
              <Badge count={unreadCount} size="small">
                <ActionButton
                  icon={<BellOutlined />}
                  size="large"
                />
              </Badge>
            </Tooltip>
          </Dropdown>
        </Space>
      </RightSection>
    </StyledHeader>
  );
};

export default Header; 