import React from 'react';
import { BrowserRouter, Routes, Route, Link } from 'react-router-dom';
import { Layout, Menu } from 'antd';
import { HomeOutlined, HistoryOutlined, SettingOutlined, ThunderboltOutlined } from '@ant-design/icons';
import { Home } from './pages/Home';
import { History } from './pages/History';
import { Settings } from './pages/Settings';
import { BPipeline } from './pages/BPipeline';
import './App.css';

const { Header, Sider, Content } = Layout;

const App: React.FC = () => {
  const [collapsed, setCollapsed] = React.useState(false);

  const menuItems = [
    {
      key: '/',
      icon: <HomeOutlined />,
      label: <Link to="/">首页</Link>,
    },
    {
      key: '/bpipeline',
      icon: <ThunderboltOutlined />,
      label: <Link to="/bpipeline">乙方流水线</Link>,
    },
    {
      key: '/history',
      icon: <HistoryOutlined />,
      label: <Link to="/history">历史记录</Link>,
    },
    {
      key: '/settings',
      icon: <SettingOutlined />,
      label: <Link to="/settings">设置</Link>,
    },
  ];

  return (
    <BrowserRouter>
      <Layout style={{ minHeight: '100vh' }}>
        <Sider
          collapsible
          collapsed={collapsed}
          onCollapse={setCollapsed}
          theme="light"
        >
          <div className="logo" style={{
            height: 64,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            fontSize: 18,
            fontWeight: 'bold',
            borderBottom: '1px solid #f0f0f0'
          }}>
            {collapsed ? '清关' : '清关优化系统'}
          </div>
          <Menu
            mode="inline"
            defaultSelectedKeys={['/']}
            items={menuItems}
            style={{ borderRight: 0 }}
          />
        </Sider>
        <Layout>
          <Header style={{
            padding: '0 24px',
            background: '#fff',
            borderBottom: '1px solid #f0f0f0',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between'
          }}>
            <h2 style={{ margin: 0 }}>清关优化系统</h2>
          </Header>
          <Content style={{
            margin: 0,
            background: '#f0f2f5',
            overflow: 'auto'
          }}>
            <Routes>
              <Route path="/" element={<Home />} />
              <Route path="/bpipeline" element={<BPipeline />} />
              <Route path="/history" element={<History />} />
              <Route path="/settings" element={<Settings />} />
            </Routes>
          </Content>
        </Layout>
      </Layout>
    </BrowserRouter>
  );
};

export default App;
