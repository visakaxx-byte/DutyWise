import React, { useEffect, useState } from 'react';
import { Table, Card, Tag, Button, Space, message } from 'antd';
import { ReloadOutlined } from '@ant-design/icons';
import type { ColumnsType } from 'antd/es/table';
import { getShipments, getShipmentResult } from '../services/api';
import type { Shipment } from '../types';
import dayjs from 'dayjs';

export const History: React.FC = () => {
  const [loading, setLoading] = useState(false);
  const [shipments, setShipments] = useState<Shipment[]>([]);
  const [pagination, setPagination] = useState({
    current: 1,
    pageSize: 20,
    total: 0,
  });

  const fetchShipments = async (page = 1, pageSize = 20) => {
    try {
      setLoading(true);
      const result = await getShipments(page, pageSize);
      setShipments(result.items || []);
      setPagination({
        current: page,
        pageSize,
        total: result.total || 0,
      });
    } catch (error: any) {
      const errorMsg = error?.response?.data?.message || '获取历史记录失败';
      message.error(errorMsg);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchShipments();
  }, []);

  const handleTableChange = (pagination: any) => {
    fetchShipments(pagination.current, pagination.pageSize);
  };

  const handleRefresh = () => {
    fetchShipments(pagination.current, pagination.pageSize);
  };

  const columns: ColumnsType<Shipment> = [
    {
      title: '批次号',
      dataIndex: 'shipmentNo',
      key: 'shipmentNo',
      width: 200,
    },
    {
      title: '状态',
      dataIndex: 'status',
      key: 'status',
      width: 120,
      render: (status: string) => {
        const statusMap = {
          processing: { color: 'processing', text: '处理中' },
          completed: { color: 'success', text: '已完成' },
          failed: { color: 'error', text: '失败' },
        };
        const config = statusMap[status as keyof typeof statusMap] || { color: 'default', text: status };
        return <Tag color={config.color}>{config.text}</Tag>;
      },
    },
    {
      title: '总商品数',
      dataIndex: 'totalItems',
      key: 'totalItems',
      width: 120,
      align: 'right',
    },
    {
      title: '优化商品数',
      dataIndex: 'optimizedItems',
      key: 'optimizedItems',
      width: 120,
      align: 'right',
    },
    {
      title: '优化率',
      key: 'optimizationRate',
      width: 120,
      align: 'right',
      render: (_, record) => {
        if (record.totalItems === 0) return '-';
        const rate = ((record.optimizedItems / record.totalItems) * 100).toFixed(1);
        return `${rate}%`;
      },
    },
    {
      title: '创建时间',
      dataIndex: 'createdAt',
      key: 'createdAt',
      width: 180,
      render: (date: string) => dayjs(date).format('YYYY-MM-DD HH:mm:ss'),
    },
    {
      title: '完成时间',
      dataIndex: 'completedAt',
      key: 'completedAt',
      width: 180,
      render: (date: string) => date ? dayjs(date).format('YYYY-MM-DD HH:mm:ss') : '-',
    },
    {
      title: '操作',
      key: 'action',
      width: 150,
      fixed: 'right',
      render: (_, record) => (
        <Space size="small">
          {record.status === 'completed' && (
            <Button
              type="link"
              size="small"
              icon={<span>⬇</span>}
              onClick={async () => {
                try {
                  const result = await getShipmentResult(record.id);
                  const outputFile = result.files?.output_file;
                  const logFile = result.files?.log_file;
                  if (outputFile) {
                    const filename = outputFile.split('/').pop() || 'output.xlsx';
                    window.open(
                      `http://localhost:8000/api/v1/files/download?path=${encodeURIComponent(outputFile)}&name=${encodeURIComponent(filename)}`,
                      '_blank'
                    );
                  } else if (logFile) {
                    const filename = logFile.split('/').pop() || 'log.xlsx';
                    window.open(
                      `http://localhost:8000/api/v1/files/download?path=${encodeURIComponent(logFile)}&name=${encodeURIComponent(filename)}`,
                      '_blank'
                    );
                  } else {
                    message.warning('没有可下载的文件');
                  }
                } catch {
                  message.error('获取文件信息失败');
                }
              }}
            >
              下载
            </Button>
          )}
        </Space>
      ),
    },
  ];

  return (
    <div style={{ maxWidth: 1280, margin: '0 auto', padding: 24 }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 24 }}>
        <h1 style={{ fontSize: 24, fontWeight: 'bold' }}>历史记录</h1>
        <Button
          icon={<ReloadOutlined />}
          onClick={handleRefresh}
          loading={loading}
        >
          刷新
        </Button>
      </div>

      <Card>
        <Table
          columns={columns}
          dataSource={shipments}
          rowKey="id"
          loading={loading}
          pagination={pagination}
          onChange={handleTableChange}
          scroll={{ x: 1200 }}
        />
      </Card>
    </div>
  );
};
