import React from 'react';
import { Card, Button, Statistic, Row, Col, Divider } from 'antd';
import { DownloadOutlined, CheckCircleOutlined } from '@ant-design/icons';
import type { ProcessResult } from '../types';
import { downloadFile } from '../services/api';

interface ResultCardProps {
  result: ProcessResult;
  onReset: () => void;
}

export const ResultCard: React.FC<ResultCardProps> = ({ result, onReset }) => {
  const handleDownload = () => {
    downloadFile(result.outputFileId);
  };

  return (
    <Card>
      <div className="text-center mb-6">
        <CheckCircleOutlined style={{ fontSize: 64, color: '#52c41a' }} />
        <h2 className="text-2xl font-bold mt-4">处理完成！</h2>
      </div>

      <Divider />

      <Row gutter={16} className="mb-6">
        <Col span={8}>
          <Statistic
            title="总商品数"
            value={result.totalItems}
            suffix="个"
          />
        </Col>
        <Col span={8}>
          <Statistic
            title="优化商品数"
            value={result.optimizedItems}
            suffix="个"
          />
        </Col>
        <Col span={8}>
          <Statistic
            title="优化率"
            value={((result.optimizedItems / result.totalItems) * 100).toFixed(1)}
            suffix="%"
          />
        </Col>
      </Row>

      <Divider />

      <Row gutter={16} className="mb-6">
        <Col span={8}>
          <Statistic
            title="原始税费"
            value={result.statistics.totalTax}
            precision={2}
            prefix="¥"
          />
        </Col>
        <Col span={8}>
          <Statistic
            title="优化后税费"
            value={result.statistics.optimizedTax}
            precision={2}
            prefix="¥"
          />
        </Col>
        <Col span={8}>
          <Statistic
            title="节省金额"
            value={result.statistics.savedTax}
            precision={2}
            prefix="¥"
            valueStyle={{ color: '#cf1322' }}
          />
        </Col>
      </Row>

      <Divider />

      <div className="flex gap-4 justify-center">
        <Button
          type="primary"
          size="large"
          icon={<DownloadOutlined />}
          onClick={handleDownload}
        >
          下载结果文件
        </Button>
        <Button size="large" onClick={onReset}>
          处理新文件
        </Button>
      </div>
    </Card>
  );
};
