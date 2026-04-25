import React from 'react';
import { Card, Button, Statistic, Row, Col, Divider } from 'antd';
import { DownloadOutlined, CheckCircleOutlined } from '@ant-design/icons';
import type { ProcessResult } from '../types';
interface ResultCardProps {
  result: ProcessResult;
  onReset: () => void;
}

export const ResultCard: React.FC<ResultCardProps> = ({ result, onReset }) => {
  const handleDownload = () => {
    const outputFile = result.files?.output_file;
    const logFile = result.files?.log_file;
    if (outputFile) {
      // 后端提供文件下载
      const filename = outputFile.split('/').pop() || 'output.xlsx';
      window.open(`http://localhost:8000/api/v1/files/download?path=${encodeURIComponent(outputFile)}&name=${encodeURIComponent(filename)}`, '_blank');
    } else if (logFile) {
      const filename = logFile.split('/').pop() || 'log.xlsx';
      window.open(`http://localhost:8000/api/v1/files/download?path=${encodeURIComponent(logFile)}&name=${encodeURIComponent(filename)}`, '_blank');
    }
  };

  const totalItems = result.statistics?.total_items ?? result.totalItems ?? 0;
  const optimizedItems = result.statistics?.optimized_items ?? result.optimizedItems ?? 0;
  const optimizationRate = result.statistics?.optimization_rate ?? 0;
  const hasFiles = !!(result.files?.output_file || result.files?.log_file);

  return (
    <Card>
      <div style={{ textAlign: 'center', marginBottom: 24 }}>
        <CheckCircleOutlined style={{ fontSize: 64, color: '#52c41a' }} />
        <h2 style={{ fontSize: 24, fontWeight: 'bold', marginTop: 16 }}>处理完成！</h2>
        {!hasFiles && (
          <p style={{ color: '#faad14', marginTop: 8 }}>注意：未生成输出文件（可能需要优化匹配项或模板文件）</p>
        )}
      </div>

      <Divider />

      <Row gutter={16} style={{ marginBottom: 24 }}>
        <Col span={8}>
          <Statistic
            title="总商品数"
            value={totalItems}
            suffix="个"
          />
        </Col>
        <Col span={8}>
          <Statistic
            title="优化商品数"
            value={optimizedItems}
            suffix="个"
          />
        </Col>
        <Col span={8}>
          <Statistic
            title="优化率"
            value={Number(optimizationRate).toFixed(1)}
            suffix="%"
          />
        </Col>
      </Row>

      <Divider />

      <div style={{ display: 'flex', gap: 16, justifyContent: 'center' }}>
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
