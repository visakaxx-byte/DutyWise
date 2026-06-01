import React, { useState } from 'react';
import {
  Button,
  message,
  Card,
  Tag,
  Radio,
  Collapse,
  Statistic,
  Row,
  Col,
  Divider,
  Space,
} from 'antd';
import {
  DownloadOutlined,
  CheckCircleOutlined,
  ThunderboltOutlined,
  FileExcelOutlined,
  FilePdfOutlined,
  DeleteOutlined,
} from '@ant-design/icons';
import axios from 'axios';
import { FileUpload } from '../components/FileUpload';
import { ProgressBar } from '../components/ProgressBar';

type Status = 'idle' | 'processing' | 'completed' | 'failed';

interface PipelineStats {
  total_items: number;
  optimized_items: number;
  avg_tax_reduction: number;
  max_tax_reduction: number;
  total_value_usd: number;
  output_file?: string;
}

interface PipelineResult {
  code: number;
  message: string;
  data: {
    task_id: string;
    stats: PipelineStats;
    download_url: string;
  };
}

const BPIPE_API = axios.create({
  baseURL: 'http://localhost:8001',
  timeout: 300000,
});

// 6-step pipeline progress
const BPIPE_STEPS = [
  { name: '解析文档', threshold: 16, status: 'pending' as const },
  { name: '字段映射', threshold: 33, status: 'pending' as const },
  { name: '查询税率', threshold: 50, status: 'pending' as const },
  { name: '品名归类', threshold: 66, status: 'pending' as const },
  { name: '跨章优化', threshold: 83, status: 'pending' as const },
  { name: '生成清关', threshold: 100, status: 'pending' as const },
];

// 文件类型配置
const EXCEL_ACCEPT = {
  extensions: ['.xlsx', '.xls'],
  mimeTypes: [
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-excel',
    'application/octet-stream',
  ],
  mimeLabel: 'Excel',
};

const PDF_ACCEPT = {
  extensions: ['.pdf'],
  mimeTypes: ['application/pdf', 'application/octet-stream'],
  mimeLabel: 'PDF',
};

export const BPipeline: React.FC = () => {
  const [manifestFile, setManifestFile] = useState<File | null>(null);
  const [blFile, setBlFile] = useState<File | null>(null);
  const [status, setStatus] = useState<Status>('idle');
  const [mode, setMode] = useState<'conservative' | 'aggressive'>('conservative');
  const [result, setResult] = useState<PipelineStats | null>(null);
  const [downloadUrl, setDownloadUrl] = useState<string>('');
  const [errorMessage, setErrorMessage] = useState<string>('');
  const [progress, setProgress] = useState(0);

  const handleManifestSelected = (files: File[]) => {
    setManifestFile(files[0] || null);
  };

  const handleBLSelected = (files: File[]) => {
    setBlFile(files[0] || null);
  };

  const handleRemoveManifest = () => setManifestFile(null);
  const handleRemoveBL = () => setBlFile(null);

  const canStart = manifestFile !== null && blFile !== null;

  const handleStartProcessing = async () => {
    if (!manifestFile || !blFile) {
      message.warning('请同时上传清单文件和提单文件');
      return;
    }

    try {
      setStatus('processing');
      setProgress(5);

      const formData = new FormData();
      formData.append('file', manifestFile);
      formData.append('bl_file', blFile);
      formData.append('mode', mode);
      formData.append('use_crawler', 'true');

      // Simulate progress updates
      const progressTimer = setInterval(() => {
        setProgress((prev) => {
          if (prev >= 90) return prev;
          return prev + Math.random() * 10;
        });
      }, 800);

      const response = await BPIPE_API.post<PipelineResult>('/process', formData);

      clearInterval(progressTimer);
      setProgress(100);

      if (response.data.code === 200) {
        setStatus('completed');
        setResult(response.data.data.stats);
        setDownloadUrl(response.data.data.download_url);
        message.success('流水线处理完成！');
      } else {
        throw new Error(response.data.message || '处理失败');
      }
    } catch (error: any) {
      setStatus('failed');
      const errorMsg =
        error?.response?.data?.detail ||
        error?.response?.data?.message ||
        error?.message ||
        '流水线执行失败，请检查文件格式或重试';
      setErrorMessage(errorMsg);
      message.error(errorMsg);
    }
  };

  const handleDownload = () => {
    if (downloadUrl) {
      window.open(`http://localhost:8001${downloadUrl}`, '_blank');
    }
  };

  const handleReset = () => {
    setManifestFile(null);
    setBlFile(null);
    setStatus('idle');
    setResult(null);
    setDownloadUrl('');
    setErrorMessage('');
    setProgress(0);
  };

  const formatSize = (bytes: number) => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / 1024 / 1024).toFixed(2)} MB`;
  };

  return (
    <div style={{ maxWidth: 1024, margin: '0 auto', padding: 24 }}>
      <h1 style={{ fontSize: 30, fontWeight: 'bold', marginBottom: 8 }}>
        <ThunderboltOutlined style={{ marginRight: 12, color: '#2F5496' }} />
        乙方清关流水线
      </h1>
      <p style={{ color: '#666', marginBottom: 32, fontSize: 15 }}>
        上传清单 + 提单 → 自动品名归类 → 跨章HS优化 → 低申报调整 → 生成清关文件
      </p>

      {status === 'idle' && (
        <>
          {/* Mode Selection */}
          <Card style={{ marginBottom: 16 }}>
            <h3 style={{ fontSize: 18, fontWeight: 600, marginBottom: 16 }}>
              策略模式
            </h3>
            <Radio.Group
              value={mode}
              onChange={(e) => setMode(e.target.value)}
              buttonStyle="solid"
              size="large"
            >
              <Radio.Button value="conservative">
                Conservative 保守
              </Radio.Button>
              <Radio.Button value="aggressive">
                Aggressive 激进
              </Radio.Button>
            </Radio.Group>

            <Collapse
              ghost
              style={{ marginTop: 16 }}
              items={[
                {
                  key: 'advanced',
                  label: '高级配置',
                  children: (
                    <Row gutter={16}>
                      <Col span={8}>
                        <Statistic
                          title="单价乘数"
                          value={mode === 'conservative' ? 0.15 : 0.08}
                          precision={2}
                        />
                      </Col>
                      <Col span={8}>
                        <Statistic
                          title="数量乘数"
                          value={mode === 'conservative' ? 1.0 : 0.7}
                          precision={1}
                        />
                      </Col>
                      <Col span={8}>
                        <Statistic
                          title="总价封顶"
                          value={mode === 'conservative' ? 15000 : 10000}
                          prefix="$"
                          precision={0}
                        />
                      </Col>
                    </Row>
                  ),
                },
              ]}
            />
          </Card>

          {/* File Upload Area */}
          <Row gutter={16}>
            <Col span={12}>
              <h4 style={{ fontSize: 16, fontWeight: 600, marginBottom: 8 }}>
                <FileExcelOutlined style={{ color: '#52c41a', marginRight: 8 }} />
                清单文件（必填）
              </h4>
              {!manifestFile ? (
                <FileUpload
                  onFilesSelected={handleManifestSelected}
                  accept={EXCEL_ACCEPT}
                  title="点击或拖拽上传清单"
                  hint="支持 .xlsx / .xls 格式"
                  maxSizeMB={10}
                />
              ) : (
                <Card size="small">
                  <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                    <div>
                      <FileExcelOutlined style={{ color: '#52c41a', marginRight: 8, fontSize: 20 }} />
                      <span style={{ fontSize: 14, fontWeight: 500 }}>{manifestFile.name}</span>
                      <span style={{ color: '#999', marginLeft: 12, fontSize: 13 }}>
                        {formatSize(manifestFile.size)}
                      </span>
                    </div>
                    <Space>
                      <Tag color="success">就绪</Tag>
                      <Button
                        type="text"
                        danger
                        icon={<DeleteOutlined />}
                        onClick={handleRemoveManifest}
                        size="small"
                      />
                    </Space>
                  </div>
                </Card>
              )}
            </Col>

            <Col span={12}>
              <h4 style={{ fontSize: 16, fontWeight: 600, marginBottom: 8 }}>
                <FilePdfOutlined style={{ color: '#ff4d4f', marginRight: 8 }} />
                提单文件（必填）
              </h4>
              {!blFile ? (
                <FileUpload
                  onFilesSelected={handleBLSelected}
                  accept={PDF_ACCEPT}
                  title="点击或拖拽上传提单"
                  hint="支持 .pdf 格式"
                  maxSizeMB={10}
                />
              ) : (
                <Card size="small">
                  <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                    <div>
                      <FilePdfOutlined style={{ color: '#ff4d4f', marginRight: 8, fontSize: 20 }} />
                      <span style={{ fontSize: 14, fontWeight: 500 }}>{blFile.name}</span>
                      <span style={{ color: '#999', marginLeft: 12, fontSize: 13 }}>
                        {formatSize(blFile.size)}
                      </span>
                    </div>
                    <Space>
                      <Tag color="success">就绪</Tag>
                      <Button
                        type="text"
                        danger
                        icon={<DeleteOutlined />}
                        onClick={handleRemoveBL}
                        size="small"
                      />
                    </Space>
                  </div>
                </Card>
              )}
            </Col>
          </Row>

          {/* Start Button */}
          <Card style={{ marginTop: 24 }}>
            {!canStart && (
              <p style={{ color: '#faad14', textAlign: 'center', marginBottom: 12 }}>
                {!manifestFile && !blFile
                  ? '请上传清单文件和提单文件'
                  : !manifestFile
                  ? '请上传清单文件'
                  : '请上传提单文件'}
              </p>
            )}
            <Button
              type="primary"
              size="large"
              style={{ width: '100%' }}
              disabled={!canStart}
              onClick={handleStartProcessing}
            >
              开始流水线
            </Button>
          </Card>
        </>
      )}

      {status === 'processing' && (
        <ProgressBar
          progress={Math.min(Math.round(progress), 99)}
          currentStep="流水线执行中"
          steps={BPIPE_STEPS}
        />
      )}

      {status === 'completed' && result && (
        <Card>
          <div style={{ textAlign: 'center', marginBottom: 24 }}>
            <CheckCircleOutlined
              style={{ fontSize: 64, color: '#52c41a' }}
            />
            <h2 style={{ fontSize: 24, fontWeight: 'bold', marginTop: 16 }}>
              流水线完成！
            </h2>
          </div>

          <Divider />

          <Row gutter={16} style={{ marginBottom: 24 }}>
            <Col span={8}>
              <Statistic
                title="总商品数"
                value={result.total_items}
                suffix="个"
              />
            </Col>
            <Col span={8}>
              <Statistic
                title="优化商品数"
                value={result.optimized_items}
                suffix="个"
              />
            </Col>
            <Col span={8}>
              <Statistic
                title="优化率"
                value={
                  result.total_items > 0
                    ? (
                        (result.optimized_items / result.total_items) *
                        100
                      ).toFixed(1)
                    : '0'
                }
                suffix="%"
              />
            </Col>
          </Row>

          <Row gutter={16} style={{ marginBottom: 24 }}>
            <Col span={8}>
              <Statistic
                title="平均降税"
                value={result.avg_tax_reduction.toFixed(1)}
                suffix="%"
              />
            </Col>
            <Col span={8}>
              <Statistic
                title="最大降税"
                value={result.max_tax_reduction.toFixed(1)}
                suffix="%"
              />
            </Col>
            <Col span={8}>
              <Statistic
                title="申报总价"
                value={result.total_value_usd.toFixed(2)}
                prefix="$"
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
              下载清关文件
            </Button>
            <Button size="large" onClick={handleReset}>
              处理新文件
            </Button>
          </div>
        </Card>
      )}

      {status === 'failed' && (
        <Card>
          <div
            style={{
              textAlign: 'center',
              paddingTop: 48,
              paddingBottom: 48,
            }}
          >
            <p
              style={{
                fontSize: 20,
                color: '#ff4d4f',
                marginBottom: 16,
              }}
            >
              流水线执行失败
            </p>
            <p
              style={{
                color: '#666',
                marginBottom: 24,
                whiteSpace: 'pre-wrap',
                lineHeight: 1.6,
              }}
            >
              {errorMessage || '请检查文件格式是否正确，或稍后重试'}
            </p>
            <Space>
              <Button type="primary" onClick={handleReset}>
                重新开始
              </Button>
              <Button onClick={() => setStatus('idle')}>
                返回修改
              </Button>
            </Space>
          </div>
        </Card>
      )}
    </div>
  );
};
