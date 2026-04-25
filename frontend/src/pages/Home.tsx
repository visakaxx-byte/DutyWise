import React, { useState } from 'react';
import { Button, message, Card, List, Tag } from 'antd';
import { FileUpload } from '../components/FileUpload';
import { ProgressBar } from '../components/ProgressBar';
import { ResultCard } from '../components/ResultCard';
import { uploadFiles, startProcessing, getTaskStatus, getShipmentResult } from '../services/api';
import type { TaskStatus, ProcessResult } from '../types';

type Status = 'idle' | 'processing' | 'completed' | 'failed';

export const Home: React.FC = () => {
  const [files, setFiles] = useState<File[]>([]);
  const [status, setStatus] = useState<Status>('idle');
  const [taskStatus, setTaskStatus] = useState<TaskStatus | null>(null);
  const [result, setResult] = useState<ProcessResult | null>(null);

  const handleFilesSelected = (selectedFiles: File[]) => {
    setFiles(selectedFiles);
  };

  const handleRemoveFile = (index: number) => {
    setFiles(files.filter((_, i) => i !== index));
  };

  const handleStartProcessing = async () => {
    if (files.length === 0) {
      message.warning('请先上传文件');
      return;
    }

    try {
      setStatus('processing');
      message.loading('正在上传文件...', 0);

      // 1. 上传文件
      const uploadResult = await uploadFiles(files);
      const { shipment_id } = uploadResult;

      message.destroy();
      message.success('文件上传成功，开始处理...');

      // 2. 开始处理
      const processResult = await startProcessing(shipment_id);
      const task_id = processResult.task_id;

      // 3. 轮询任务状态
      const pollInterval = setInterval(async () => {
        try {
          const status = await getTaskStatus(task_id);
          setTaskStatus(status);

          if (status.status === 'completed') {
            clearInterval(pollInterval);
            setStatus('completed');

            // 获取结果
            const result = await getShipmentResult(shipment_id);
            setResult(result);
            message.success('处理完成！');
          } else if (status.status === 'failed') {
            clearInterval(pollInterval);
            setStatus('failed');
            message.error('处理失败，请检查文件格式或重试');
          }
        } catch (error) {
          clearInterval(pollInterval);
          setStatus('failed');
          message.error('获取任务状态失败');
        }
      }, 2000);

    } catch (error: any) {
      setStatus('failed');
      const errorMsg = error?.response?.data?.message || '处理失败，请重试';
      message.error(errorMsg);
    }
  };

  const handleReset = () => {
    setFiles([]);
    setStatus('idle');
    setTaskStatus(null);
    setResult(null);
  };

  return (
    <div style={{ maxWidth: 1024, margin: '0 auto', padding: 24 }}>
      <h1 style={{ fontSize: 30, fontWeight: 'bold', marginBottom: 32 }}>清关优化系统</h1>

      {status === 'idle' && (
        <>
          <FileUpload onFilesSelected={handleFilesSelected} />

          {files.length > 0 && (
            <Card style={{ marginTop: 24 }}>
              <h3 style={{ fontSize: 18, fontWeight: 600, marginBottom: 16 }}>已上传文件</h3>
              <List
                dataSource={files}
                renderItem={(file, index) => (
                  <List.Item
                    actions={[
                      <Button
                        key="delete"
                        danger
                        size="small"
                        onClick={() => handleRemoveFile(index)}
                      >
                        删除
                      </Button>
                    ]}
                  >
                    <List.Item.Meta
                      title={file.name}
                      description={`大小: ${(file.size / 1024).toFixed(2)} KB`}
                    />
                    <Tag color="success">就绪</Tag>
                  </List.Item>
                )}
              />
              <Button
                type="primary"
                size="large"
                style={{ marginTop: 24, width: '100%' }}
                onClick={handleStartProcessing}
              >
                开始处理
              </Button>
            </Card>
          )}
        </>
      )}

      {status === 'processing' && taskStatus && (
        <ProgressBar
          progress={taskStatus.progress}
          currentStep={taskStatus.currentStep}
          steps={taskStatus.steps}
        />
      )}

      {status === 'completed' && result && (
        <ResultCard result={result} onReset={handleReset} />
      )}

      {status === 'failed' && (
        <Card>
          <div style={{ textAlign: 'center', paddingTop: 48, paddingBottom: 48 }}>
            <p style={{ fontSize: 20, color: '#ff4d4f', marginBottom: 16 }}>处理失败</p>
            <p style={{ color: '#666', marginBottom: 24 }}>请检查文件格式是否正确，或稍后重试</p>
            <Button type="primary" onClick={handleReset}>
              重新开始
            </Button>
          </div>
        </Card>
      )}
    </div>
  );
};
