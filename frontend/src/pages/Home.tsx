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
      const { shipment_id, task_id } = uploadResult;

      message.destroy();
      message.success('文件上传成功，开始处理...');

      // 2. 开始处理
      await startProcessing(shipment_id);

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
    <div className="max-w-4xl mx-auto p-6">
      <h1 className="text-3xl font-bold mb-8">清关优化系统</h1>

      {status === 'idle' && (
        <>
          <FileUpload onFilesSelected={handleFilesSelected} />

          {files.length > 0 && (
            <Card className="mt-6">
              <h3 className="text-lg font-semibold mb-4">已上传文件</h3>
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
                className="mt-6 w-full"
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
          <div className="text-center py-12">
            <p className="text-xl text-red-500 mb-4">处理失败</p>
            <p className="text-gray-600 mb-6">请检查文件格式是否正确，或稍后重试</p>
            <Button type="primary" onClick={handleReset}>
              重新开始
            </Button>
          </div>
        </Card>
      )}
    </div>
  );
};
