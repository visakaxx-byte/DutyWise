import React from 'react';
import { Progress, Card, Steps } from 'antd';
import type { TaskStatus } from '../types';

interface ProgressBarProps {
  progress: number;
  currentStep: string;
  steps: TaskStatus['steps'];
}

export const ProgressBar: React.FC<ProgressBarProps> = ({
  progress,
  currentStep,
  steps
}) => {
  // 将步骤状态转换为 Steps 组件的状态
  const getStepStatus = (status: string) => {
    switch (status) {
      case 'completed':
        return 'finish';
      case 'processing':
        return 'process';
      case 'pending':
        return 'wait';
      default:
        return 'wait';
    }
  };

  const currentStepIndex = steps.findIndex(step => step.status === 'processing');

  return (
    <Card>
      <div className="space-y-6">
        <div>
          <div className="flex justify-between mb-2">
            <span className="text-lg font-semibold">处理进度</span>
            <span className="text-lg font-semibold">{progress}%</span>
          </div>
          <Progress
            percent={progress}
            status={progress === 100 ? 'success' : 'active'}
            strokeColor={{
              '0%': '#108ee9',
              '100%': '#87d068',
            }}
          />
        </div>

        <div>
          <p className="text-gray-600 mb-4">当前步骤: {currentStep}</p>
          <Steps
            current={currentStepIndex}
            direction="vertical"
            items={steps.map((step) => ({
              title: step.name,
              status: getStepStatus(step.status),
            }))}
          />
        </div>
      </div>
    </Card>
  );
};
