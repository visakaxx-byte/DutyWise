import React from 'react';
import { Progress, Card, Steps } from 'antd';
import type { TaskStatus } from '../types';

const DEFAULT_STEPS = [
  { name: '解析文档', threshold: 10, status: 'wait' as const },
  { name: '字段映射', threshold: 30, status: 'wait' as const },
  { name: '查询税率', threshold: 50, status: 'wait' as const },
  { name: '优化HS编码', threshold: 70, status: 'wait' as const },
  { name: '生成文件', threshold: 90, status: 'wait' as const },
];

interface ProgressBarProps {
  progress: number;
  currentStep: string;
  steps?: TaskStatus['steps'];
}

function computeSteps(progress: number, _currentStep: string): Array<{ name: string; status: string }> {
  return DEFAULT_STEPS.map((step, idx) => {
    if (progress >= DEFAULT_STEPS[idx].threshold) {
      return { name: step.name, status: 'completed' };
    }
    if (idx === 0 || (idx > 0 && progress >= DEFAULT_STEPS[idx - 1].threshold)) {
      return { name: step.name, status: 'processing' };
    }
    return { name: step.name, status: 'pending' };
  });
}

export const ProgressBar: React.FC<ProgressBarProps> = ({
  progress,
  currentStep,
  steps
}) => {
  const stepItems = steps && steps.length > 0
    ? steps.map((step) => ({
      title: step.name,
      status: step.status === 'completed' ? 'finish' as const
        : step.status === 'processing' ? 'process' as const
        : 'wait' as const,
    }))
    : computeSteps(progress, currentStep).map((step) => ({
      title: step.name,
      status: step.status === 'completed' ? 'finish' as const
        : step.status === 'processing' ? 'process' as const
        : 'wait' as const,
    }));

  const currentStepIndex = computeSteps(progress, currentStep).findIndex(
    s => s.status === 'processing'
  );

  return (
    <Card>
      <div style={{ marginBottom: 24 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 8 }}>
          <span style={{ fontSize: 18, fontWeight: 600 }}>处理进度</span>
          <span style={{ fontSize: 18, fontWeight: 600 }}>{progress}%</span>
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
        <p style={{ color: '#666', marginBottom: 16 }}>当前步骤: {currentStep}</p>
        <Steps
          current={currentStepIndex >= 0 ? currentStepIndex : stepItems.filter(s => s.status === 'finish').length}
          direction="vertical"
          items={stepItems}
        />
      </div>
    </Card>
  );
};
