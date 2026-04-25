// 类型定义

export interface LLMConfig {
  baseUrl: string;
  apiKey: string;
  modelId: string;
}

export interface Settings {
  llm: LLMConfig;
  excludeAntiDumping: boolean;
  minSimilarity: number;
  crawlerUsername: string;
  crawlerPassword: string;
}

export interface Shipment {
  id: number;
  shipmentNo: string;
  status: 'processing' | 'completed' | 'failed';
  totalItems: number;
  optimizedItems: number;
  createdAt: string;
  completedAt?: string;
}

export interface TaskStatus {
  taskId: string;
  status: 'processing' | 'completed' | 'failed';
  progress: number;
  currentStep: string;
  steps: Array<{
    name: string;
    status: 'completed' | 'processing' | 'pending';
  }>;
}

export interface ProcessResult {
  shipmentId: number;
  totalItems: number;
  optimizedItems: number;
  savedAmount: number;
  outputFileId: string;
  outputFileName: string;
  statistics: {
    totalTax: number;
    optimizedTax: number;
    savedTax: number;
  };
}
