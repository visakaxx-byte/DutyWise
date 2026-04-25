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
  status: 'pending' | 'processing' | 'completed' | 'failed';
  progress: number;
  currentStep: string;
  errorMessage?: string;
  steps?: Array<{
    name: string;
    status: 'completed' | 'processing' | 'pending';
  }>;
}

export interface ProcessResult {
  shipmentId: number;
  shipmentNo: string;
  status: string;
  statistics: {
    total_items: number;
    optimized_items: number;
    optimization_rate: number;
  };
  files: {
    output_file: string | null;
    log_file: string | null;
  };
  totalItems?: number;
  optimizedItems?: number;
  savedAmount?: number;
  outputFileId?: string;
  outputFileName?: string;
}
