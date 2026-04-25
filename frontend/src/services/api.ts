import axios from 'axios';
import type { Settings, TaskStatus, ProcessResult } from '../types';

// 创建 axios 实例
const api = axios.create({
  baseURL: import.meta.env.VITE_API_BASE_URL || 'http://localhost:8000/api/v1',
  timeout: 30000,
});

// 请求拦截器
api.interceptors.request.use(
  (config) => {
    return config;
  },
  (error) => {
    return Promise.reject(error);
  }
);

// 响应拦截器
api.interceptors.response.use(
  (response) => {
    return response;
  },
  (error) => {
    console.error('API Error:', error);
    return Promise.reject(error);
  }
);

// 上传文件
export const uploadFiles = async (files: File[]) => {
  const formData = new FormData();
  files.forEach(file => formData.append('files', file));

  const { data } = await api.post('/shipments/upload', formData, {
    headers: { 'Content-Type': 'multipart/form-data' }
  });
  return data.data;
};

// 开始处理
export const startProcessing = async (shipmentId: number, options?: object) => {
  const { data } = await api.post(`/shipments/${shipmentId}/process`, options || {});
  return data.data;
};

// 查询任务状态
export const getTaskStatus = async (taskId: string): Promise<TaskStatus> => {
  const { data } = await api.get(`/tasks/${taskId}/status`);
  return data.data;
};

// 获取处理结果
export const getShipmentResult = async (shipmentId: number): Promise<ProcessResult> => {
  const { data } = await api.get(`/shipments/${shipmentId}/result`);
  return data.data;
};

// 下载文件
export const downloadFile = (fileId: string) => {
  const url = `${api.defaults.baseURL}/files/download/${fileId}`;
  window.open(url, '_blank');
};

// 获取历史记录
export const getShipments = async (page = 1, pageSize = 20) => {
  const { data } = await api.get('/shipments', {
    params: { page, page_size: pageSize }
  });
  return data.data;
};

// 测试 LLM 连接
export const testLLMConnection = async (baseUrl: string, apiKey: string, modelId: string) => {
  const { data } = await api.post('/settings/test-llm', {
    baseUrl,
    apiKey,
    modelId
  });
  return data.data;
};

// 保存设置
export const saveSettings = async (settings: Settings) => {
  const { data } = await api.post('/settings', settings);
  return data.data;
};

// 获取设置
export const getSettings = async (): Promise<Settings> => {
  const { data } = await api.get('/settings');
  return data.data;
};

export default api;
