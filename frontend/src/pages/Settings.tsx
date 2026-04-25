import React, { useState } from 'react';
import { Form, Input, Button, Radio, Switch, message, Card, Space } from 'antd';
import { useSettingsStore } from '../store/useSettingsStore';
import { testLLMConnection, saveSettings } from '../services/api';

export const Settings: React.FC = () => {
  const { settings, updateSettings, resetSettings } = useSettingsStore();
  const [form] = Form.useForm();
  const [testing, setTesting] = useState(false);
  const [saving, setSaving] = useState(false);

  const presetModels = [
    {
      label: '豆包 Doubao-pro-32k (推荐)',
      value: 'doubao-pro',
      config: {
        baseUrl: 'https://ark.cn-beijing.volces.com/api/v3',
        modelId: 'ep-20240611125520-lmk7v'
      }
    },
    {
      label: '豆包 Doubao-lite-32k',
      value: 'doubao-lite',
      config: {
        baseUrl: 'https://ark.cn-beijing.volces.com/api/v3',
        modelId: 'ep-20240611125520-xxxxx'
      }
    },
    {
      label: '自定义',
      value: 'custom',
      config: { baseUrl: '', modelId: '' }
    }
  ];

  const handleSave = async (values: any) => {
    try {
      setSaving(true);
      await saveSettings(values);
      updateSettings(values);
      message.success('设置已保存');
    } catch (error: any) {
      const errorMsg = error?.response?.data?.message || '保存失败';
      message.error(errorMsg);
    } finally {
      setSaving(false);
    }
  };

  const handleTestConnection = async () => {
    try {
      setTesting(true);
      const values = form.getFieldsValue();
      await testLLMConnection(
        values.llm.baseUrl,
        values.llm.apiKey,
        values.llm.modelId
      );
      message.success('连接成功！');
    } catch (error: any) {
      const errorMsg = error?.response?.data?.message || '连接失败，请检查配置';
      message.error(errorMsg);
    } finally {
      setTesting(false);
    }
  };

  const handlePresetChange = (value: string) => {
    const preset = presetModels.find(p => p.value === value);
    if (preset && preset.value !== 'custom') {
      form.setFieldsValue({
        llm: {
          ...form.getFieldValue('llm'),
          ...preset.config
        }
      });
    }
  };

  const handleReset = () => {
    resetSettings();
    form.resetFields();
    message.info('已恢复默认设置');
  };

  return (
    <div className="max-w-3xl mx-auto p-6">
      <h1 className="text-2xl font-bold mb-6">设置</h1>

      <Form
        form={form}
        layout="vertical"
        initialValues={settings}
        onFinish={handleSave}
      >
        {/* LLM 配置 */}
        <Card title="LLM 配置" className="mb-6">
          <Form.Item label="预设模型" name="preset">
            <Radio.Group onChange={(e) => handlePresetChange(e.target.value)}>
              <Space direction="vertical">
                {presetModels.map(model => (
                  <Radio key={model.value} value={model.value}>
                    {model.label}
                  </Radio>
                ))}
              </Space>
            </Radio.Group>
          </Form.Item>

          <Form.Item
            label="Base URL"
            name={['llm', 'baseUrl']}
            rules={[{ required: true, message: '请输入 Base URL' }]}
          >
            <Input placeholder="https://ark.cn-beijing.volces.com/api/v3" />
          </Form.Item>

          <Form.Item
            label="API Key"
            name={['llm', 'apiKey']}
            rules={[{ required: true, message: '请输入 API Key' }]}
          >
            <Input.Password placeholder="输入你的 API Key" />
          </Form.Item>

          <Form.Item
            label="Model ID"
            name={['llm', 'modelId']}
            rules={[{ required: true, message: '请输入 Model ID' }]}
          >
            <Input placeholder="ep-20240611125520-lmk7v" />
          </Form.Item>

          <Button loading={testing} onClick={handleTestConnection}>
            测试连接
          </Button>
        </Card>

        {/* 优化设置 */}
        <Card title="优化设置" className="mb-6">
          <Form.Item
            label="排除有反倾销标记的HS编码"
            name="excludeAntiDumping"
            valuePropName="checked"
          >
            <Switch />
          </Form.Item>

          <Form.Item
            label="最低相似度阈值"
            name="minSimilarity"
            extra="0-1之间，越高越严格"
            rules={[
              { required: true, message: '请输入相似度阈值' },
              { type: 'number', min: 0, max: 1, message: '请输入0-1之间的数值' }
            ]}
          >
            <Input type="number" min={0} max={1} step={0.1} />
          </Form.Item>
        </Card>

        {/* 爬虫配置 */}
        <Card title="爬虫配置" className="mb-6">
          <Form.Item
            label="网站账号"
            name="crawlerUsername"
            rules={[{ required: true, message: '请输入网站账号' }]}
          >
            <Input placeholder="请输入账号" />
          </Form.Item>

          <Form.Item
            label="网站密码"
            name="crawlerPassword"
            rules={[{ required: true, message: '请输入网站密码' }]}
          >
            <Input.Password placeholder="请输入密码" />
          </Form.Item>
        </Card>

        <div className="flex gap-4">
          <Button
            type="primary"
            htmlType="submit"
            size="large"
            loading={saving}
          >
            保存设置
          </Button>
          <Button size="large" onClick={handleReset}>
            恢复默认
          </Button>
        </div>
      </Form>
    </div>
  );
};
