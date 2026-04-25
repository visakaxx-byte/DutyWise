import React, { useCallback } from 'react';
import { Upload, Card, message } from 'antd';
import { InboxOutlined } from '@ant-design/icons';
import type { UploadProps } from 'antd';

const { Dragger } = Upload;

interface FileUploadProps {
  onFilesSelected: (files: File[]) => void;
}

export const FileUpload: React.FC<FileUploadProps> = ({ onFilesSelected }) => {
  const handleChange: UploadProps['onChange'] = useCallback((info: any) => {
    const { fileList } = info;

    // 过滤出有效的文件
    const validFiles = fileList
      .filter((file: any) => file.originFileObj)
      .map((file: any) => file.originFileObj as File);

    onFilesSelected(validFiles);
  }, [onFilesSelected]);

  const beforeUpload = (file: File) => {
    const isExcel = file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
                    file.type === 'application/vnd.ms-excel' ||
                    file.name.endsWith('.xlsx') ||
                    file.name.endsWith('.xls');

    if (!isExcel) {
      message.error('只能上传 Excel 文件（.xlsx 或 .xls）');
      return Upload.LIST_IGNORE;
    }

    const isLt10M = file.size / 1024 / 1024 < 10;
    if (!isLt10M) {
      message.error('文件大小不能超过 10MB');
      return Upload.LIST_IGNORE;
    }

    return false; // 阻止自动上传
  };

  return (
    <Card>
      <Dragger
        name="files"
        multiple
        beforeUpload={beforeUpload}
        onChange={handleChange}
        accept=".xlsx,.xls"
        showUploadList={false}
      >
        <p className="ant-upload-drag-icon">
          <InboxOutlined />
        </p>
        <p className="ant-upload-text">点击或拖拽文件到此区域上传</p>
        <p className="ant-upload-hint">
          支持单个或批量上传 Excel 文件（.xlsx, .xls）
        </p>
      </Dragger>
    </Card>
  );
};
