import React, { useCallback } from 'react';
import { Upload, Card, message } from 'antd';
import { InboxOutlined } from '@ant-design/icons';
import type { UploadProps } from 'antd';

const { Dragger } = Upload;

interface FileTypeConfig {
  extensions: string[];
  mimeTypes: string[];
  mimeLabel: string;   // 用于错误提示，如 "Excel" / "PDF"
}

interface FileUploadProps {
  onFilesSelected: (files: File[]) => void;
  accept: FileTypeConfig;
  title?: string;
  hint?: string;
  maxSizeMB?: number;
  multiple?: boolean;
}

export const FileUpload: React.FC<FileUploadProps> = ({
  onFilesSelected,
  accept,
  title = '点击或拖拽文件到此区域上传',
  hint,
  maxSizeMB = 10,
  multiple = false,
}) => {
  const handleChange: UploadProps['onChange'] = useCallback((info: any) => {
    const { fileList } = info;

    const validFiles = fileList
      .filter((file: any) => file.originFileObj)
      .map((file: any) => file.originFileObj as File);

    onFilesSelected(validFiles);
  }, [onFilesSelected]);

  const beforeUpload = (file: File) => {
    const fileName = file.name.toLowerCase();

    const isValidType =
      accept.mimeTypes.includes(file.type) ||
      accept.extensions.some((ext) => fileName.endsWith(ext));

    if (!isValidType) {
      const extList = accept.extensions.join(', ');
      message.error(`只能上传 ${accept.mimeLabel} 文件（${extList}）`);
      return Upload.LIST_IGNORE;
    }

    const isLtLimit = file.size / 1024 / 1024 < maxSizeMB;
    if (!isLtLimit) {
      message.error(`文件大小不能超过 ${maxSizeMB}MB`);
      return Upload.LIST_IGNORE;
    }

    return false; // 阻止自动上传
  };

  return (
    <Card>
      <Dragger
        name="files"
        multiple={multiple}
        beforeUpload={beforeUpload}
        onChange={handleChange}
        showUploadList={false}
      >
        <p className="ant-upload-drag-icon">
          <InboxOutlined />
        </p>
        <p className="ant-upload-text">{title}</p>
        <p className="ant-upload-hint">
          {hint || `支持上传 ${accept.mimeLabel} 文件（${accept.extensions.join(', ')}）`}
        </p>
      </Dragger>
    </Card>
  );
};
