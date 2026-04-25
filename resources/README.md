# 资源文件说明

## 图标文件

请在此目录放置应用图标：

### Windows
- `icon.ico` - Windows 应用图标（256x256 或更高）

### macOS
- `icon.icns` - macOS 应用图标

### Linux
- `icon.png` - Linux 应用图标（512x512 PNG）

## 图标生成工具

可以使用以下工具生成多平台图标：

1. **在线工具**
   - https://www.icoconverter.com/
   - https://cloudconvert.com/

2. **命令行工具**
   ```bash
   # 安装 electron-icon-builder
   npm install -g electron-icon-builder

   # 从 PNG 生成所有平台图标
   electron-icon-builder --input=./icon.png --output=./resources
   ```

3. **手动创建**
   - Windows: 使用 GIMP 或 Photoshop 导出为 .ico
   - macOS: 使用 iconutil 或 Image2Icon
   - Linux: 直接使用 PNG 格式

## 图标尺寸要求

- Windows ICO: 16x16, 32x32, 48x48, 64x64, 128x128, 256x256
- macOS ICNS: 16x16, 32x32, 64x64, 128x128, 256x256, 512x512, 1024x1024
- Linux PNG: 512x512 (推荐)

## 临时方案

如果暂时没有图标，electron-builder 会使用默认图标。
建议尽快添加自定义图标以提升应用专业度。
