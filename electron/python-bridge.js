const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');
const axios = require('axios');

class PythonBridge {
  constructor() {
    this.process = null;
    this.port = 8000;
    this.host = '127.0.0.1';
    this.baseUrl = `http://${this.host}:${this.port}`;
    this.maxRetries = 30;
    this.retryDelay = 1000;
  }

  /**
   * 获取 Python 可执行文件路径
   */
  getPythonPath() {
    if (process.env.NODE_ENV === 'development') {
      // 开发环境：使用系统 Python
      return process.platform === 'win32' ? 'python' : 'python3';
    } else {
      // 生产环境：使用打包的 Python 运行时
      const pythonExe = process.platform === 'win32' ? 'python.exe' : 'python';
      return path.join(process.resourcesPath, 'python', pythonExe);
    }
  }

  /**
   * 获取后端代码路径
   */
  getBackendPath() {
    if (process.env.NODE_ENV === 'development') {
      // 开发环境：使用项目中的 backend 目录
      return path.join(__dirname, '..', 'backend');
    } else {
      // 生产环境：使用打包的 backend 目录
      return path.join(process.resourcesPath, 'backend');
    }
  }

  /**
   * 启动 Python 后端
   */
  async start() {
    return new Promise((resolve, reject) => {
      const pythonPath = this.getPythonPath();
      const backendPath = this.getBackendPath();

      console.log('Python 路径:', pythonPath);
      console.log('后端路径:', backendPath);

      // 检查后端目录是否存在
      if (!fs.existsSync(backendPath)) {
        reject(new Error(`后端目录不存在: ${backendPath}`));
        return;
      }

      // 启动 Python 进程
      const args = [
        '-m', 'uvicorn',
        'app.main:app',
        '--host', this.host,
        '--port', this.port.toString()
      ];

      console.log('启动命令:', pythonPath, args.join(' '));

      this.process = spawn(pythonPath, args, {
        cwd: backendPath,
        env: {
          ...process.env,
          PYTHONUNBUFFERED: '1'
        }
      });

      // 监听标准输出
      this.process.stdout.on('data', (data) => {
        const output = data.toString();
        console.log('[Python]', output);

        // 检测启动成功
        if (output.includes('Uvicorn running on') || output.includes('Application startup complete')) {
          this.waitForBackend().then(resolve).catch(reject);
        }
      });

      // 监听标准错误
      this.process.stderr.on('data', (data) => {
        const error = data.toString();
        console.error('[Python Error]', error);

        // 检测端口占用错误
        if (error.includes('Address already in use')) {
          reject(new Error(`端口 ${this.port} 已被占用`));
        }
      });

      // 监听进程退出
      this.process.on('close', (code) => {
        console.log(`Python 进程退出，代码: ${code}`);
        this.process = null;
      });

      // 监听进程错误
      this.process.on('error', (error) => {
        console.error('Python 进程错误:', error);
        reject(error);
      });

      // 超时处理
      setTimeout(() => {
        if (this.process && !this.process.killed) {
          reject(new Error('Python 后端启动超时'));
        }
      }, 30000);
    });
  }

  /**
   * 等待后端就绪
   */
  async waitForBackend() {
    for (let i = 0; i < this.maxRetries; i++) {
      try {
        const response = await axios.get(`${this.baseUrl}/health`, {
          timeout: 1000
        });

        if (response.status === 200) {
          console.log('后端健康检查通过');
          return true;
        }
      } catch (error) {
        // 继续重试
      }

      await new Promise(resolve => setTimeout(resolve, this.retryDelay));
    }

    throw new Error('后端启动失败：健康检查超时');
  }

  /**
   * 停止 Python 后端
   */
  stop() {
    if (this.process) {
      console.log('正在停止 Python 后端...');

      if (process.platform === 'win32') {
        // Windows: 使用 taskkill
        spawn('taskkill', ['/pid', this.process.pid, '/f', '/t']);
      } else {
        // Unix: 发送 SIGTERM
        this.process.kill('SIGTERM');
      }

      this.process = null;
    }
  }

  /**
   * 获取后端状态
   */
  getStatus() {
    return {
      running: this.process !== null && !this.process.killed,
      pid: this.process ? this.process.pid : null,
      url: this.baseUrl
    };
  }

  /**
   * 检查后端健康状态
   */
  async checkHealth() {
    try {
      const response = await axios.get(`${this.baseUrl}/health`, {
        timeout: 3000
      });
      return response.status === 200;
    } catch (error) {
      return false;
    }
  }
}

module.exports = PythonBridge;
