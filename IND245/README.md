# 小磅秤运行

要运行这个小磅秤配置文件项目，你可以按照以下步骤操作：

### 1. 创建并激活虚拟环境

首先，确保你已经安装了 `venv` 模块。在项目根目录（`f:\小磅秤配置文件`）下打开命令行，执行以下命令来创建并激活虚拟环境，有 python 环境可直接第二步：

```bash
# 创建虚拟环境
python -m venv venv

# 激活虚拟环境（Windows）
.\venv\Scripts\activate

# 激活虚拟环境（Linux/Mac）
source venv/bin/activate
```

### 2. 安装依赖

在激活虚拟环境后，使用 `pip` 安装项目所需的依赖：

```bash
pip install -r requirements.txt
```

### 3. 运行项目

安装完依赖后，就可以运行 `weighing_scale.py` 文件了：

```bash
python weighing_scale.py
```

### 完整操作示例

以下是在 Windows 系统下的完整操作示例：

```bash
# 进入项目目录
cd f:\小磅秤配置文件

# 创建虚拟环境
python -m venv venv

# 激活虚拟环境
.\venv\Scripts\activate

# 安装依赖
pip install -r requirements.txt

# 运行项目
python weighing_scale.py
```

按照以上步骤操作，你就可以成功运行这个小磅秤配置文件项目了。如果在运行过程中遇到问题，请检查串口连接和配置是否正确。

# 设置称重仪服务开机自启并持续运行

要将称重仪服务设置为开机自启并持续运行，我们可以创建一个 Windows 服务。这里我推荐使用 NSSM (Non-Sucking Service Manager)工具来实现这一目标。

## 步骤实施

### 1. 首先，我们需要创建一个启动脚本

创建一个批处理文件，用于启动 Python 脚本：

```batch
@echo off
e:
cd e:\github\InterOfHardware\IND245
python weighing_scale.py
```

### 2. 下载并安装 NSSM 并配置环境变量

NSSM 是一个可以将普通程序注册为 Windows 服务的工具。

```bash
# 下载NSSM (可以从 https://nssm.cc/download 下载)
# 解压后将nssm.exe放在一个固定位置，如C:\Windows\System32
```

**添加到环境变量**

添加 NSSM 目录到环境变量：

- 右键点击"此电脑" → 属性 → 高级系统设置 → 环境变量
- 在"系统变量"中找到 Path 变量并编辑
- 添加 NSSM 所在目录路径 (如 e:\tools\nssm-2.24\win64\ )
- 点击确定保存更改

### 3. 使用 NSSM 创建服务

打开管理员权限的命令提示符，执行以下命令：

```bash
nssm install WeighingScaleService
```

这会打开 NSSM 的图形界面，在界面中设置：

在 Application 选项卡中：

- Path: `C:\Users\Xxiang\AppData\Local\Programs\Python\Python311\python.exe`
- Startup directory: `E:\github\InterOfHardware\IND245`
- Auguments: `E:\github\InterOfHardware\IND245\weighing_scale.py`
- Service name: `WeighingScaleService`

在 Details 选项卡中：

- Display name: `IND245称重仪服务`
- Description: `提供IND245称重仪数据的HTTP服务`
- Startup type: `Automatic`

在 I/O 选项卡中设置日志文件路径：

- Output (stdout): `e:\github\InterOfHardware\IND245\service_output.log`
- Error (stderr): `e:\github\InterOfHardware\IND245\service_error.log`

### 4. 启动服务

```bash
nssm start WeighingScaleService
```

### 5. 验证服务是否正常运行

```bash
nssm status WeighingScaleService # 查看服务状态
# nssm start WeighingScaleService # 启动服务
# nssm stop WeighingScaleService # 停止服务
# nssm remove WeighingScaleService # 移除服务
# nssm restart WeighingScaleService # 重启服务
# nssm edit WeighingScaleService # 编辑服务
```

可以通过访问 http://localhost:5000/weight 来验证服务是否正常运行。

### 6. 修改代码以提高服务稳定性

为了使服务更加稳定，我建议对代码进行一些修改：

```python:e:\github\InterOfHardware\IND245\weighing_scale.py
# ... 现有代码 ...

def start_weight_reading(scale):
    """后台线程函数，持续读取重量数据"""
    try:
        while True:
            try:
                scale.get_weight()
            except Exception as e:
                logger.error(f"读取重量数据异常: {str(e)}")
                # 尝试重新连接
                try:
                    scale.connect()
                except:
                    pass
            time.sleep(0.5)  # 每0.5秒读取一次数据
    except Exception as e:
        logger.error(f"读取重量数据线程异常: {str(e)}")
    finally:
        scale.close()

# ... 现有代码 ...

def main():
    global scale_instance

    # 配置日志
    logger.remove()  # 移除默认的日志处理器
    logger.add(sys.stderr, level="INFO")  # 添加标准错误输出处理器
    log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "weighing_scale.log")
    logger.add(log_path, rotation="10 MB", level="DEBUG")

    # 配置参数
    PORT = 'COM1'  # 串口地址
    BAUDRATE = 9600  # 波特率
    HTTP_HOST = '0.0.0.0'  # 监听所有网络接口
    HTTP_PORT = 5000  # HTTP服务器端口

    logger.info("称重仪服务启动中...")
    print(f"正在连接串口 {PORT}...")
    scale_instance = IND245_WeighingScale(PORT, baudrate=BAUDRATE)

    try:
        # 启动重量读取线程
        weight_thread = threading.Thread(target=start_weight_reading, args=(scale_instance,), daemon=True)
        weight_thread.start()

        # 启动HTTP服务器
        logger.info(f"启动HTTP服务器在 http://{HTTP_HOST}:{HTTP_PORT}/weight")
        print(f"启动HTTP服务器在 http://{HTTP_HOST}:{HTTP_PORT}/weight")
        start_http_server(host=HTTP_HOST, port=HTTP_PORT)

    except KeyboardInterrupt:
        print("\n程序已停止")
        logger.info("用户终止操作")
    except Exception as e:
        logger.error(f"服务异常: {str(e)}")
    finally:
        if scale_instance:
            scale_instance.close()

# ... 现有代码 ...
```

## 其他建议

1. **监控服务状态**：可以使用 Windows 任务计划程序定期检查服务是否正常运行，如果不正常则重启服务。

2. **设置自动重启**：在 NSSM 的设置中，可以配置服务崩溃后自动重启的选项。

3. **防火墙设置**：确保 Windows 防火墙允许 5000 端口的访问，以便其他设备可以访问称重数据。

4. **开机延迟启动**：如果需要，可以设置服务延迟启动，确保其他必要服务先启动。

这样设置后，称重仪服务将在系统启动时自动运行，并持续提供 HTTP 接口供其他应用程序访问称重数据。
