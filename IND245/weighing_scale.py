import serial
import time
import sys
import os
import signal
from loguru import logger
# 添加Flask相关导入
from flask import Flask, jsonify
import threading

class IND245_WeighingScale:
    def __init__(self, port, baudrate=9600, timeout=2, max_retries=5):
        self.ser = None
        self.max_retries = max_retries
        self.port = port
        self.baudrate = baudrate
        self.timeout = timeout
        # 添加一个变量存储最新的重量数据
        self.latest_weight = (None, "未初始化")
        self._release_port()
        self.connect()

    def _release_port(self):
        """释放被占用的串口"""
        try:
            # 由于 Windows 系统不支持 lsof 命令，这里注释掉该部分代码
            # stream = os.popen(f'lsof {self.port}')
            # output = stream.read()
            # stream.close()
            
            # 解析输出找到 PID
            # for line in output.split('\n')[1:]:
            #     if line.strip():
            #         pid = int(line.split()[1])
            #         if pid != os.getpid():
            #             try:
            #                 os.kill(pid, signal.SIGTERM)
            #                 logger.info(f"终止进程 PID: {pid}")
            #                 time.sleep(1)
            #             except ProcessLookupError:
            #                 pass
            return True
        except Exception as e:
            logger.error(f"释放串口失败: {str(e)}")
            return False

    def connect(self):
        retry_count = 0
        while retry_count < self.max_retries:
            try:
                if self.ser and self.ser.is_open:
                    self.ser.close()
                    time.sleep(1)  # 等待串口完全关闭
                
                try:
                    self.ser = serial.Serial(
                        port=self.port,
                        baudrate=self.baudrate,
                        bytesize=serial.EIGHTBITS,
                        parity=serial.PARITY_NONE,
                        stopbits=serial.STOPBITS_ONE,
                        timeout=self.timeout,
                        write_timeout=1
                    )
                except serial.SerialException as e:
                    if 'Resource busy' in str(e):
                        logger.warning("串口被占用，尝试释放...")
                        if self._release_port():
                            continue  # 重试连接
                    raise e
                
                if self.ser.is_open:
                    logger.info(f"成功连接到串口 {self.port}，波特率：{self.baudrate}")
                    return True
                else:
                    raise Exception("串口打开失败")
                    
            except Exception as e:
                retry_count += 1
                logger.warning(f"连接尝试 {retry_count}/{self.max_retries} 失败: {str(e)}")
                if retry_count < self.max_retries:
                    time.sleep(3)  # 增加重试间隔
                else:
                    logger.error(f"连接失败，已达到最大重试次数: {str(e)}")
                    return False

    def get_weight(self):
        """获取重量数据"""
        if not self.ser or not self.ser.is_open:
            logger.error("串口未连接")
            print("串口未连接，尝试重新连接...")
            self.connect()
            return (None, "串口未连接")

        try:
            # 清空缓冲区
            self.ser.reset_input_buffer()
            print("正在读取重量数据...")
            self.timeout = 1  # 调整超时设置为1秒
            # 读取返回数据
            response = self.ser.read_until()  # 使用readline替代read_all，并设置超时时间为1秒
            # print(f"原始数据: {response}")
            # 以\r\x02;0分隔原始数据
            response = response.replace(b' ', b'')
            response_list = [str(item, 'latin-1') for item in response.split(b'00\r\x02;0')]
            if response_list:
                response_list = response_list[1:-1]
            print(f"分隔后的数据: {response_list}")
            # 统计各串出现的次数
            from collections import Counter
            counter = Counter(response_list)
            # 找出出现次数最多的串
            most_common = counter.most_common(1)[0][0]
            weight = float(most_common)
            result = (weight * 0.1, 'kg')
            # 更新最新的重量数据
            self.latest_weight = result
            return result
            
        except serial.SerialTimeoutException:
            error_msg = "串口通信超时"
            print(error_msg)
            logger.error(error_msg)
            return (None, error_msg)
        except serial.SerialException as e:
            error_msg = f"串口通信错误: {str(e)}"
            print(error_msg)
            logger.error(error_msg)
            return (None, error_msg)
        except Exception as e:
            error_msg = f"未知错误: {str(e)}"
            print(error_msg)
            logger.error(error_msg)
            return (None, error_msg)

    def close(self):
        if self.ser and self.ser.is_open:
            self.ser.close()
            logger.info("串口已关闭")
            print("串口已关闭")

# 创建Flask应用
app = Flask(__name__)

# 全局变量存储称重仪实例
scale_instance = None

@app.route('/weight', methods=['GET'])
def get_weight():
    """API端点，返回当前重量数据"""
    if scale_instance:
        weight, unit = scale_instance.latest_weight
        return jsonify({
            'weight': weight,
            'unit': unit,
            'timestamp': time.time()
        })
    else:
        return jsonify({
            'error': '称重仪未初始化',
            'timestamp': time.time()
        }), 500

def start_weight_reading(scale):
    """后台线程函数，持续读取重量数据"""
    try:
        while True:
            scale.get_weight()
            time.sleep(0.5)  # 每0.5秒读取一次数据
    except Exception as e:
        logger.error(f"读取重量数据线程异常: {str(e)}")
    finally:
        scale.close()

def start_http_server(host='0.0.0.0', port=5000):
    """启动HTTP服务器"""
    app.run(host=host, port=port, debug=False, use_reloader=False)

def main():
    global scale_instance
    
    # 配置日志
    logger.remove()  # 移除默认的日志处理器
    logger.add(sys.stderr, level="INFO")  # 添加标准错误输出处理器
    logger.add("weighing_scale.log", rotation="10 MB", level="DEBUG")
    
    # 配置参数
    PORT = 'COM1'  # 串口地址
    BAUDRATE = 9600  # 波特率
    HTTP_HOST = '0.0.0.0'  # 监听所有网络接口
    HTTP_PORT = 5000  # HTTP服务器端口

    print(f"正在连接串口 {PORT}...")
    scale_instance = IND245_WeighingScale(PORT, baudrate=BAUDRATE)
    
    try:
        # 启动重量读取线程
        weight_thread = threading.Thread(target=start_weight_reading, args=(scale_instance,), daemon=True)
        weight_thread.start()
        
        # 启动HTTP服务器
        print(f"启动HTTP服务器在 http://{HTTP_HOST}:{HTTP_PORT}/weight")
        start_http_server(host=HTTP_HOST, port=HTTP_PORT)
            
    except KeyboardInterrupt:
        print("\n程序已停止")
        logger.info("\n用户终止操作")
    finally:
        if scale_instance:
            scale_instance.close()

if __name__ == "__main__":
    main()
