import serial
import time
import sys
import os
import signal
from loguru import logger

class IND245_WeighingScale:
    def __init__(self, port, baudrate=9600, timeout=2, max_retries=5):
        self.ser = None
        self.max_retries = max_retries
        self.port = port
        self.baudrate = baudrate
        self.timeout = timeout
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
            return (weight * 0.1, 'kg')
            
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

def main():
    # 配置日志
    logger.remove()  # 移除默认的日志处理器
    logger.add(sys.stderr, level="INFO")  # 添加标准错误输出处理器
    logger.add("weighing_scale.log", rotation="10 MB", level="DEBUG")
    
    # 配置参数
    PORT = 'COM1'  # 串口地址
    BAUDRATE = 9600  # 波特率

    print(f"正在连接串口 {PORT}...")
    scale = IND245_WeighingScale(PORT, baudrate=BAUDRATE)
    
    try:
        print("开始读取重量数据...")
        while True:
            weight, unit = scale.get_weight()
            if weight is not None:
                print(f"当前重量: {weight}{unit}")
            # time.sleep(0.1)  # 增加读取间隔
            
    except KeyboardInterrupt:
        print("\n程序已停止")
        logger.info("\n用户终止操作")
    finally:
        scale.close()

if __name__ == "__main__":
    main()
