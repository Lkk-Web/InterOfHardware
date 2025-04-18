import win32print
import win32ui
import win32con
import time
import qrcode
from PIL import Image
from flask import Flask, request, jsonify
import logging
import os
from datetime import datetime

# 创建Flask应用
app = Flask(__name__)

# 配置日志
class CustomFormatter(logging.Formatter):
    def format(self, record):
        if "Running on" in record.getMessage():
            return "="*50 + "\n标签打印服务已成功启动!\n" + \
                   f"API地址: http://0.0.0.0:5001/print\n" + \
                   "支持POST请求，请使用正确的参数格式\n" + "="*50
        return super().format(record)

# 应用自定义日志格式
logger = logging.getLogger('werkzeug')
handler = logging.StreamHandler()
handler.setFormatter(CustomFormatter())
logger.handlers = [handler]

# 创建打印日志记录器
print_logger = logging.getLogger('print_logger')
print_logger.setLevel(logging.INFO)

# 确保日志目录存在
log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs')
os.makedirs(log_dir, exist_ok=True)

# 创建日志文件处理器，按日期命名
log_file = os.path.join(log_dir, f'print_log_{datetime.now().strftime("%Y%m%d")}.log')
file_handler = logging.FileHandler(log_file, encoding='utf-8')
file_handler.setLevel(logging.INFO)

# 设置日志格式
log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
file_handler.setFormatter(log_formatter)
print_logger.addHandler(file_handler)

# 打印机名称
PRINTER_NAME = 'Deli DL-820T(NEW)'
# 添加一个标志，用于控制是否强制要求打印机存在
REQUIRE_PRINTER = True  # 设置为False时，即使打印机不存在也不会报错

def check_printer():
    """检查打印机是否存在且已连接"""
    # 获取所有打印机 - 修改为正确枚举打印机
    printers = [printer[2] for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
    
    print("当前连接的打印机列表：")
    for p in printers:
        print(p)
    
    if PRINTER_NAME not in printers:
        if REQUIRE_PRINTER:
            raise Exception(f"找不到已连接的打印机: {PRINTER_NAME}")
        else:
            print(f"警告: 找不到已连接的打印机 {PRINTER_NAME}，但将继续运行服务")
            print(f"请在需要打印时连接打印机")
            return False
    return True

def print_label(text):
    """使用Windows GDI打印"""
    try:
        printer_available = check_printer()
        if not printer_available and not REQUIRE_PRINTER:
            print(f"模拟打印: {text[:50]}...")
            return True  # 返回成功，但实际上没有打印
        
        # 获取打印机状态并检查
        hprinter = win32print.OpenPrinter(PRINTER_NAME)
        printer_info = win32print.GetPrinter(hprinter, 2)
        printer_status = printer_info['Status']
        
        # 检查打印机状态
        if printer_status != 0:  # 0表示正常
            status_codes = {
                win32print.PRINTER_STATUS_PAUSED: "打印机已暂停",
                win32print.PRINTER_STATUS_ERROR: "打印机错误",
                win32print.PRINTER_STATUS_PAPER_JAM: "打印机卡纸",
                win32print.PRINTER_STATUS_PAPER_OUT: "打印机缺纸",
                win32print.PRINTER_STATUS_PAPER_PROBLEM: "打印机纸张问题",
                win32print.PRINTER_STATUS_OFFLINE: "打印机离线",
                win32print.PRINTER_STATUS_USER_INTERVENTION: "打印机需要用户干预",
                win32print.PRINTER_STATUS_DOOR_OPEN: "打印机门已打开"
            }
            
            status_msg = ""
            for code, msg in status_codes.items():
                if printer_status & code:
                    status_msg += msg + ", "
            
            error_msg = f"打印机状态异常: {status_msg[:-2] if status_msg else '未知错误'}"
            print(error_msg)
            win32print.ClosePrinter(hprinter)
            return False
        
        # 清理打印队列中的错误作业
        jobs = win32print.EnumJobs(hprinter, 0, 999)
        for job in jobs:
            status = job.get('Status', 0)
            if status & win32print.JOB_STATUS_ERROR:
                print(f"发现错误作业，正在清理: {job['pDocument']}")
                try:
                    win32print.SetJob(hprinter, job['JobId'], 0, None, win32print.JOB_CONTROL_DELETE)
                except:
                    pass
        
        dc = win32ui.CreateDC()
        dc.CreatePrinterDC(PRINTER_NAME)
        
        dc.StartDoc('Label Print')
        dc.StartPage()
        dc.SetMapMode(win32con.MM_TEXT)
        
        # 创建两种字体：普通和加粗
        normal_font = win32ui.CreateFont({
            'name': '宋体',
            'height': 28,
            'weight': 700,
            'width': 0,
            'quality': win32con.PROOF_QUALITY
        })
        
        bold_font = win32ui.CreateFont({
            'name': '宋体',
            'height': 40,
            'weight': 700,  # 加粗版本
            'width': 0,
            'quality': win32con.PROOF_QUALITY
        })
        
        # 打印文本
        lines = text.split('\n')
        y = 30
        x = 30
        line_spacing = 42
        
        for i, line in enumerate(lines):
            if i == 0:  # 第一行使用加粗字体
                dc.SelectObject(bold_font)
            else:
                dc.SelectObject(normal_font)
            dc.TextOut(x, y, line)
            y += line_spacing
        
        # 生成二维码
        qr = qrcode.QRCode(version=1, box_size=2, border=1)
        # 修改二维码数据，使用JSON格式
        package_info = lines[1].split('：')[1].strip() if len(lines) > 1 else ""
        qr_data = f'{{"packageCode":"{package_info}"}}'  # 扫描后得出的数据，JSON格式
        qr.add_data(qr_data)
        qr.make(fit=True)
        qr_img = qr.make_image(fill_color="black", back_color="white")
        
        # 调整二维码大小并转换为黑白图像
        qr_img = qr_img.resize((100, 100))
        qr_img = qr_img.convert('1')  # 转换为黑白图像
        
        # 使用矩形绘制二维码
        qr_x = 270  # 二维码左上角X坐标
        qr_y = 120  # 二维码左上角Y坐标
        dot_size = 2  # 每个点的大小
        
        # 创建黑色画笔和画刷
        black_pen = win32ui.CreatePen(win32con.PS_SOLID, 1, 0)
        black_brush = win32ui.CreateBrush(win32con.BS_SOLID, 0, 0)
        
        old_pen = dc.SelectObject(black_pen)
        old_brush = dc.SelectObject(black_brush)
        
        # 直接在DC上绘制二维码
        for y_pos in range(qr_img.height):
            for x_pos in range(qr_img.width):
                if qr_img.getpixel((x_pos, y_pos)) == 0:  # 黑色像素
                    # 使用矩形代替像素
                    rect = (qr_x + x_pos*dot_size, qr_y + y_pos*dot_size, 
                           qr_x + (x_pos+1)*dot_size, qr_y + (y_pos+1)*dot_size)
                    dc.Rectangle(rect)
        
        # 恢复原来的画笔和画刷
        dc.SelectObject(old_pen)
        dc.SelectObject(old_brush)
        
        # 结束打印
        dc.EndPage()
        dc.EndDoc()
        
        # 释放资源
        del dc
        win32print.ClosePrinter(hprinter)
        
        # 等待打印就绪
        time.sleep(1)
        
        print(f"打印成功: {package_info}")
        return True
            
    except Exception as e:
        print(f"打印出错: {str(e)}")
        return False

def test_print():
    """测试打印"""
    # 测试商品标签（优化布局）
    product_label = (
        '批次:WGC2025031801\n'
        '包号：WGC2025031801-001\n'
        '\n'  # 包号后添加空行
        '品名：Z0118\n'
        '厂家：RY\n'
        '日期：2025-4-14\n'
        '车牌：粤A12345\n'
    )
    print_label(product_label)

    # 测试带备注的标签
    product_label_with_desc = (
        '批次:WGC2025031801\n'
        '包号：WGC2025031801-002\n'
        '品名：Z0118\n'
        '备注：测试备注信息\n'
        '厂家：RY\n'
        '日期：2025-4-14\n'
        '车牌：粤A12345\n'
    )
    print_label(product_label_with_desc)

@app.route('/print', methods=['POST'])
def batch_print():
    """接收POST请求进行批量打印
    
    参数:
    - dto: 包含批次、品名、厂家、日期、车牌的字典
    - start_package: 起始包号
    - end_package: 结束包号
    """
    try:
        # 首先检查打印机状态
        try:
            check_printer()
        except Exception as e:
            error_msg = f'打印机检查失败: {str(e)}'
            print_logger.error(error_msg)
            return jsonify({'success': False, 'message': error_msg}), 503
            
        data = request.json
        
        # 验证必要参数
        if not data or 'dto' not in data or 'start_package' not in data or 'end_package' not in data:
            error_msg = '缺少必要参数'
            print_logger.error(error_msg)
            return jsonify({'success': False, 'message': error_msg}), 400
        
        dto = data['dto']
        start_package = int(data['start_package'])
        end_package = int(data['end_package'])
        
        # 验证DTO中的必要字段
        required_fields = ['batch', 'product_name', 'manufacturer', 'date', 'license_plate']
        for field in required_fields:
            if field not in dto:
                error_msg = f'DTO缺少{field}字段'
                print_logger.error(error_msg)
                return jsonify({'success': False, 'message': error_msg}), 400
        
        # 验证包号范围
        if start_package > end_package:
            error_msg = '起始包号不能大于结束包号'
            print_logger.error(error_msg)
            return jsonify({'success': False, 'message': error_msg}), 400
        
        # 记录打印任务开始
        print_logger.info(f"========== 打印任务开始 ==========")
        print_logger.info(f"批次: {dto['batch']}, 包号范围: {start_package}-{end_package}")
        print_logger.info(f"品名: {dto['product_name']}, 厂家: {dto['manufacturer']}")
        # 记录备注信息（如果有）
        if 'description' in dto and dto['description']:
            print_logger.info(f"备注: {dto['description']}")
        print_logger.info(f"日期: {dto['date']}, 车牌: {dto['license_plate']}")
        
        # 批量打印
        success_count = 0
        success_packages = []  # 记录成功的包号
        failed_packages = []  # 记录失败的包号
        
        # 检查打印机状态，如有必要清空打印队列
        try:
            hprinter = win32print.OpenPrinter(PRINTER_NAME)
            jobs = win32print.EnumJobs(hprinter, 0, 999)
            job_count = len(jobs)
            print(f"当前打印队列中有 {job_count} 个作业")
            print_logger.info(f"打印前检查: 打印队列中有 {job_count} 个作业")
            win32print.ClosePrinter(hprinter)
        except Exception as e:
            error_msg = f"检查打印队列失败: {str(e)}"
            print(error_msg)
            print_logger.error(error_msg)

        # 顺序处理所有包号
        for package_num in range(start_package, end_package + 1):
            # 格式化包号为三位数，如001、002、003等
            formatted_package_num = f"{package_num:03d}"
            
            # 基本标签内容
            label_parts = [
                f'批次:{dto["batch"]}',
                f'包号：{dto["batch"]}-{formatted_package_num}'
            ]
            
            # 如果有description，直接添加品名和备注；否则在包号后添加空行再添加品名
            if 'description' in dto and dto['description']:
                label_parts.extend([
                    f'品名：{dto["product_name"]}',
                    f'备注：{dto["description"]}'
                ])
            else:
                label_parts.extend([
                    '',  # 包号后添加空行
                    f'品名：{dto["product_name"]}'
                ])
            
            # 添加其余内容
            label_parts.extend([
                f'厂家：{dto["manufacturer"]}',
                f'日期：{dto["date"]}',
                f'车牌：{dto["license_plate"]}'
            ])
            
            # 合并成最终标签文本
            label_text = '\n'.join(label_parts) + '\n'
            
            package_id = f"{dto['batch']}-{formatted_package_num}"
            print(f"正在打印包号: {package_id}...")
            print_logger.info(f"正在打印包号: {package_id}")
            
            # 每个包号最多尝试3次
            max_retries = 3
            success = False
            
            for retry in range(max_retries):
                if print_label(label_text):
                    success_count += 1
                    success_packages.append(package_num)
                    print_logger.info(f"包号 {package_id} 打印成功")
                    time.sleep(2)  # 打印间隔
                    success = True
                    break  # 打印成功，跳出重试循环
                else:
                    retry_msg = f"包号 {package_id} 打印失败，尝试 {retry+1}/{max_retries}"
                    print(retry_msg)
                    print_logger.warning(retry_msg)
                    time.sleep(1 + retry)  # 每次重试增加等待时间
            
            # 记录失败的包号        
            if not success:
                failed_msg = f"包号 {package_id} 打印失败，已达到最大重试次数"
                print(failed_msg)
                print_logger.error(failed_msg)
                failed_packages.append(package_num)
                
                # 每5个失败的包号，打印一次摘要，避免忘记处理
                if len(failed_packages) % 5 == 0:
                    summary_msg = f"已有 {len(failed_packages)} 个包号打印失败: {failed_packages[-5:]}"
                    print(summary_msg)
                    print_logger.warning(summary_msg)
        
        # 第一轮打印摘要
        first_round_msg = f"第一轮打印完成: 总计 {end_package - start_package + 1} 个包号, 成功 {success_count} 个, 失败 {len(failed_packages)} 个"
        print(first_round_msg)
        print_logger.info(first_round_msg)
        
        # 如果有失败的包号，尝试重新打印一次
        if failed_packages:
            retry_msg = f"开始重新打印失败的包号: {failed_packages}"
            print(retry_msg)
            print_logger.info(retry_msg)
            
            retry_success = 0
            retry_success_list = []
            still_failed = []
            
            for package_num in failed_packages:
                # 格式化包号为三位数
                formatted_package_num = f"{package_num:03d}"
                
                # 基本标签内容
                label_parts = [
                    f'批次:{dto["batch"]}',
                    f'包号：{dto["batch"]}-{formatted_package_num}'
                ]
                
                # 如果有description，直接添加品名和备注；否则在包号后添加空行再添加品名
                if 'description' in dto and dto['description']:
                    label_parts.extend([
                        f'品名：{dto["product_name"]}',
                        f'备注：{dto["description"]}'
                    ])
                else:
                    label_parts.extend([
                        '',  # 包号后添加空行
                        f'品名：{dto["product_name"]}'
                    ])
                
                # 添加其余内容
                label_parts.extend([
                    f'厂家：{dto["manufacturer"]}',
                    f'日期：{dto["date"]}',
                    f'车牌：{dto["license_plate"]}'
                ])
                
                # 合并成最终标签文本
                label_text = '\n'.join(label_parts) + '\n'
                
                package_id = f"{dto['batch']}-{formatted_package_num}"
                print(f"重新打印包号: {package_id}")
                print_logger.info(f"重新打印包号: {package_id}")
                
                # 重新打印时给更多重试机会和更长等待时间
                time.sleep(3)  # 较长的初始间隔
                if print_label(label_text):
                    retry_success += 1
                    retry_success_list.append(package_num)
                    success_packages.append(package_num)
                    print_logger.info(f"包号 {package_id} 重新打印成功")
                    time.sleep(3)  # 更长的间隔
                else:
                    final_fail_msg = f"包号 {package_id} 重新打印仍然失败"
                    print(final_fail_msg)
                    print_logger.error(final_fail_msg)
                    still_failed.append(package_num)
                    time.sleep(1)  # 短暂间隔
            
            # 更新统计信息
            success_count += retry_success
            failed_packages = still_failed
            
            # 第二轮打印摘要
            second_round_msg = f"第二轮打印完成: 总计重试 {len(retry_success_list) + len(still_failed)} 个包号, 成功 {retry_success} 个, 仍然失败 {len(still_failed)} 个"
            print(second_round_msg)
            print_logger.info(second_round_msg)
            
            if retry_success > 0:
                print_logger.info(f"第二轮成功打印的包号: {retry_success_list}")
        
        # 返回最终结果
        result = {
            'success': True,
            'total': end_package - start_package + 1,
            'success_count': success_count,
            'failed_count': len(failed_packages),
            'failed_packages': failed_packages,  # 返回最终未成功打印的包号
            'success_packages': success_packages  # 返回所有成功打印的包号
        }
        
        # 记录最终结果到日志
        print_logger.info(f"========== 打印任务最终结果 ==========")
        print_logger.info(f"批次: {dto['batch']}")
        print_logger.info(f"包号范围: {start_package} - {end_package}")
        print_logger.info(f"总计打印: {result['total']} 个标签")
        print_logger.info(f"成功打印: {result['success_count']} 个标签")
        print_logger.info(f"失败打印: {result['failed_count']} 个标签")
        
        if result['failed_count'] > 0:
            print_logger.warning(f"失败的包号列表: {failed_packages}")
            # 将失败的包号记录到单独的文件，方便后续处理
            try:
                failed_log_file = os.path.join(log_dir, f'failed_packages_{dto["batch"]}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log')
                with open(failed_log_file, 'w', encoding='utf-8') as f:
                    f.write(f"批次: {dto['batch']}\n")
                    f.write(f"日期: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                    f.write(f"失败包号数量: {result['failed_count']}\n")
                    f.write("失败包号列表:\n")
                    for pkg in failed_packages:
                        # 使用三位数格式化失败的包号
                        formatted_pkg = f"{pkg:03d}"
                        f.write(f"{dto['batch']}-{formatted_pkg}\n")
                print_logger.info(f"失败包号已记录到文件: {failed_log_file}")
            except Exception as e:
                print_logger.error(f"记录失败包号到文件时出错: {str(e)}")
        else:
            print_logger.info("所有包号均打印成功")
        
        print_logger.info(f"========== 打印任务结束 ==========")
        
        return jsonify(result)
    
    except Exception as e:
        error_msg = str(e)
        print_logger.error(f"打印过程发生异常: {error_msg}")
        return jsonify({'success': False, 'message': error_msg}), 500

if __name__ == "__main__": 
    # 记录服务启动
    print_logger.info("标签打印服务启动")
    app.run(host='0.0.0.0', port=5001, debug=False)
    #test_print()
    # 启动Flask应用
    # 如果只想测试打印功能，可以注释上面的app.run()，取消注释下面的test_print()