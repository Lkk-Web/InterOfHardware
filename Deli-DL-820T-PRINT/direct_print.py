import win32print
import win32ui
import win32con
import time
import qrcode
from PIL import Image
from flask import Flask, request, jsonify
import logging

# 创建Flask应用
app = Flask(__name__)

# 配置日志
class CustomFormatter(logging.Formatter):
    def format(self, record):
        if "Running on" in record.getMessage():
            return "="*50 + "\n标签打印服务已成功启动!\n" + \
                   f"API地址: http://0.0.0.0:5010/print\n" + \
                   "支持POST请求，请使用正确的参数格式\n" + "="*50
        return super().format(record)

# 应用自定义日志格式
logger = logging.getLogger('werkzeug')
handler = logging.StreamHandler()
handler.setFormatter(CustomFormatter())
logger.handlers = [handler]

# 打印机名称
PRINTER_NAME = 'Deli DL-820T(NEW)'

def check_printer():
    """检查打印机是否存在"""
    printers = [printer[2] for printer in win32print.EnumPrinters(2)]
    print("系统中的打印机列表：")
    for p in printers:
        print(p)
    
    if PRINTER_NAME not in printers:
        raise Exception(f"找不到打印机: {PRINTER_NAME}")
    return True

def print_label(text):
    """使用Windows GDI打印"""
    try:
        check_printer()
        hprinter = win32print.OpenPrinter(PRINTER_NAME)
        printer_info = win32print.GetPrinter(hprinter, 2)
        
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
        
        del dc
        win32print.ClosePrinter(hprinter)
        print("打印成功")
        return True
            
    except Exception as e:
        print(f"打印出错: {str(e)}")
        return False

def test_print():
    """测试打印"""
    # 测试商品标签（优化布局）
    product_label = (
        '批次:WGC2025031801\n'
        '包号：WGC2025031801-100\n'
        '\n'
        '品名：Z0118\n'
        '厂家：RY\n'
        '日期：2025-3-17\n'
        '车牌：粤A12345\n'
    )
    print_label(product_label)

@app.route('/print', methods=['POST'])
def batch_print():
    """接收POST请求进行批量打印
    
    参数:
    - dto: 包含批次、品名、厂家、日期、车牌的字典
    - start_package: 起始包号
    - end_package: 结束包号
    """
    try:
        data = request.json
        
        # 验证必要参数
        if not data or 'dto' not in data or 'start_package' not in data or 'end_package' not in data:
            return jsonify({'success': False, 'message': '缺少必要参数'}), 400
        
        dto = data['dto']
        start_package = int(data['start_package'])
        end_package = int(data['end_package'])
        
        # 验证DTO中的必要字段
        required_fields = ['batch', 'product_name', 'manufacturer', 'date', 'license_plate']
        for field in required_fields:
            if field not in dto:
                return jsonify({'success': False, 'message': f'DTO缺少{field}字段'}), 400
        
        # 验证包号范围
        if start_package > end_package:
            return jsonify({'success': False, 'message': '起始包号不能大于结束包号'}), 400
        
        # 批量打印
        success_count = 0
        failed_packages = []
        
        for package_num in range(start_package, end_package + 1):
            # 构建标签文本
            label_text = (
                f'批次:{dto["batch"]}\n'
                f'包号：{dto["batch"]}-{package_num}\n'
                f'\n'
                f'品名：{dto["product_name"]}\n'
                f'厂家：{dto["manufacturer"]}\n'
                f'日期：{dto["date"]}\n'
                f'车牌：{dto["license_plate"]}\n'
            )
            
            # 打印标签
            if print_label(label_text):
                success_count += 1
            else:
                failed_packages.append(package_num)
            
            # 打印间隔，避免打印机缓冲区溢出
            time.sleep(0.5)
        
        # 返回结果
        result = {
            'success': True,
            'total': end_package - start_package + 1,
            'success_count': success_count,
            'failed_count': len(failed_packages),
            'failed_packages': failed_packages
        }
        
        return jsonify(result)
    
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500

if __name__ == "__main__":  
    # 启动Flask应用
    app.run(host='0.0.0.0', port=5012, debug=False)
    # 如果只想测试打印功能，可以注释上面的app.run()，取消注释下面的test_print()
    # test_print()