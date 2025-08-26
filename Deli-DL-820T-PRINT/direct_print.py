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

# 配置日志（简化，仅保留关键日志）
class CustomFormatter(logging.Formatter):
    def format(self, record):
        if "Running on" in record.getMessage():
            return "="*50 + "\n标签打印服务已成功启动!\n" + \
                   f"API地址: http://0.0.0.0:5001/print\n" + \
                   "支持POST请求，请使用正确的参数格式\n" + "="*50
        return super().format(record)

logger = logging.getLogger('werkzeug')
handler = logging.StreamHandler()
handler.setFormatter(CustomFormatter())
logger.handlers = [handler]

# 打印日志记录器
print_logger = logging.getLogger('print_logger')
print_logger.setLevel(logging.INFO)

log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs')
os.makedirs(log_dir, exist_ok=True)

log_file = os.path.join(log_dir, f'print_log_{datetime.now().strftime("%Y%m%d")}.log')
file_handler = logging.FileHandler(log_file, encoding='utf-8')
file_handler.setLevel(logging.INFO)

log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
file_handler.setFormatter(log_formatter)
print_logger.addHandler(file_handler)

# 打印机名称
PRINTER_NAME = 'Deli DL-820T(NEW)'
REQUIRE_PRINTER = True

def check_printer():
    printers = [printer[2] for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
    if PRINTER_NAME not in printers:
        if REQUIRE_PRINTER:
            raise Exception(f"找不到已连接的打印机: {PRINTER_NAME}")
        else:
            print(f"警告: 找不到打印机 {PRINTER_NAME}，但继续运行")
            return False
    return True

def print_label(text):
    try:
        if not check_printer():
            print("模拟打印（无真实打印机）:", text[:30])
            return True

        hprinter = win32print.OpenPrinter(PRINTER_NAME)
        printer_info = win32print.GetPrinter(hprinter, 2)
        if printer_info['Status'] != 0:
            win32print.ClosePrinter(hprinter)
            return False

        dc = win32ui.CreateDC()
        dc.CreatePrinterDC(PRINTER_NAME)
        dc.StartDoc('Label Print')
        dc.StartPage()
        dc.SetMapMode(win32con.MM_TEXT)

        normal_font = win32ui.CreateFont({'name': '宋体', 'height': 28, 'weight': 700})
        bold_font = win32ui.CreateFont({'name': '宋体', 'height': 40, 'weight': 700})

        lines = text.split('\n')
        y, x, line_spacing = 30, 30, 42

        for i, line in enumerate(lines):
            font = bold_font if i == 0 else normal_font
            dc.SelectObject(font)
            dc.TextOut(x, y, line)
            y += line_spacing

        # 二维码（简化异常处理）
        try:
            if len(lines) > 1:
                package_info = lines[1].split('：')[1].strip()
                qr = qrcode.QRCode(version=1, box_size=2, border=1)
                qr_data = f'{{"packageCode":"{package_info}"}}'
                qr.add_data(qr_data)
                qr.make(fit=True)
                qr_img = qr.make_image(fill_color="black", back_color="white").convert('1').resize((100, 100))

                qr_x, qr_y, dot_size = 270, 120, 2
                black_pen = win32ui.CreatePen(win32con.PS_SOLID, 1, 0)
                black_brush = win32ui.CreateBrush(win32con.BS_SOLID, 0, 0)
                old_pen = dc.SelectObject(black_pen)
                old_brush = dc.SelectObject(black_brush)

                for y_pos in range(qr_img.height):
                    for x_pos in range(qr_img.width):
                        if qr_img.getpixel((x_pos, y_pos)) == 0:
                            rect = (qr_x + x_pos * dot_size, qr_y + y_pos * dot_size,
                                    qr_x + (x_pos + 1) * dot_size, qr_y + (y_pos + 1) * dot_size)
                            dc.Rectangle(rect)
                dc.SelectObject(old_pen)
                dc.SelectObject(old_brush)
        except Exception as e:
            print_logger.error(f"二维码生成/打印失败: {e}")

        dc.EndPage()
        dc.EndDoc()
        del dc
        win32print.ClosePrinter(hprinter)
        time.sleep(1)
        return True
    except Exception as e:
        print_logger.error(f"打印失败: {e}")
        return False

@app.route('/print', methods=['POST'])
def batch_print():
    try:
        check_printer()

        data = request.json
        if not data or 'dto' not in data or 'start_package' not in data or 'end_package' not in data:
            return jsonify({'success': False, 'message': '缺少必要参数'}), 400

        dto = data['dto']
        start_package = int(data['start_package'])
        end_package = int(data['end_package'])

        required_fields = ['batch', 'product_name', 'manufacturer', 'date', 'license_plate']
        for field in required_fields:
            if field not in dto:
                return jsonify({'success': False, 'message': f'Missing field: {field}'}), 400

        if start_package > end_package:
            return jsonify({'success': False, 'message': '起始包号不能大于结束包号'}), 400

        print_logger.info(f"========== 打印任务开始 ==========")
        print_logger.info(f"批次: {dto['batch']}, 包号范围: {start_package} - {end_package}")

        success_count = 0
        success_packages = []
        failed_packages = []

        for package_num in range(start_package, end_package + 1):
            formatted_num = f"{package_num:03d}"
            package_id = f"{dto['batch']}-{formatted_num}"

            label_parts = [
                f'批次:{dto["batch"]}',
                f'包号：{dto["batch"]}-{formatted_num}'
            ]
            if 'description' in dto:
                label_parts.extend([f'品名：{dto["product_name"]}', f'备注：{dto["description"]}'])
            else:
                label_parts.extend(['', f'品名：{dto["product_name"]}'])
            label_parts.extend([
                f'厂家：{dto["manufacturer"]}',
                f'日期：{dto["date"]}',
                f'车牌：{dto["license_plate"]}'
            ])
            label_text = '\n'.join(label_parts) + '\n'

            print(f"正在打印包号: {package_id}")
            print_logger.info(f"正在打印包号: {package_id}")

            success = print_label(label_text)
            if success:
                success_count += 1
                success_packages.append(package_num)
                print_logger.info(f"✅ 包号 {package_id} 打印成功")
            else:
                failed_packages.append(package_num)
                print_logger.error(f"❌ 包号 {package_id} 打印失败")

            time.sleep(1)  # 避免打印机过载

        # 返回结果
        return jsonify({
            'success': True,
            'total': end_package - start_package + 1,
            'success_count': success_count,
            'failed_count': len(failed_packages),
            'success_packages': success_packages,
            'failed_packages': failed_packages
        })

    except Exception as e:
        print_logger.error(f"打印任务异常: {e}")
        return jsonify({'success': False, 'message': str(e)}), 500

if __name__ == '__main__':
    print_logger.info("标签打印服务启动")
    app.run(host='0.0.0.0', port=5001, debug=False)