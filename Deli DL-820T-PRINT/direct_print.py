import win32print
import win32ui
import win32con
import time
import qrcode
from PIL import Image

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
        qr_data = f"批次:{lines[0].split(':')[1].strip()}"
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
            
    except Exception as e:
        print(f"打印出错: {str(e)}")

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

if __name__ == "__main__":
    test_print()