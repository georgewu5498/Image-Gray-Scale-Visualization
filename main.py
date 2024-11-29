from PIL import Image
import xlsxwriter
import matplotlib.pyplot as plt

# 读取RGB图像并转换为灰度图
image_path = 'coin.png'  # 替换为你的图片路径
img = Image.open(image_path)
gray_img = img.convert('L')

# 获取图像尺寸
width, height = gray_img.size

# 创建Excel工作簿和工作表
workbook = xlsxwriter.Workbook('coin.xlsx')
worksheet = workbook.add_worksheet()

# 使用matplotlib的autumn颜色映射
autumn_map = plt.get_cmap('autumn')

# 设置列宽和行高，使单元格看起来是正方形的
worksheet.set_column(0, width - 1, 6)  # 设置列宽为1个字符宽度
for row in range(height):
    worksheet.set_row(row, 30)  # 设置行高为20个像素高度

# 将灰度值和坐标存储到Excel，并设置单元格颜色
for row in range(height):
    for col in range(width):
        gray_value = gray_img.getpixel((col, row))
        # 将灰度值归一化到0-1之间
        norm_value = gray_value / 255.0
        # 获取autumn颜色映射中的颜色
        color = autumn_map(norm_value)[:3]  # 取RGB值，忽略透明度
        # 将RGB值转换为十六进制颜色代码
        hex_color = '#{:02X}{:02X}{:02X}'.format(int(color[0] * 255), int(color[1] * 255), int(color[2] * 255))

        # 创建格式对象并设置背景颜色、对齐方式和边框
        format = workbook.add_format({
            'bg_color': hex_color,
            'align': 'center',
            'valign': 'top',  # 垂直居中对齐
            'font_size': 8,  # 设置字体大小为8，可以根据需要调整
            'border': 1,  # 设置边框为1，即细边框
        })

        # 写入灰度值和坐标，并应用格式
        worksheet.write(row, col, f"{col},{row},{gray_value}", format)

# 保存Excel文件
workbook.close()
