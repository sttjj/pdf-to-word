from reportlab.pdfgen import canvas

def create_test_pdf():
    """
    创建一个简单的测试PDF文件
    """
    c = canvas.Canvas('test.pdf')
    c.drawString(100, 750, "测试PDF文件")
    c.drawString(100, 700, "这是一个用于测试PDF转Word功能的示例文件")
    c.drawString(100, 650, "第一页内容")
    c.showPage()
    c.drawString(100, 750, "第二页")
    c.drawString(100, 700, "这是第二页的内容")
    c.save()
    print("测试PDF文件创建成功: test.pdf")

if __name__ == "__main__":
    create_test_pdf()