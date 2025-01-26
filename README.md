```python

from unotools import Socket, Desktop
import os

def save_slide_as_image(input_ppt, output_folder):
    """使用 unotools 将 PPT 的每一页保存为图片"""
    # 确保输出文件夹存在
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 连接 LibreOffice
    connection = Socket(host="127.0.0.1", port=2002)  # 与 LibreOffice 服务通信
    desktop = Desktop(connection)

    # 打开 PPT 文件
    file_url = f"file://{os.path.abspath(input_ppt)}"
    presentation = desktop.open_document(file_url)

    # 遍历幻灯片并保存为图片
    slides = presentation.getDrawPages()
    for i, slide in enumerate(slides):
        output_path = os.path.join(output_folder, f"slide_{i+1}.png")
        output_url = f"file://{os.path.abspath(output_path)}"

        # 设置导出参数
        filter_data = {"FilterName": "draw_png_Export"}
        presentation.storeToURL(output_url, (filter_data,))
        print(f"Saved slide {i+1} as image: {output_path}")

    # 关闭文档
    presentation.close()

if __name__ == "__main__":
    input_ppt = "your_presentation.pptx"  # 输入的 PPT 文件路径
    output_folder = "output_images"      # 输出的图片文件夹路径
    save_slide_as_image(input_ppt, output_folder)


```
