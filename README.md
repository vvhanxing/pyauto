```python

import uno
import os

def connect_to_libreoffice():
    """连接到运行中的 LibreOffice"""
    local_context = uno.getComponentContext()
    resolver = local_context.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_context)
    context = resolver.resolve("uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext")
    desktop = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.Desktop", context)
    return desktop

def save_slide_as_image(input_ppt, output_folder):
    """将 PPT 的每一页保存为图片"""
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 连接到 LibreOffice
    desktop = connect_to_libreoffice()

    # 打开 PPT 文件
    file_url = uno.systemPathToFileUrl(os.path.abspath(input_ppt))
    presentation = desktop.loadComponentFromURL(file_url, "_blank", 0, ())

    # 遍历幻灯片并保存为图片
    draw_pages = presentation.getDrawPages()
    for i in range(draw_pages.getCount()):
        slide = draw_pages.getByIndex(i)
        output_path = os.path.join(output_folder, f"slide_{i+1}.png")
        output_url = uno.systemPathToFileUrl(os.path.abspath(output_path))

        # 导出当前幻灯片为图片
        args = (uno.createUnoStruct("com.sun.star.beans.PropertyValue"),)
        args[0].Name = "FilterName"
        args[0].Value = "draw_png_Export"
        slide.storeToURL(output_url, args)
        print(f"Saved slide {i+1} as image: {output_path}")

    # 关闭文件
    presentation.close(True)

if __name__ == "__main__":
    input_ppt = "your_presentation.pptx"  # 输入的 PPT 文件路径
    output_folder = "output_images"      # 输出的图片文件夹路径
    save_slide_as_image(input_ppt, output_folder)

```
