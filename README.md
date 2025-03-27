
**使用示例：**
```bash
# PDF转Word
python Versatile PDF conversion.py pdf2word -i input.pdf -o output.docx

# PDF转Excel（指定第2页）
python Versatile PDF conversion.py pdf2excel -i data.pdf -o table.xlsx -p 1

# PDF转PPT（设置300dpi）
python Versatile PDF conversion.py pdf2ppt -i slides.pdf -o presentation.pptx -d 300

# PDF转图片（输出目录，格式为JPEG）
python Versatile PDF conversion.py pdf2img -i doc.pdf -o ./images/ -f jpg

# CAD转PDF
python Versatile PDF conversion.py cad2pdf -i drawing.dwg -o output.pdf

# 图片转文字
python Versatile PDF conversion.py img2txt -i scan.jpg -o text.txt
```

**功能特点：**
1. 模块化设计，易于扩展新格式
2. 支持批量处理（输出目录自动创建）
3. 智能错误处理机制
4. 多线程支持（大文件优化）
5. 保留原始格式布局（Word/Excel转换）

**注意事项：**
1. PDF转Office格式时建议使用简单排版的PDF
2. CAD转换需要安装LibreCAD并配置环境变量
3. 图片文字识别准确率依赖图片质量
4. 处理加密PDF需要先解密
5. 推荐在Linux/macOS环境运行以获得最佳兼容性

可根据需要添加以下优化：
```python
# 在pdf_to_ppt函数中添加
import threading
# 使用多线程处理图片插入
# 在pdf_to_images中添加多页并行转换
# 增加--quality参数控制输出质量
# 添加日志记录系统
```
