import zipfile
import os

def zip_file(source_path, output_path):
    """将文件压缩成ZIP格式"""
    try:
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # 获取文件名
            filename = os.path.basename(source_path)
            # 将文件添加到ZIP中
            zipf.write(source_path, filename)
        print(f"成功创建压缩文件: {output_path}")
        return True
    except Exception as e:
        print(f"创建压缩文件时出错: {e}")
        return False

# 压缩WordFormatter.exe
source_file = "dist/WordFormatter.exe"
# 设置输出的zip文件名
output_zip = "Wordformatter_V1.0.1.zip"

if zip_file(source_file, output_zip):
    # 显示压缩文件信息
    print(f"\n压缩文件信息:")
    print(f"- 路径: {os.path.abspath(output_zip)}")
    print(f"- 大小: {os.path.getsize(output_zip) / (1024 * 1024):.2f} MB")


