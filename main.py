import os
import shutil
import subprocess


def convert_ppt_to_pdf(ppt_file, output_folder):
    """将PPT文件转换为PDF"""
    # 使用LibreOffice的命令行工具进行转换
    command = [
        'soffice', '--headless', '--convert-to', 'pdf',
        '--outdir', output_folder, ppt_file
    ]
    try:
        subprocess.run(command, check=True)
        print(f"已将 {ppt_file} 转换为 PDF.")
    except subprocess.CalledProcessError as e:
        print(f"转换失败: {ppt_file}. 错误信息: {e}")


def extract_files(source_directory, target_directory):
    """从源文件夹中提取所有PDF文件及PPT文件"""
    # 确保目标文件夹存在，如果不存在则创建
    if not os.path.exists(target_directory):
        os.makedirs(target_directory)

    # 遍历源文件夹中的所有子文件夹
    for root, dirs, files in os.walk(source_directory):
        for file in files:
            source_file = os.path.join(root, file)
            # 如果是PDF文件，直接复制
            if file.lower().endswith('.pdf'):
                target_file = os.path.join(target_directory, file)
                if os.path.exists(target_file):
                    base, ext = os.path.splitext(file)
                    counter = 1
                    while os.path.exists(target_file):
                        target_file = os.path.join(target_directory, f"{base}_{counter}{ext}")
                        counter += 1
                shutil.copy(source_file, target_file)
                print(f"复制 {source_file} 到 {target_file}")

            # 如果是PPT文件，转换为PDF后复制
            elif file.lower().endswith('.ppt') or file.lower().endswith('.pptx'):
                # 转换PPT为PDF
                convert_ppt_to_pdf(source_file, target_directory)

    print("所有文件已成功提取并转换到目标文件夹。")


# 示例调用
source_directory = 'D:/reading/金融/经济学原理/课件'  # 源文件夹路径
target_directory = 'D:/reading/金融/经济学原理/集成课件'  # 目标文件夹路径

extract_files(source_directory, target_directory)
