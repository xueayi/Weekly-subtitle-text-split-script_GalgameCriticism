import os
import re
import docx2txt


# 定义函数以拆分文本
def split_text(text, line_max, line_min):
    sentences = re.split(r'(?<=[。])', text)  # 使用句号、感叹号、问号作为句子的分隔符
    result = []
    current_line = ""

    for sentence in sentences:
        if len(current_line) + len(sentence) <= line_max:
            current_line += sentence
        else:
            result.append(current_line)
            current_line = sentence

    if current_line:
        result.append(current_line)

    return [line.strip() for line in result if len(line) >= line_min]


# 获取文件类型（txt或docx）
def get_file_type(file_path):
    _, ext = os.path.splitext(file_path)
    return ext.lower()


# 处理文本文件
def process_text_file(file_path, line_max, line_min):
    with open(file_path, 'r', encoding='utf-8') as file:
        text = file.read()

        # 删除HTTP和HTTPS网址
        text = remove_urls(text)

        lines = split_text(text, line_max, line_min)
    return lines


# 处理docx文件
def process_docx_file(file_path, line_max, line_min):
    text = docx2txt.process(file_path)

    # 删除HTTP和HTTPS网址
    text = remove_urls(text)

    lines = split_text(text, line_max, line_min)
    return lines


# 删除文本中的HTTP和HTTPS网址
def remove_urls(text):
    # 正则表达式模式，用于匹配HTTP和HTTPS网址
    url_pattern = r'https?://\S+|www\.\S+'

    # 使用sub函数将匹配的网址替换为空字符串
    text_without_urls = re.sub(url_pattern, '', text)

    return text_without_urls


# 输入路径
file_path = input("请输入文件路径或直接将文件拖入框内：").strip('"')  # 去除双引号
line_max = int(input("请输入每行最大字数（周报推荐值115-120）："))
line_min = int(input("请输入每行最小字数（不知道多少可以输入10）："))

file_type = get_file_type(file_path)

if file_type == '.txt':
    lines = process_text_file(file_path, line_max, line_min)
elif file_type == '.docx':
    lines = process_docx_file(file_path, line_max, line_min)
else:
    print("不支持的文件类型")
    exit()  # 不支持的文件类型时，退出程序

# 保存结果到文件
output_file = f"{os.path.splitext(file_path)[0]}_{line_min}_{line_max}_拆分.txt"
with open(output_file, 'w', encoding='utf-8') as file:
    for line in lines:
        file.write(line + '\n')

print(f"拆分完成，结果保存在{output_file}")

# 等待用户确认后退出
input("按 Enter 键退出程序，如有改进建议联系雪阿宜")
