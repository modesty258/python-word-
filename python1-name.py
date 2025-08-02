import os
import re
import win32com.client

def extract_vuln_items(text):
    pattern = re.compile(r'5\.2\.(\d)\s+([^\n\r]+)')
    return [(f'5.2.{num}', title.strip()) for num, title in pattern.findall(text) if 1 <= int(num) <= 8]

def read_doc_text(word, file_path):
    doc = word.Documents.Open(file_path)
    text = doc.Content.Text
    doc.Close()
    return text

def main(folder_path):
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    try:
        for filename in os.listdir(folder_path):
            if filename.lower().endswith('.doc'):
                file_path = os.path.join(folder_path, filename)
                try:
                    text = read_doc_text(word, file_path)
                    items = extract_vuln_items(text)
                    print(f'文件: {filename}')
                    for idx, title in items:
                        print(f'{idx} {title}')
                    print('-' * 40)
                except Exception as e:
                    print(f'处理文件 {filename} 时出错: {e}')
    finally:
        word.Quit()

if __name__ == '__main__':
    folder = r'E:\desk\第一批漏洞扫描结果'
    main(folder)