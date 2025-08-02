import os
import win32com.client
import logging
import re

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def read_doc_text(word, file_path):
    """读取Word文档内容"""
    try:
        doc = word.Documents.Open(file_path, ReadOnly=True)
        text = doc.Content.Text
        doc.Close()
        return text
    except Exception as e:
        logging.error(f"读取文档失败 {file_path}: {e}")
        return None

def extract_vulnerability_numbers(text, filename):
    """正确提取漏洞数据 - 处理分散在多行的数据"""
    lines = text.splitlines()
    keyword = "漏洞总计（次）"
    occurrences = []
    
    # 找到所有包含关键词的行
    for i, line in enumerate(lines):
        if keyword in line:
            occurrences.append(i)
    
    print(f"\n📄 文件: {filename}")
    print(f"找到 {len(occurrences)} 次 '{keyword}' 出现")
    
    if len(occurrences) >= 2:
        # 使用第二次出现
        second_occurrence_line = occurrences[1]
        print(f"✅ 使用第二次出现 (行号: {second_occurrence_line + 1})")
        
        # 从第二次出现的位置向后查找数据
        # 根据调试结果，数据在后续的几行中，每行一个数字
        extracted_numbers = []
        
        # 向后查找最多10行，寻找数字数据
        for i in range(second_occurrence_line + 1, min(second_occurrence_line + 11, len(lines))):
            line = lines[i].strip()
            
            # 移除 \x07 分隔符，获取纯净内容
            clean_line = line.replace('\x07', '').strip()
            
            # 如果是纯数字，添加到结果中
            if clean_line.isdigit():
                extracted_numbers.append(clean_line)
                print(f"  行{i+1}: 提取数字 '{clean_line}'")
            # 如果行为空或只有分隔符，跳过
            elif not clean_line or clean_line == '':
                continue
            # 如果遇到非数字内容，可能是表格结束了
            else:
                # 但是先检查是否包含数字
                numbers_in_line = re.findall(r'\d+', clean_line)
                if numbers_in_line:
                    extracted_numbers.extend(numbers_in_line)
                    print(f"  行{i+1}: 从'{clean_line}'中提取数字 {numbers_in_line}")
                else:
                    break
        
        print(f"📊 提取到的数字: {extracted_numbers}")
        
        # 确保有5个数字（高风险、中风险、低风险、信息、总计）
        while len(extracted_numbers) < 5:
            extracted_numbers.append('0')
        
        # 只取前5个数字
        final_numbers = extracted_numbers[:5]
        print(f"🎯 最终数据: {final_numbers}")
        
        return final_numbers
    
    elif len(occurrences) == 1:
        print(f"⚠️  只找到1次出现，尝试使用第一次")
        first_occurrence_line = occurrences[0]
        
        extracted_numbers = []
        for i in range(first_occurrence_line + 1, min(first_occurrence_line + 11, len(lines))):
            line = lines[i].strip()
            clean_line = line.replace('\x07', '').strip()
            
            if clean_line.isdigit():
                extracted_numbers.append(clean_line)
                print(f"  行{i+1}: 提取数字 '{clean_line}'")
            elif not clean_line:
                continue
            else:
                numbers_in_line = re.findall(r'\d+', clean_line)
                if numbers_in_line:
                    extracted_numbers.extend(numbers_in_line)
                else:
                    break
        
        while len(extracted_numbers) < 5:
            extracted_numbers.append('0')
        
        return extracted_numbers[:5]
    
    else:
        print(f"❌ 未找到关键词")
        return ['0', '0', '0', '0', '0']

def main():
    folder = r'E:\desk\第一批漏洞扫描结果'
    
    if not os.path.exists(folder):
        print(f"错误: 文件夹不存在 - {folder}")
        return
    
    # 初始化Word应用
    try:
        word = win32com.client.Dispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False
        print("Word应用启动成功")
    except Exception as e:
        print(f"无法启动Word应用: {e}")
        return
    
    results = []
    
    try:
        # 获取所有Word文档
        doc_files = [f for f in os.listdir(folder) 
                    if f.lower().endswith(('.doc', '.docx')) and not f.startswith('~$')]
        
        print(f"找到 {len(doc_files)} 个文档")
        
        for filename in doc_files:
            file_path = os.path.join(folder, filename)
            logging.info(f"正在处理: {filename}")
            
            try:
                text = read_doc_text(word, file_path)
                if text is None:
                    continue
                
                numbers = extract_vulnerability_numbers(text, filename)
                
                results.append({
                    'filename': filename,
                    'data': numbers
                })
                
                print("-" * 60)
                
            except Exception as e:
                logging.error(f'处理文件 {filename} 时出错: {e}')
        
        # 保存结果
        if results:
            output_file = os.path.join(folder, '修正后的漏洞统计结果.txt')
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write("文件名\t高风险\t中风险\t低风险\t信息\t漏洞总计\n")
                
                for result in results:
                    data_line = '\t'.join(result['data'])
                    f.write(f"{result['filename']}\t{data_line}\n")
            
            print(f"\n💾 结果已保存到: {output_file}")
            
            # 显示汇总
            print(f"\n🎯 提取汇总:")
            for result in results:
                print(f"{result['filename']}: {result['data']}")
                
    except Exception as e:
        print(f"处理过程中出错: {e}")
    finally:
        try:
            word.Quit()
            print("Word应用已关闭")
        except:
            pass

if __name__ == '__main__':
    main()