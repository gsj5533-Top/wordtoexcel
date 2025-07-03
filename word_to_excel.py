import os
import shutil
import json
import pandas as pd
from docx import Document
import re

def load_config(config_path):
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)

def load_symbol_config(config_path):
    # 新增：加载符号映射和tick_symbols配置
    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)
    symbol_maps = config.get('symbol_maps', {})
    tick_symbols = config.get('tick_symbols', ['✓', '✔', '☑', '☒', '√'])
    empty_box = config.get('empty_box', '☐')
    return symbol_maps, tick_symbols, empty_box

def extract_text_from_docx(docx_path):
    doc = Document(docx_path)
    return '\n'.join([p.text for p in doc.paragraphs])
#解决word中wingdings123对号无法识别
def extract_text_with_symbols(docx_path, symbol_maps):
    doc = Document(docx_path)
    lines = []
    # 读取配置的映射表
    wingdings_map = symbol_maps.get('wingdings', {})
    wingdings2_map = symbol_maps.get('wingdings2', {})
    wingdings3_map = symbol_maps.get('wingdings3', {})
    for para in doc.paragraphs:
        line = ''
        for run in para.runs:
            text = run.text
            font = run.font.name or ''
            orig_text = text
            # Wingdings 1
            if 'Wingdings' == font:
                for k, v in wingdings_map.items():
                    if k in text:
                        text = text.replace(k, v)
            # Wingdings 2
            elif 'Wingdings 2' == font:
                for k, v in wingdings2_map.items():
                    if k in text:
                        text = text.replace(k, v)
            # Wingdings 3
            elif 'Wingdings 3' == font:
                for k, v in wingdings3_map.items():
                    if k in text:
                        text = text.replace(k, v)
            line += text
        lines.append(line)
    result = '\n'.join(lines)
    return result

def format_date_to_str(text):
    # 匹配如2016.01、2016.10、2016-01、2016/01等，统一转为'2016.01'
    def repl(m):
        y, mth = m.group(1), m.group(2)
        return f"'{y}.{mth.zfill(2)}"  # 前导单引号防止Excel自动转格式
    return re.sub(r'(\d{4})[./-](\d{1,2})', repl, text)

def main():
    config = load_config('config.json')
    src = config['source_folder']
    backup = config['backup_folder']
    excel_path = config['excel_path']
    keywords = config['keywords']

    print("配置内容：", config)
    print("源目录：", src)
    print("目录下文件：", os.listdir(src))

    # 输出文件保存到“结果”文件夹
    result_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '结果')
    os.makedirs(result_dir, exist_ok=True)
    out_filename = os.path.basename(excel_path)
    excel_path = os.path.join(result_dir, out_filename)

    # 读取或新建Excel/CSV
    if excel_path.lower().endswith('.csv'):
        if os.path.exists(excel_path):
            try:
                df = pd.read_csv(excel_path, dtype=str, encoding='utf-8-sig')
            except UnicodeDecodeError:
                df = pd.read_csv(excel_path, dtype=str, encoding='gbk')
        else:
            df = pd.DataFrame({k: [] for k in keywords}, dtype=str)
    else:
        if os.path.exists(excel_path):
            df = pd.read_excel(excel_path, dtype=str, engine='openpyxl')
        else:
            df = pd.DataFrame({k: [] for k in keywords}, dtype=str)

    total_written = 0
    # main函数内加载符号配置
    symbol_maps, tick_symbols, empty_box = load_symbol_config('config.json')
    for fname in os.listdir(src):
        # 跳过Word临时文件
        if fname.lower().endswith('.docx') and not fname.startswith('~$'):
            print(f"处理文件: {fname}")
            fpath = os.path.join(src, fname)
            # text = extract_text_from_docx(fpath)
            text = extract_text_with_symbols(fpath, symbol_maps)
            print("------ Word原始内容 ------")
            print(text)
            print("------ End ------")
            text = format_date_to_str(str(text))  # 处理日期格式
            lines = text.split('\n')
            row = {k: '' for k in keywords}  # 新增：每个文件一行
            for idx, line in enumerate(lines):
                kw_matches = []
                for kw in keywords:
                    kw_pattern = r'{}[ ]*[：:][ ]*'.format(re.escape(kw))
                    for m in re.finditer(kw_pattern, line):
                        kw_matches.append({'kw': kw, 'start': m.start(), 'end': m.end()})
                kw_matches.sort(key=lambda x: x['start'])
                for i, match in enumerate(kw_matches):
                    kw = match['kw']
                    start = match['end']
                    end = kw_matches[i+1]['start'] if i+1 < len(kw_matches) else len(line)
                    val = line[start:end].strip()
                    if any(sym in val for sym in tick_symbols) or empty_box in val:
                        stack = []
                        i2 = 0
                        while i2 < len(val):
                            if val[i2] == empty_box:
                                j = i2 + 1
                                while j < len(val) and val[j] not in tick_symbols + [empty_box]:
                                    j += 1
                                content = val[i2+1:j].strip()
                                if content:
                                    stack.append({'type': 'empty', 'content': content, 'pos': i2})
                                i2 = j
                            elif val[i2] in tick_symbols:
                                j = i2 + 1
                                while j < len(val) and val[j] not in tick_symbols + [empty_box]:
                                    j += 1
                                content = val[i2+1:j].strip()
                                if content:
                                    stack.append({'type': 'tick', 'content': content, 'pos': i2})
                                i2 = j
                            else:
                                i2 += 1
                        tick_indices = [ii for ii, item in enumerate(stack) if item['type'] == 'tick']
                        if tick_indices:
                            first_tick = tick_indices[0]
                            stack = stack[first_tick:]
                        val = ' '.join([item['content'] for item in stack if item['type'] == 'tick']).strip()
                        if val == '' or val.startswith(' '):
                            val = ''
                    row[kw] = str(val)
            # 在每个文件所有关键字处理完后，输出本行所有关键字的最终值
            print(f"[行{idx+1}写入结果] {row}")
            # 检查所有字段，未被赋值的也填空
            for kw in keywords:
                if kw not in row or row[kw] == '':
                    row[kw] = ''
            df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
            total_written += 1
            os.makedirs(backup, exist_ok=True)
            shutil.move(fpath, os.path.join(backup, fname))
    # 全部转为字符串再写入
    df = df.astype(str)
    if excel_path.lower().endswith('.csv'):
        df.to_csv(excel_path, index=False, encoding='utf-8-sig')
    else:
        df.to_excel(excel_path, index=False, engine='openpyxl')
    print(f"\n共写入 {total_written} 条数据到Excel/CSV。")
    print("实际输出路径：", excel_path)

if __name__ == '__main__':
    main()
