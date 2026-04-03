import os
import re
import pandas as pd
from flask import Flask, render_template, request, send_file, jsonify
import Levenshtein
from io import BytesIO

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

VERSION = "1.2.0"

# 加载匹配数据
DIRECTORY_FIELDS = []   # 目录
SHEET_G_DATA = {}      # G列

def load_directory():
    """加载目录"""
    global DIRECTORY_FIELDS
    
    local_file = os.path.join(os.path.dirname(__file__), 'templates', '工商库.xlsx')
    if not os.path.exists(local_file):
        return
    
    try:
        df = pd.read_excel(local_file, sheet_name='目录')
        if '对应数据名称' in df.columns:
            DIRECTORY_FIELDS = df['对应数据名称'].dropna().astype(str).tolist()
            DIRECTORY_FIELDS = [x.strip() for x in DIRECTORY_FIELDS if x.strip()]
            print(f"目录: {len(DIRECTORY_FIELDS)} 条")
    except Exception as e:
        print(f"加载目录失败: {e}")

def load_g_columns():
    """延迟加载G列"""
    global SHEET_G_DATA
    
    if SHEET_G_DATA:
        return
    
    local_file = os.path.join(os.path.dirname(__file__), 'templates', '工商库.xlsx')
    if not os.path.exists(local_file):
        return
    
    try:
        xl = pd.ExcelFile(local_file)
        
        # 常用sheet（年报相关）
        priority_sheets = ['工商-年报基本信息', '经营信息-地块公示', '特殊主体-事业单位']
        
        for sheet in priority_sheets:
            if sheet in xl.sheet_names:
                try:
                    df = pd.read_excel(local_file, sheet_name=sheet)
                    if len(df.columns) >= 7:
                        g_col = df.columns[6]
                        g_data = df[g_col].dropna().astype(str).tolist()
                        g_data = [x.strip() for x in g_data if x.strip() and len(x) > 1]
                        if g_data:
                            SHEET_G_DATA[sheet] = g_data
                            print(f"{sheet}: {len(g_data)} 条")
                except:
                    pass
        
        print(f"G列加载完成: {len(SHEET_G_DATA)} 个sheet")
    except Exception as e:
        print(f"加载G列失败: {e}")

# 启动时只加载目录
load_directory()

def clean_text(s):
    return s.replace('工商-', '').replace('企业', '').replace('公司', '').replace('信息', '').replace('数据', '').replace('记录', '').replace(' ', '').strip()

def parse_txt_fields(filepath):
    fields = []
    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            line = line.rstrip('；').rstrip(';')
            if '；' in line or ';' in line:
                line = line.replace(';', '、').replace('；', '、')
            
            match = re.match(r'^\d+、[^：]+：(.+)$', line)
            if match:
                content = match.group(1)
                parts = content.split('、')
                fields.extend([p.strip() for p in parts if p.strip()])
            else:
                for sep in ['、', '，', ',']:
                    if sep in line:
                        fields.extend([p.strip() for p in line.split(sep) if p.strip()])
                        break
                else:
                    if line.strip():
                        fields.append(line.strip())
    return fields

def find_match(user_field):
    user_field = str(user_field).strip()
    if not user_field:
        return None
    
    user_clean = clean_text(user_field)
    
    best_match = None
    
    # 1. 匹配目录
    for target in DIRECTORY_FIELDS:
        target = str(target).strip()
        if not target:
            continue
        target_clean = clean_text(target)
        
        score = 0
        match_type = ''
        
        if user_field == target or user_clean == target_clean:
            score = 100
            match_type = '完全匹配'
        else:
            try:
                sim = Levenshtein.ratio(user_clean, target_clean)
                if sim >= 0.4:
                    score = int(sim * 100)
                    match_type = '推荐'
            except:
                pass
        
        if score > 0:
            if best_match is None or score > best_match['score']:
                best_match = {'matched': target, 'source': '目录', 'type': match_type, 'score': score}
    
    # 2. 匹配G列
    load_g_columns()
    
    for sheet_name, g_data in SHEET_G_DATA.items():
        for target in g_data:
            target = str(target).strip()
            if not target:
                continue
            target_clean = clean_text(target)
            
            score = 0
            match_type = ''
            
            if user_field == target or user_clean == target_clean:
                score = 100
                match_type = '完全匹配'
            else:
                try:
                    sim = Levenshtein.ratio(user_clean, target_clean)
                    if sim >= 0.4:
                        score = int(sim * 100)
                        match_type = '推荐'
                except:
                    pass
            
            if score > 0:
                if best_match is None or score > best_match['score']:
                    best_match = {'matched': target, 'source': sheet_name, 'type': match_type, 'score': score}
    
    return best_match

@app.route('/')
def index():
    return render_template('index.html', version=VERSION)

@app.route('/template/txt')
def download_template():
    content = "1、公司概况：基本信息、联系方式、变更记录，主要人员；\n2、股东信息：股东信息、对外投资；"
    output = BytesIO(content.encode('utf-8'))
    return send_file(output, download_name='模板.txt', as_attachment=True)

@app.route('/match', methods=['POST'])
def match_fields():
    try:
        if 'file' not in request.files:
            return jsonify({'error': '请上传文件'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': '请选择文件'}), 400
        
        ext = os.path.splitext(file.filename)[1].lower()
        if ext != '.txt':
            return jsonify({'error': '只支持TXT文件'}), 400
        
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filepath)
        
        user_fields = parse_txt_fields(filepath)
        if not user_fields:
            return jsonify({'error': '未能解析出字段'}), 400
        
        results = []
        for field in user_fields:
            result = find_match(field)
            if result:
                results.append({
                    'user_field': field,
                    'matched': result['matched'],
                    'source': result['source'],
                    'match_type': result['type'],
                    'score': result['score']
                })
            else:
                results.append({
                    'user_field': field,
                    'matched': '-',
                    'source': '-',
                    'match_type': '匹配失败',
                    'score': 0
                })
        
        total = len(results)
        exact = len([r for r in results if r['match_type'] == '完全匹配'])
        recommend = len([r for r in results if r['match_type'] == '推荐'])
        failed = len([r for r in results if r['match_type'] == '匹配失败'])
        
        result_df = pd.DataFrame(results)
        result_df.to_excel(os.path.join(app.config['OUTPUT_FOLDER'], 'result.xlsx'), index=False)
        
        return jsonify({
            'success': True,
            'stats': {'total': total, 'exact': exact, 'recommend': recommend, 'failed': failed},
            'results': results[:100]
        })
    
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/download')
def download_result():
    result_path = os.path.join(app.config['OUTPUT_FOLDER'], 'result.xlsx')
    if os.path.exists(result_path):
        return send_file(result_path, as_attachment=True)
    return "文件未找到", 404

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)