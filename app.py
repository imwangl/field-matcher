import os
import re
import pandas as pd
from flask import Flask, render_template, request, send_file, jsonify
import requests
import Levenshtein

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

VERSION = "1.0.0"
SHEET_ID = "1V6uygE_6POZjS8kHuvtGpWxn5LdpUdwxRg9g87RLWuE"

# 从Google Sheet获取数据
def get_target_fields():
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet=%E7%9B%AE%E5%BD%95"
    response = requests.get(url, timeout=30)
    
    if response.status_code == 200:
        # 用pandas解析CSV
        from io import StringIO
        df = pd.read_csv(StringIO(response.text))
        # 返回"对应数据名称"列
        if '对应数据名称' in df.columns:
            return df['对应数据名称'].dropna().tolist()
    return []

def clean_text(s):
    """清理文本进行匹配"""
    return s.replace('工商-', '').replace('企业', '').replace('公司', '').replace('信息', '').replace('数据', '').replace('记录', '').replace(' ', '').strip()

def parse_txt_fields(filepath):
    """解析TXT文件中的字段"""
    fields = []
    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            line = line.rstrip('；').rstrip(';')
            if '；' in line or ';' in line:
                line = line.replace(';', '、').replace('；', '、')
            
            # 结构化格式 "1、公司概况：基本信息、联系方式"
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

def find_match(user_field, target_fields):
    """匹配字段"""
    user_field = str(user_field).strip()
    if not user_field:
        return None
    
    user_clean = clean_text(user_field)
    
    for target in target_fields:
        target = str(target).strip()
        if not target:
            continue
        
        target_clean = clean_text(target)
        
        # 精确匹配
        if user_field == target or user_clean == target_clean:
            return {'matched': target, 'source': '目录', 'type': '完全匹配', 'score': 100}
        
        # 相似度匹配
        try:
            sim = Levenshtein.ratio(user_clean, target_clean)
            if sim >= 0.4:
                return {'matched': target, 'source': '目录', 'type': '推荐', 'score': int(sim*100)}
        except:
            pass
    
    return None

# 启动时获取一次数据
TARGET_DATA = get_target_fields()
print(f"获取到 {len(TARGET_DATA)} 条目标数据")

@app.route('/')
def index():
    return render_template('index.html', version=VERSION)

@app.route('/match', methods=['POST'])
def match_fields():
    try:
        if 'file' not in request.files:
            return jsonify({'error': '请上传文件'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': '请选择文件'}), 400
        
        # 只接受txt
        ext = os.path.splitext(file.filename)[1].lower()
        if ext != '.txt':
            return jsonify({'error': '只支持TXT文件'}), 400
        
        # 保存文件
        filepath = '/tmp/upload.txt'
        file.save(filepath)
        
        # 解析字段
        user_fields = parse_txt_fields(filepath)
        if not user_fields:
            return jsonify({'error': '未能解析出字段'}), 400
        
        # 匹配
        results = []
        for field in user_fields:
            result = find_match(field, TARGET_DATA)
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
        
        # 统计
        total = len(results)
        exact = len([r for r in results if r['match_type'] == '完全匹配'])
        recommend = len([r for r in results if r['match_type'] == '推荐'])
        failed = len([r for r in results if r['match_type'] == '匹配失败'])
        
        # 保存结果
        result_df = pd.DataFrame(results)
        result_df.to_excel('/tmp/result.xlsx', index=False)
        
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
    return send_file('/tmp/result.xlsx', as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)