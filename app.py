import os
import re
import json
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

VERSION = "1.5.0"

# 加载匹配数据 - 从JSON读取
DIRECTORY_FIELDS = []
G_INDEX = {}

def load_match_data():
    global DIRECTORY_FIELDS, G_INDEX
    
    local_file = os.path.join(os.path.dirname(__file__), 'templates', '工商库.xlsx')
    json_file = os.path.join(os.path.dirname(__file__), 'templates', 'g_index.json')
    
    # 加载目录
    if os.path.exists(local_file):
        try:
            df = pd.read_excel(local_file, sheet_name='目录')
            if '对应数据名称' in df.columns:
                DIRECTORY_FIELDS = df['对应数据名称'].dropna().astype(str).tolist()
                DIRECTORY_FIELDS = [x.strip() for x in DIRECTORY_FIELDS if x.strip()]
                print(f"目录: {len(DIRECTORY_FIELDS)} 条")
        except Exception as e:
            print(f"加载目录失败: {e}")
    
    # 加载G列索引（自动刷新）
    try:
        if os.path.exists(json_file):
            # 检查文件是否过期（超过7天）
            import time
            json_mtime = os.path.getmtime(json_file)
            json_age_days = (time.time() - json_mtime) / 86400
            
            if json_age_days > 7:
                print(f"JSON索引已过期({json_age_days:.1f}天)，重新生成...")
                os.remove(json_file)
            else:
                with open(json_file, 'r', encoding='utf-8') as f:
                    G_INDEX = json.load(f)
                print(f"G列索引: {len(G_INDEX)} 个sheet")
        else:
            # 重新生成JSON索引
            print("JSON索引不存在，正在生成...")
        
        # 如果JSON不存在，重新生成
        if not G_INDEX:
            xl = pd.ExcelFile(local_file)
            for sheet in xl.sheet_names:
                if sheet in ['目录', 'Sheet1']:
                    continue
                try:
                    df = pd.read_excel(local_file, sheet_name=sheet)
                    if len(df.columns) >= 7:
                        g_col = df.columns[6]
                        g_data = df[g_col].dropna().astype(str).tolist()
                        g_data = [x.strip() for x in g_data if x.strip() and len(x) > 1]
                        if g_data:
                            G_INDEX[sheet] = g_data
                except:
                    pass
            
            # 保存到JSON
            with open(json_file, 'w', encoding='utf-8') as f:
                json.dump(G_INDEX, f, ensure_ascii=False, indent=2)
            print(f"G列索引已生成并保存: {len(G_INDEX)} 个sheet")
    
    except Exception as e:
        print(f"加载G列索引失败: {e}")

load_match_data()

# 语义相关词映射（用户词 -> 相关词）
SEMANTIC_MAP = {
    '评级': ['等级', '评级', '信用'],
    '等级': ['等级', '评级', '信用'],
    '税务': ['纳税', '税务', '税'],
    '纳税': ['纳税', '税务', '税'],
    '联系': ['联系', '电话', '邮箱', '地址'],
    '电话': ['联系', '电话', '号码'],
    '邮箱': ['联系', '邮箱', '邮件'],
    '地址': ['联系', '地址', '所在地'],
    '处罚': ['处罚', '违法', '惩罚'],
    '许可': ['许可', '批准', '资质'],
    '变更': ['变更', '变动', '修改'],
    '投资': ['投资', '出资', '股权'],
    '股东': ['股东', '出资', '投资'],
    '法人': ['法人', '代表', '负责人'],
    '年报': ['年报', '年度报告', '年度'],
    '基本': ['基本', '基础', '主要'],
}

def clean_text(s):
    return s.replace('工商-', '').replace('企业', '').replace('公司', '').replace('信息', '').replace('数据', '').replace('记录', '').replace(' ', '').strip()

def get_semantic_score(user_field, target_field):
    """计算语义相关性分数"""
    score = 0
    
    # 检查是否有语义相关的词
    for user_word, related_words in SEMANTIC_MAP.items():
        if user_word in user_field:
            for related in related_words:
                if related in target_field:
                    score += 20  # 每个语义相关加20分
    
    return score

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
        
        # 基础相似度
        base_score = 0
        match_type = ''
        
        if user_field == target or user_clean == target_clean:
            base_score = 100
            match_type = '完全匹配'
        else:
            try:
                sim = Levenshtein.ratio(user_clean, target_clean)
                if sim >= 0.4:
                    base_score = int(sim * 100)
                    match_type = '推荐'
            except:
                pass
        
        # 加上语义分数
        if base_score > 0:
            semantic_bonus = get_semantic_score(user_field, target)
            total_score = min(100, base_score + semantic_bonus)  # 最高100分
            
            if best_match is None or total_score > best_match['score']:
                best_match = {'matched': target, 'source': '目录', 'type': match_type, 'score': total_score}
    
    # 2. 从G列索引中找
    if best_match is None or best_match['score'] < 100:
        for sheet_name, g_data in G_INDEX.items():
            for target in g_data:
                target = str(target).strip()
                if not target:
                    continue
                target_clean = clean_text(target)
                
                base_score = 0
                match_type = ''
                
                if user_field == target or user_clean == target_clean:
                    base_score = 100
                    match_type = '完全匹配'
                else:
                    try:
                        sim = Levenshtein.ratio(user_clean, target_clean)
                        if sim >= 0.4:
                            base_score = int(sim * 100)
                            match_type = '推荐'
                    except:
                        pass
                
                if base_score > 0:
                    semantic_bonus = get_semantic_score(user_field, target)
                    total_score = min(100, base_score + semantic_bonus)
                    
                    if best_match is None or total_score > best_match['score']:
                        best_match = {'matched': target, 'source': sheet_name, 'type': match_type, 'score': total_score}
    
    return best_match

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
    
    # 2. 如果目录没匹配，从G列索引中找
    if best_match is None:
        # 从JSON索引中查找
        for sheet_name, g_data in G_INDEX.items():
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

from flask import make_response

@app.route('/template/txt')
def download_template():
    template_file = os.path.join(os.path.dirname(__file__), 'templates', '模板.txt')
    return send_file(template_file, as_attachment=True, mimetype='text/plain')

@app.route('/match', methods=['POST'])
def match_fields():
    try:
        # 支持单个字段查询（通过URL参数）
        single_field = request.form.get('single_field') or request.args.get('single_field')
        
        if single_field:
            # 单个字段查询
            user_fields = [single_field.strip()]
        elif 'file' not in request.files:
            return jsonify({'error': '请上传文件或输入字段'}), 400
        else:
            # 文件上传
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