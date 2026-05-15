import os
import io
import json
import re
import tempfile
from flask import Flask, request, jsonify, render_template_string, Response
import openai
import PyPDF2
import docx
import openpyxl
from threading import Thread
from queue import Queue
import time

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024

# 默认 API 配置（可被前端覆盖）
DEFAULT_API_KEY = os.getenv("AI_API_KEY", "")
DEFAULT_API_BASE = os.getenv("AI_API_BASE", "https://open.bigmodel.cn/api/paas/v4/")
DEFAULT_MODEL = os.getenv("AI_MODEL", "glm-4")

# ------------------- 前端页面 -------------------
HTML_PAGE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>星际猎聘 · AI 智能匹配</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0-alpha1/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { background: radial-gradient(ellipse at top, #0a0f2a 0%, #030614 100%); color: #eef5ff; font-family: 'Segoe UI', system-ui; overflow-x: hidden; }
        body::before { content: ''; position: fixed; top:0; left:0; width:100%; height:100%; background-image: radial-gradient(2px 2px at 20px 30px, #fff, transparent), radial-gradient(1px 1px at 80px 110px, #f8f9fc, transparent); background-size: 400px 400px, 350px 350px; opacity:0.4; pointer-events:none; }
        .glass-card { background: rgba(15, 25, 45, 0.65); backdrop-filter: blur(12px); border: 1px solid rgba(72, 187, 255, 0.3); border-radius: 28px; }
        .card-header-custom { background: linear-gradient(135deg, rgba(0,30,60,0.9), rgba(0,80,120,0.7)); border-bottom: 1px solid #3b82f6; border-radius: 28px 28px 0 0; color: #b3e4ff; }
        .btn-neon { background: linear-gradient(95deg, #0a2f5a, #0a4c7a); border: none; box-shadow: 0 0 8px #00a6ff; font-weight: 600; color: white; }
        .btn-neon:hover { transform: scale(1.02); box-shadow: 0 0 18px #3b82f6; }
        .upload-area { background: rgba(10,20,35,0.6); border: 1.5px dashed #2f6ea2; border-radius: 1.5rem; cursor: pointer; transition: 0.2s; text-align: center; padding: 1rem; }
        .upload-area:hover { border-color: #3b82f6; background: rgba(30,80,120,0.4); }
        .file-list { background: rgba(0,0,0,0.4); border-radius: 16px; border: 1px solid #2c5a7a; max-height: 180px; overflow-y: auto; padding: 8px; }
        .result-table th { background: #0f2a44; color: #aaf0ff; }
        .result-table td { color: #e2f0ff; }
        .badge-score { background: linear-gradient(145deg, #00b4ff, #0066cc); font-size: 1rem; padding: 6px 14px; border-radius: 40px; font-weight: bold; }
        input, select { background: #0a1428 !important; border: 1px solid #2c6e9e !important; color: #eef5ff !important; }
        input:focus { border-color: #00aaff !important; box-shadow: 0 0 8px #00aaff !important; }
        .spinner-overlay { position: fixed; top:0; left:0; width:100%; height:100%; background: rgba(0,0,0,0.8); backdrop-filter: blur(6px); z-index: 9999; display: flex; align-items: center; justify-content: center; }
    </style>
</head>
<body>
<nav class="navbar navbar-dark sticky-top" style="background: rgba(3,10,25,0.9); backdrop-filter: blur(10px); border-bottom:1px solid #2d5f8b;">
    <div class="container"><span class="navbar-brand"><i class="fas fa-galaxy me-2"></i> 星际猎聘 · AI 智能匹配</span></div>
</nav>
<div class="container py-4" style="position:relative; z-index:2;">
    <div class="row mb-4">
        <div class="col-12">
            <div class="glass-card p-3 p-md-4">
                <div class="row g-3">
                    <div class="col-md-4"><label>API Key (留空则用服务器预设)</label><input type="password" id="apiKey" class="form-control" placeholder="可选填你的 Key"></div>
                    <div class="col-md-3"><label>API Base URL</label><input type="text" id="apiBase" class="form-control" value="https://open.bigmodel.cn/api/paas/v4/"></div>
                    <div class="col-md-3"><label>模型名称</label><input type="text" id="modelName" class="form-control" value="glm-4"></div>
                    <div class="col-md-2"><label>匹配阈值</label><input type="number" id="threshold" class="form-control" value="80" min="0" max="100"></div>
                </div>
                <small class="text-muted mt-2 d-block">支持智谱、DeepSeek、OpenAI 等兼容接口。Key 仅本次使用，不记录。</small>
            </div>
        </div>
    </div>

    <div class="row g-4">
        <div class="col-md-6">
            <div class="glass-card h-100">
                <div class="card-header-custom p-3"><i class="fas fa-user-astronaut me-2"></i> 上传人才简历</div>
                <div class="card-body p-3">
                    <div class="upload-area" id="resumeDrop"><i class="fas fa-cloud-upload-alt fa-2x mb-2"></i><p>点击或拖拽简历文件</p><small>PDF / Word / Excel / TXT (每文件一人)</small><input type="file" id="resumeFiles" multiple class="d-none" accept=".pdf,.docx,.xlsx,.xls,.txt,.doc"></div>
                    <div class="mt-3"><strong>已选简历 (<span id="resumeCount">0</span>)</strong><div id="resumeFileList" class="file-list mt-2"></div></div>
                </div>
            </div>
        </div>
        <div class="col-md-6">
            <div class="glass-card h-100">
                <div class="card-header-custom p-3"><i class="fas fa-building me-2"></i> 上传岗位需求</div>
                <div class="card-body p-3">
                    <div class="upload-area" id="jobDrop"><i class="fas fa-chart-network fa-2x mb-2"></i><p>点击或拖拽岗位文件</p><small>每文件一个岗位 (含企业+职责)</small><input type="file" id="jobFiles" multiple class="d-none" accept=".pdf,.docx,.xlsx,.xls,.txt,.doc"></div>
                    <div class="mt-3"><strong>已选岗位 (<span id="jobCount">0</span>)</strong><div id="jobFileList" class="file-list mt-2"></div></div>
                </div>
            </div>
        </div>
    </div>

    <div class="row my-4">
        <div class="col text-center">
            <button id="matchBtn" class="btn btn-neon btn-lg px-5 py-3"><i class="fas fa-satellite me-2"></i> 启动匹配引擎</button>
            <div id="progressMsg" class="mt-3 small"></div>
        </div>
    </div>

    <div class="row">
        <div class="col-12">
            <div class="glass-card p-0 overflow-hidden">
                <div class="card-header-custom p-3 d-flex justify-content-between"><h5 class="mb-0"><i class="fas fa-trophy text-warning me-2"></i> 匹配结果</h5><span class="badge bg-secondary" id="resultCountBadge">0</span></div>
                <div class="table-responsive result-table">
                    <table class="table table-hover align-middle mb-0">
                        <thead><tr><th>人才</th><th>匹配岗位</th><th>AI 分析理由</th><th>分数</th></tr></thead>
                        <tbody id="resultBody"><tr><td colspan="4" class="text-center py-5">✨ 上传文件并点击匹配 ✨</td></tr></tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
</div>
<div id="loadingOverlay" style="display:none;" class="spinner-overlay"><div class="glass-card p-4 text-center"><div class="spinner-border text-info" style="width:3rem;height:3rem;"></div><p class="mt-3">AI 计算中 ...</p><small id="loadingDetail"></small></div></div>
<script>
    let resumeFiles = [], jobFiles = [];
    const elements = {
        resumeInput: document.getElementById('resumeFiles'),
        jobInput: document.getElementById('jobFiles'),
        resumeCount: document.getElementById('resumeCount'),
        jobCount: document.getElementById('jobCount'),
        resumeList: document.getElementById('resumeFileList'),
        jobList: document.getElementById('jobFileList'),
        matchBtn: document.getElementById('matchBtn'),
        progress: document.getElementById('progressMsg'),
        overlay: document.getElementById('loadingOverlay'),
        detail: document.getElementById('loadingDetail'),
        resultBody: document.getElementById('resultBody'),
        resultBadge: document.getElementById('resultCountBadge')
    };

    function updateUI(files, countElem, listElem, icon) {
        countElem.textContent = files.length;
        if (!files.length) {
            listElem.innerHTML = '<div class="text-muted small">暂无文件</div>';
            return;
        }
        let html = '<ul class="list-unstyled mb-0">';
        files.forEach(f => {
            html += `<li><i class="fas ${icon}"></i> ${escapeHtml(f.name)} (${(f.size/1024).toFixed(1)}KB)</li>`;
        });
        html += '</ul>';
        listElem.innerHTML = html;
    }

    function escapeHtml(s) { return String(s).replace(/[&<>]/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[m])); }

    elements.resumeInput.addEventListener('change', e => {
        resumeFiles = Array.from(e.target.files);
        updateUI(resumeFiles, elements.resumeCount, elements.resumeList, 'fa-file-alt');
    });
    elements.jobInput.addEventListener('change', e => {
        jobFiles = Array.from(e.target.files);
        updateUI(jobFiles, elements.jobCount, elements.jobList, 'fa-briefcase');
    });

    function setupDrag(areaId, inputId, isResume) {
        const area = document.getElementById(areaId);
        const input = document.getElementById(inputId);
        area.addEventListener('dragover', e => { e.preventDefault(); area.style.borderColor='#3b82f6'; });
        area.addEventListener('dragleave', () => area.style.borderColor='#2f6ea2');
        area.addEventListener('drop', e => {
            e.preventDefault();
            area.style.borderColor='#2f6ea2';
            const files = Array.from(e.dataTransfer.files);
            if (isResume) { resumeFiles = files; updateUI(resumeFiles, elements.resumeCount, elements.resumeList, 'fa-file-alt'); }
            else { jobFiles = files; updateUI(jobFiles, elements.jobCount, elements.jobList, 'fa-briefcase'); }
        });
        area.addEventListener('click', () => input.click());
    }
    setupDrag('resumeDrop', 'resumeFiles', true);
    setupDrag('jobDrop', 'jobFiles', false);

    elements.matchBtn.onclick = async () => {
        if (!resumeFiles.length || !jobFiles.length) {
            alert('请同时上传简历和岗位文件');
            return;
        }
        const apiKey = document.getElementById('apiKey').value.trim();
        const apiBase = document.getElementById('apiBase').value.trim();
        const modelName = document.getElementById('modelName').value.trim();
        const threshold = parseInt(document.getElementById('threshold').value) || 80;

        const formData = new FormData();
        formData.append('api_key', apiKey);
        formData.append('api_base', apiBase);
        formData.append('model', modelName);
        formData.append('threshold', threshold);
        resumeFiles.forEach(f => formData.append('resumes', f));
        jobFiles.forEach(f => formData.append('jobs', f));

        elements.overlay.style.display = 'flex';
        elements.progress.innerHTML = '<i class="fas fa-spinner fa-pulse"></i> 上传并解析文件...';
        elements.detail.textContent = '';

        try {
            const resp = await fetch('/api/match', { method: 'POST', body: formData });
            const data = await resp.json();
            if (!resp.ok) throw new Error(data.error || '匹配失败');
            renderResults(data.matches, threshold);
            elements.progress.innerHTML = `<i class="fas fa-check-circle"></i> 完成！共 ${data.matches.length} 条高分匹配 (≥${threshold})`;
        } catch (err) {
            alert('匹配失败: ' + err.message);
            elements.progress.innerHTML = `<span class="text-danger">错误: ${err.message}</span>`;
        } finally {
            elements.overlay.style.display = 'none';
        }
    };

    function renderResults(matches, th) {
        elements.resultBody.innerHTML = '';
        elements.resultBadge.textContent = matches.length;
        if (!matches.length) {
            elements.resultBody.innerHTML = `<tr><td colspan="4" class="text-center py-4">⚠️ 无分数 ≥${th} 的匹配</td></tr>`;
            return;
        }
        matches.forEach(m => {
            elements.resultBody.innerHTML += `
                <tr>
                    <td><strong>${escapeHtml(m.candidate)}</strong></td>
                    <td><span class="badge bg-primary bg-opacity-75">${escapeHtml(m.job)}</span></td>
                    <td><small>${escapeHtml(m.reason)}</small></td>
                    <td><span class="badge-score">${m.score}分</span></td>
                </tr>`;
        });
    }
</script>
</body>
</html>"""

# ------------------- 后端工具函数 -------------------
def extract_text_from_file(file_storage):
    """从上传的文件中提取文本"""
    filename = file_storage.filename.lower()
    try:
        if filename.endswith('.pdf'):
            reader = PyPDF2.PdfReader(file_storage)
            text = []
            for page in reader.pages:
                page_text = page.extract_text()
                if page_text:
                    text.append(page_text)
            return ' '.join(text)[:5000]  # 限制长度
        elif filename.endswith('.docx'):
            doc = docx.Document(file_storage)
            return ' '.join([para.text for para in doc.paragraphs])[:5000]
        elif filename.endswith(('.xlsx', '.xls')):
            wb = openpyxl.load_workbook(file_storage, read_only=True)
            rows = []
            for sheet in wb:
                for row in sheet.iter_rows(values_only=True):
                    row_text = ' '.join([str(cell) if cell else '' for cell in row])
                    rows.append(row_text)
            return ' '.join(rows)[:5000]
        elif filename.endswith('.txt'):
            return file_storage.read().decode('utf-8')[:5000]
        else:
            return ""
    except Exception as e:
        app.logger.error(f"文件解析失败: {filename}, 错误: {e}")
        return ""

def call_ai(prompt, api_key, api_base, model, temperature=0.3, json_mode=True):
    """统一 AI 调用"""
    client = openai.OpenAI(api_key=api_key, base_url=api_base)
    kwargs = {
        "model": model,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": temperature,
    }
    if json_mode:
        kwargs["response_format"] = {"type": "json_object"}
    response = client.chat.completions.create(**kwargs)
    return response.choices[0].message.content

def extract_candidates(texts_batch, api_key, api_base, model):
    """批量提取候选人信息"""
    items = []
    for idx, text in enumerate(texts_batch):
        items.append({"id": idx, "text": text})
    prompt = "你是一个专业HR。请从以下简历文本中提取：姓名、唯一编号(若无则生成R开头编号)、技术专长(核心技能)。\n"
    for item in items:
        prompt += f"\n简历 {item['id']+1}: {item['text'][:2000]}\n"
    prompt += "\n输出一个JSON数组，每个元素为 {\"name\":\"\", \"id\":\"\", \"expertise\":\"\"}，保持顺序。"
    try:
        res = call_ai(prompt, api_key, api_base, model, temperature=0.2)
        data = json.loads(res)
        return data
    except Exception as e:
        app.logger.error(f"批量提取候选人失败: {e}")
        # 降级：逐条提取
        results = []
        for item in items:
            try:
                p = f"提取姓名、编号、专长。文本：{item['text'][:1500]}。输出JSON：{{\"name\":\"\",\"id\":\"\",\"expertise\":\"\"}}"
                r = call_ai(p, api_key, api_base, model, temperature=0.1)
                results.append(json.loads(r))
            except:
                results.append({"name": f"候选人{item['id']+1}", "id": f"R{item['id']+1}", "expertise": item['text'][:200]})
        return results

def extract_jobs(texts_batch, api_key, api_base, model):
    """批量提取岗位信息"""
    items = []
    for idx, text in enumerate(texts_batch):
        items.append({"id": idx, "text": text})
    prompt = "从以下岗位描述中提取：企业/公司名称、岗位名称、核心技能要求。\n"
    for item in items:
        prompt += f"\n岗位 {item['id']+1}: {item['text'][:2000]}\n"
    prompt += "\n输出JSON数组，每个元素 {\"company\":\"\", \"title\":\"\", \"requirements\":\"\"}，保持顺序。"
    try:
        res = call_ai(prompt, api_key, api_base, model, temperature=0.2)
        return json.loads(res)
    except:
        results = []
        for item in items:
            try:
                p = f"提取公司、岗位、要求。文本：{item['text'][:1500]}。输出JSON：{{\"company\":\"\",\"title\":\"\",\"requirements\":\"\"}}"
                r = call_ai(p, api_key, api_base, model, temperature=0.1)
                results.append(json.loads(r))
            except:
                results.append({"company": "未知", "title": f"岗位{item['id']+1}", "requirements": item['text'][:200]})
        return results

def score_pairs(candidates, jobs, api_key, api_base, model):
    """批量评分：将所有对组合成一个提示，减少调用次数"""
    pairs = []
    for c_idx, c in enumerate(candidates):
        for j_idx, j in enumerate(jobs):
            pairs.append((c_idx, j_idx))
    batch_size = 5  # 每次请求最多5对，避免提示过长
    all_scores = []
    for i in range(0, len(pairs), batch_size):
        batch = pairs[i:i+batch_size]
        prompt = "作为技术招聘专家，请评估以下人才与岗位的技术匹配度(0-100分)并给简短理由。\n"
        for k, (ci, ji) in enumerate(batch):
            prompt += f"\n对{k+1}: 人才专长: {candidates[ci]['expertise']}; 岗位要求: {jobs[ji]['requirements']}\n"
        prompt += "\n输出JSON数组，每个元素 {\"score\":整数, \"reason\":\"简短理由\"}，严格按顺序。"
        try:
            res = call_ai(prompt, api_key, api_base, model, temperature=0.1)
            scores_batch = json.loads(res)
            for item in scores_batch:
                all_scores.append(item)
        except:
            # 降级为单对评分
            for (ci, ji) in batch:
                try:
                    p = f"专长: {candidates[ci]['expertise']}\n要求: {jobs[ji]['requirements']}\n输出JSON:{{\"score\":整数,\"reason\":\"\"}}"
                    r = call_ai(p, api_key, api_base, model, temperature=0.1)
                    all_scores.append(json.loads(r))
                except:
                    all_scores.append({"score": 0, "reason": "AI评分失败"})
    return all_scores

@app.route('/api/match', methods=['POST'])
def match():
    try:
        # 读取参数
        api_key = request.form.get('api_key', '').strip()
        api_base = request.form.get('api_base', '').strip() or DEFAULT_API_BASE
        model = request.form.get('model', '').strip() or DEFAULT_MODEL
        threshold = int(request.form.get('threshold', 80))
        if not api_key and DEFAULT_API_KEY:
            api_key = DEFAULT_API_KEY
        if not api_key:
            return jsonify({"error": "请填写 API Key 或服务器预设"}), 400

        # 解析文件
        resume_files = request.files.getlist('resumes')
        job_files = request.files.getlist('jobs')
        if not resume_files or not job_files:
            return jsonify({"error": "请上传文件"}), 400

        resume_texts = []
        for f in resume_files:
            text = extract_text_from_file(f)
            if text.strip():
                resume_texts.append(text)
        job_texts = []
        for f in job_files:
            text = extract_text_from_file(f)
            if text.strip():
                job_texts.append(text)
        if not resume_texts or not job_texts:
            return jsonify({"error": "未提取到有效文本"}), 400

        # 批量提取信息
        candidates = extract_candidates(resume_texts, api_key, api_base, model)
        jobs = extract_jobs(job_texts, api_key, api_base, model)

        # 批量评分
        scores = score_pairs(candidates, jobs, api_key, api_base, model)

        # 组装结果
        matches = []
        pair_idx = 0
        for c_idx, c in enumerate(candidates):
            for j_idx, j in enumerate(jobs):
                score_data = scores[pair_idx]
                pair_idx += 1
                final_score = int(float(score_data.get('score', 0)))
                if final_score >= threshold:
                    matches.append({
                        "candidate": f"{c.get('name', '未知')} (#{c.get('id', f'R{c_idx+1}')})",
                        "job": f"{j.get('company', '未知')} · {j.get('title', '未知')}",
                        "reason": score_data.get('reason', ''),
                        "score": final_score
                    })
        matches.sort(key=lambda x: x['score'], reverse=True)
        return jsonify({"matches": matches})
    except Exception as e:
        app.logger.error(f"匹配过程异常: {e}")
        return jsonify({"error": str(e)}), 500

@app.route('/')
def index():
    return render_template_string(HTML_PAGE)

if __name__ == '__main__':
    app.run(debug=False, host='0.0.0.0', port=5000)