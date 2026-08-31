#!/usr/bin/env python3
"""
标书智能生成器 · 后端版
========================
填HTML表单 → 直接生成Word标书文件（一键下载）

用法：
  python bid_form_server.py
  # 浏览器打开 http://localhost:8080
"""

import os, sys, re, json, io, tempfile
from pathlib import Path

# 确保能找到bid_engine
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from flask import Flask, request, send_file, render_template_string

app = Flask(__name__)


def _desensitize(val, flag):
    """脱敏：flag=True 时返回 ______，否则原样返回"""
    return '______' if flag else val


# ====================== 标书Markdown生成器 ======================

def generate_bid_md(data, desensitize=False):
    """根据表单数据生成标书Markdown

    desensitize=True 时，所有公司名/联系人/电话替换为 ______。
    """
    pn = data.get('project_name', '').strip()
    bn = _desensitize(data.get('bidder_name', '').strip(), desensitize)
    date = data.get('bid_date', '').strip() or __import__('datetime').datetime.now().strftime('%Y-%m-%d')

    md = ''

    # ———— 封面 ————
    cover_title = data.get('cover_title', '').strip() or '投标文件'
    cover_logo_raw = data.get('cover_logo_text', '').strip() or data.get('bidder_name', '').strip()
    cover_logo = _desensitize(cover_logo_raw, desensitize)
    md += f'# {pn}\n\n'
    md += f'## {cover_title}\n\n'
    md += f'**项目名称**：{pn}\n\n'
    if data.get('bid_no'): md += f'**招标编号**：{data["bid_no"]}\n\n'
    md += f'**投标人**：{cover_logo}\n\n'
    md += f'**日期**：{date}\n\n---\n\n'

    # ———— 一、投标函 ————
    bp = data.get('bid_price', '______')
    sp = data.get('service_period', '______')
    lp = _desensitize(data.get('legal_person', '______'), desensitize)
    cp = _desensitize(data.get('contact_person', '______'), desensitize)
    ph = _desensitize(data.get('contact_phone', '______'), desensitize)
    bn_letter = bn or '______'
    validity = data.get('validity_days', '90')

    md += '# 一、投标函\n\n'
    md += f'致：{pn}（招标人）\n\n'
    md += f'我方 {bn_letter}，在仔细阅读了本项目招标文件后，决定参加本项目的投标。\n\n'
    md += f'1. 我方愿意按照招标文件要求，以投标报价 **{bp}万元**，服务期限 **{sp}** 提供本项目所需服务。\n\n'
    md += '2. 我方承诺所提交的投标文件及有关资料内容完整、真实、准确。\n\n'
    md += '3. 我方承诺一旦中标，将严格按照合同约定履行义务，保证服务质量。\n\n'
    md += f'4. 本投标函有效期：自开标之日起{validity}个日历日。\n\n'
    md += f'投标人（盖章）：{bn_letter}\n\n'
    md += f'法定代表人或授权代表（签字）：{lp}\n\n'
    md += f'联系人：{cp}\n\n'
    md += f'联系电话：{ph}\n\n'
    md += f'日期：{date}\n\n---\n\n'

    # ———— 二、授权委托书 ————
    md += '# 二、法定代表人授权委托书\n\n'
    md += f'本授权委托书声明：我 {lp} 系 {bn_letter} 的法定代表人，现授权委托 {cp} 为我单位代理人，以本单位名义参加 {pn} 的投标活动。\n\n'
    md += '代理人在开标、评标、合同谈判过程中所签署的一切文件和处理与之有关的一切事务，我均予以承认。\n\n'
    md += '代理人无转委托权。\n\n特此委托。\n\n'
    md += f'法定代表人（签字）：{lp}\n\n'
    md += f'代理人（签字）：{cp}\n\n'
    md += f'日期：{date}\n\n---\n\n'

    # ———— 三、公司简介 ————
    md += '# 三、公司简介与资质\n\n'
    if data.get('company_desc'):
        cd = _desensitize(data['company_desc'], desensitize) if desensitize else data['company_desc']
        md += cd + '\n\n'
    md += '| 项目 | 内容 |\n|------|------|\n'
    md += f'| 公司全称 | {bn_letter} |\n'
    if data.get('reg_capital'): md += f'| 注册资金 | {data["reg_capital"]}万元 |\n'
    if data.get('found_year'): md += f'| 成立时间 | {data["found_year"]}年 |\n'
    if data.get('staff_count'): md += f'| 员工总数 | {data["staff_count"]}人 |\n'
    md += '\n'

    if data.get('certificates'):
        md += '**资质证书**：\n\n'
        for i, c in enumerate(data['certificates'].split('\n')):
            c = c.strip()
            if c: md += f'{i+1}. {c}\n'
        md += '\n'
    md += '---\n\n'

    # ———— 四、项目类型模板分支 ————
    project_type = data.get('project_type', '其他')

    if project_type == '货物':
        md += '# 四、项目理解与技术要求\n\n'
        md += '## 4.1 技术要求响应\n\n'
        md += (data.get('service_desc') or '（此处填写技术要求响应内容）') + '\n\n'
        md += '## 4.2 货物配置清单\n\n'
        md += (data.get('challenges') or '（此处填写货物配置清单）') + '\n\n'
        md += '## 4.3 质量保证措施\n\n'
        md += (data.get('guarantee') or '（此处填写质量保证措施）') + '\n\n'
        md += '## 4.4 交货与验收方案\n\n'
        md += (data.get('staff_plan') or '（此处填写交货与验收方案）') + '\n\n'
    elif project_type == 'IT':
        md += '# 四、项目理解与技术方案\n\n'
        md += '## 4.1 技术方案概述\n\n'
        md += (data.get('service_desc') or '（此处填写技术方案概述）') + '\n\n'
        md += '## 4.2 系统架构设计\n\n'
        md += (data.get('challenges') or '（此处填写系统架构设计）') + '\n\n'
        md += '## 4.3 实施方案\n\n'
        md += (data.get('guarantee') or '（此处填写实施方案）') + '\n\n'
        md += '## 4.4 运维保障方案\n\n'
        md += (data.get('staff_plan') or '（此处填写运维保障方案）') + '\n\n'
    else:
        # 人力外包 / 物业 / 其他：默认模板
        md += '# 四、项目理解与服务方案\n\n'
        md += '## 4.1 项目理解\n\n'
        md += (data.get('service_desc') or '（此处填写对本项目的理解和总体认识）') + '\n\n'
        md += '## 4.2 项目重难点分析及对策\n\n'
        md += (data.get('challenges') or '（此处分析项目重难点及对策）') + '\n\n'
        md += '## 4.3 服务保障措施\n\n'
        md += (data.get('guarantee') or '（此处填写质量/安全/应急保障措施）') + '\n\n'
        md += '## 4.4 人员配置方案\n\n'
        md += (data.get('staff_plan') or '（此处填写拟投入人员配置方案）') + '\n\n'
    md += '---\n\n'

    # ———— 五、业绩 ————
    md += '# 五、类似项目业绩\n\n'
    if data.get('projects'):
        md += '| 项目名称 | 甲方单位 | 合同金额 | 年份 |\n|----------|----------|----------|------|\n'
        for p in data['projects'].split('\n'):
            p = p.strip()
            if not p: continue
            parts = [s.strip() for s in p.split('|')]
            while len(parts) < 4: parts.append('______')
            if desensitize and len(parts) > 1:
                parts[1] = '______'
            md += '| ' + ' | '.join(parts[:4]) + ' |\n'
    else:
        md += '（此处填写类似项目业绩）\n\n'
    md += '\n---\n\n'

    # ———— 六、报价 ————
    md += '# 六、报价说明\n\n'
    md += '## 6.1 总报价\n\n'
    md += f'本项目投标总报价：**{bp}万元**\n\n'
    if data.get('price_detail'):
        md += '## 6.2 报价明细\n\n| 项目 | 金额（万元） | 说明 |\n|------|-------------|------|\n'
        for p in data['price_detail'].split('\n'):
            p = p.strip()
            if not p: continue
            parts = [s.strip() for s in p.split('|')]
            while len(parts) < 3: parts.append('')
            md += '| ' + ' | '.join(parts[:3]) + ' |\n'
        md += '\n'
    if data.get('discount'):
        md += '## 6.3 优惠承诺\n\n' + data['discount'] + '\n\n'
    md += '---\n\n'

    # ———— 七、服务承诺 ————
    md += '# 七、服务承诺\n\n'
    md += (data.get('commitments') or '（此处填写服务承诺内容') + '\n\n---\n\n'

    # ———— 八、附件 ————
    md += '# 八、附件\n\n'
    md += '1. 营业执照副本复印件\n2. 相关资质证书复印件\n'
    md += '3. 类似项目合同复印件\n4. 其他需要提交的材料\n\n'

    return md


def md_to_docx(md_text, output_path):
    """调用bid_engine将Markdown转为Word"""
    from bid_engine import process_markdown

    # 写临时md文件
    md_path = output_path + '.tmp.md'
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(md_text)

    try:
        # 调用bid_engine
        from bid_engine import BidEngine
        engine = BidEngine(md_path, output_path)
        engine.build()
        return True
    except Exception as e:
        print(f"bid_engine调用失败: {e}")
        return False
    finally:
        if os.path.exists(md_path):
            os.remove(md_path)


# ====================== HTML表单（内置） ======================

HTML_FORM = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>标书智能生成器</title>
<style>
  *{margin:0;padding:0;box-sizing:border-box}
  body{font-family:-apple-system,'Microsoft YaHei',sans-serif;background:#f0f2f5;color:#333;padding:20px}
  .container{max-width:1000px;margin:0 auto}
  .panel{background:#fff;border-radius:12px;padding:24px;box-shadow:0 2px 8px rgba(0,0,0,0.08);margin-bottom:20px}
  h2{font-size:18px;margin-bottom:20px;padding-bottom:12px;border-bottom:2px solid #1a73e8;color:#1a73e8}
  h3{font-size:15px;margin:20px 0 12px;color:#555;border-left:3px solid #1a73e8;padding-left:10px;cursor:pointer;user-select:none}
  h3:first-of-type{margin-top:0}
  .f{margin-bottom:14px}
  .f label{display:block;font-size:13px;font-weight:600;color:#555;margin-bottom:4px}
  .f label .r{color:#e74c3c}
  .f input,.f textarea,.f select{width:100%;padding:8px 12px;border:1px solid #d9d9d9;border-radius:6px;font-size:14px;font-family:inherit}
  .f input:focus,.f textarea:focus{outline:none;border-color:#1a73e8;box-shadow:0 0 0 2px rgba(26,115,232,0.15)}
  .f textarea{resize:vertical;min-height:60px}
  .row{display:grid;grid-template-columns:1fr 1fr;gap:12px}
  .toolbar{display:flex;gap:10px;margin-bottom:16px;flex-wrap:wrap}
  .btn{padding:10px 20px;border:none;border-radius:6px;font-size:14px;font-weight:600;cursor:pointer}
  .btn-primary{background:#1a73e8;color:#fff}
  .btn-primary:hover{background:#1557b0}
  .btn-success{background:#34a853;color:#fff}
  .btn-success:hover{background:#2d8f47}
  .status{background:#e8f5e9;border-radius:6px;padding:12px;font-size:13px;color:#2e7d32;line-height:1.6;display:none}
  .status.error{background:#ffebee;color:#c62828}
  .progress{display:none;text-align:center;padding:20px;color:#666}
  .spinner{display:inline-block;width:20px;height:20px;border:3px solid #e0e0e0;border-top-color:#1a73e8;border-radius:50%;animation:spin .8s linear infinite;margin-right:8px;vertical-align:middle}
  @keyframes spin{to{transform:rotate(360deg)}}
  pre{background:#f8f9fa;border:1px solid #e8e8e8;border-radius:8px;padding:16px;font-size:12px;line-height:1.6;max-height:400px;overflow-y:auto;white-space:pre-wrap;word-break:break-all;margin-top:12px}
  .hint{font-size:12px;color:#888;margin-top:4px}
</style>
</head>
<body>
<div class="container">
  <div class="panel">
    <h2>📋 标书智能生成器</h2>
    <div class="toolbar">
      <button class="btn btn-primary" onclick="generateWord()">🚀 生成Word标书</button>
      <button class="btn btn-success" onclick="previewMd()">👁️ 预览Markdown</button>
    </div>
    <div class="status" id="status"></div>
    <div class="progress" id="progress"><span class="spinner"></span>正在生成标书文件...</div>
    <pre id="preview" style="display:none"></pre>
  </div>

  <div class="panel">
    <h2>📝 标书信息</h2>
    <div class="row">
      <div class="f"><label>项目名称 <span class="r">*</span></label><input id="project_name" placeholder="例：XX医院后勤服务外包项目"></div>
      <div class="f"><label>招标编号</label><input id="bid_no" placeholder="例：0811-DSITC260750"></div>
    </div>
    <div class="row">
      <div class="f"><label>投标人全称 <span class="r">*</span></label><input id="bidder_name" placeholder="例：XX服务管理有限公司"></div>
      <div class="f"><label>投标人简称</label><input id="bidder_short" placeholder="例：XX公司"></div>
    </div>
    <div class="row">
      <div class="f"><label>投标日期</label><input id="bid_date" type="date"></div>
      <div class="f"><label>项目类型</label>
        <select id="project_type">
          <option value="人力外包">人力资源服务外包</option>
          <option value="物业">物业服务/保安保洁</option>
          <option value="货物">货物采购</option>
          <option value="IT">IT技术服务</option>
          <option value="其他">其他</option>
        </select>
        <div class="hint">💡 不同类型自动匹配章节模板：人力/物业→服务方案；货物→技术要求；IT→技术方案</div>
      </div>
    </div>
    <div class="row">
      <div class="f"><label>投标有效期（日历日）</label><input id="validity_days" value="90" placeholder="例：90"></div>
      <div class="f"><label>预算（万元）</label><input id="budget" placeholder="例：1094.8"></div>
    </div>
    <div class="row">
      <div class="f"><label>投标报价（万元）</label><input id="bid_price" placeholder="例：1050"></div>
      <div class="f"><label>服务期限</label><input id="service_period" placeholder="例：自合同签订之日起一年"></div>
    </div>
    <div class="row">
      <div class="f"><label>法定代表人</label><input id="legal_person" placeholder="例：张三"></div>
      <div class="f"><label>联系人</label><input id="contact_person" placeholder="例：李四"></div>
    </div>
    <div class="row">
      <div class="f"><label>联系电话</label><input id="contact_phone" placeholder="例：138XXXXXXXX"></div>
      <div class="f" style="display:flex;align-items:flex-end;gap:8px;padding-bottom:4px">
        <label style="margin-bottom:0;white-space:nowrap"><input id="desensitize" type="checkbox" style="width:auto;margin-right:4px"> 脱敏模式</label>
        <span style="font-size:12px;color:#888">（公司名/联系人/电话替换为______）</span>
      </div>
    </div>
  </div>

  <div class="panel">
    <h2>🏢 公司信息</h2>
    <div class="f"><label>公司简介</label><textarea id="company_desc" rows="3" placeholder="成立时间、注册资本、主营业务、行业地位等"></textarea></div>
    <div class="row">
      <div class="f"><label>注册资金（万元）</label><input id="reg_capital" placeholder="例：1000"></div>
      <div class="f"><label>成立年份</label><input id="found_year" placeholder="例：2010"></div>
    </div>
    <div class="row">
      <div class="f"><label>员工总数</label><input id="staff_count" placeholder="例：500"></div>
      <div class="f"><label>封面logo文字</label><input id="cover_logo_text" placeholder="不填则用投标人全称"></div>
    </div>
    <div class="f"><label>资质证书（每行一个）</label><textarea id="certificates" rows="3" placeholder="例：ISO9001认证&#10;AAA企业信用证书&#10;劳务派遣经营许可证"></textarea></div>
  </div>

  <div class="panel">
    <h2>📄 服务方案</h2>
    <div class="f"><label>服务内容概述</label><textarea id="service_desc" rows="4" placeholder="服务范围、内容概述..."></textarea></div>
    <div class="f"><label>项目重难点及对策</label><textarea id="challenges" rows="3" placeholder="重难点分析及解决措施..."></textarea></div>
    <div class="f"><label>服务保障措施</label><textarea id="guarantee" rows="3" placeholder="质量/安全/应急响应措施..."></textarea></div>
    <div class="f"><label>人员配置方案</label><textarea id="staff_plan" rows="3" placeholder="人员数量、岗位结构、管理团队..."></textarea></div>
  </div>

  <div class="panel">
    <h2>📊 业绩 & 报价</h2>
    <div class="f"><label>类似项目业绩（每行：项目名|甲方|金额|年份）</label><textarea id="projects" rows="3" placeholder="XX医院后勤服务|XX医院|500万元|2023"></textarea></div>
    <div class="f"><label>报价明细（每行：项目|金额|说明）</label><textarea id="price_detail" rows="3" placeholder="人工成本|800万元|含工资社保公积金"></textarea></div>
    <div class="f"><label>优惠承诺</label><input id="discount" placeholder="例：提供免费培训服务"></div>
  </div>

  <div class="panel">
    <h2>✅ 服务承诺</h2>
    <div class="f"><label>服务承诺内容</label><textarea id="commitments" rows="3" placeholder="质量承诺、售后承诺等..."></textarea></div>
    <div class="f"><label>封面标题（正本/副本）</label><input id="cover_title" placeholder="不填则默认「投标文件」"></div>
  </div>
</div>

<script>
document.addEventListener('DOMContentLoaded',()=>{
  document.getElementById('bid_date').value=new Date().toISOString().split('T')[0]
});

function get(id){return document.getElementById(id).value.trim()}
function getChecked(id){return document.getElementById(id).checked}

function showStatus(msg,err){
  const s=document.getElementById('status');
  s.textContent=msg;s.style.display='block';
  s.className='status'+(err?' error':'');
}

function formToJSON(){
  const fields=['project_name','bid_no','bidder_name','bidder_short','bid_date','project_type','budget',
    'bid_price','service_period','legal_person','contact_person','contact_phone',
    'company_desc','reg_capital','found_year','staff_count','certificates',
    'service_desc','challenges','guarantee','staff_plan',
    'projects','price_detail','discount','commitments','cover_title','cover_logo_text','validity_days'];
  const data={};
  fields.forEach(f=>data[f]=get(f));
  data['desensitize']=getChecked('desensitize');
  return data;
}

async function generateWord(){
  const pn=get('project_name');
  if(!pn){showStatus('⚠️ 请填写项目名称！',true);return}
  const bn=get('bidder_name');
  if(!bn){showStatus('⚠️ 请填写投标人全称！',true);return}

  document.getElementById('progress').style.display='block';
  showStatus('正在生成中...');

  try{
    const resp=await fetch('/generate',{
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify(formToJSON())
    });
    if(!resp.ok){
      const err=await resp.json();
      showStatus('❌ '+err.error,true);
      return;
    }
    // 下载文件
    const blob=await resp.blob();
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    a.href=url;
    a.download=(pn.replace(/[/\\\\:*?"<>|]/g,'_'))+'_投标文件.docx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    document.getElementById('progress').style.display='none';
    showStatus('✅ Word标书已生成并下载！');
  }catch(e){
    document.getElementById('progress').style.display='none';
    showStatus('❌ 生成失败：'+e.message,true);
  }
}

async function previewMd(){
  document.getElementById('progress').style.display='block';
  try{
    const resp=await fetch('/preview',{
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify(formToJSON())
    });
    const data=await resp.json();
    document.getElementById('progress').style.display='none';
    const pre=document.getElementById('preview');
    pre.textContent=data.markdown;
    pre.style.display='block';
  }catch(e){
    document.getElementById('progress').style.display='none';
    showStatus('❌ 预览失败：'+e.message,true);
  }
}
</script>
</body>
</html>"""


# ====================== Flask路由 ======================

@app.route('/')
def index():
    return render_template_string(HTML_FORM)


@app.route('/generate', methods=['POST'])
def generate():
    data = request.json
    pn = data.get('project_name', '').strip()
    if not pn:
        return {'error': '项目名称不能为空'}, 400

    desensitize = data.get('desensitize', False)

    # 生成Markdown（传递脱敏标志）
    md = generate_bid_md(data, desensitize=desensitize)

    # 转为Word
    out_name = re.sub(r'[/\\\\:*?"<>|]', '_', pn) + '_投标文件.docx'
    out_path = os.path.join(tempfile.gettempdir(), out_name)

    # 尝试用bid_engine生成
    docx_ok = False
    try:
        sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
        from bid_engine import md_to_docx
        md_to_docx(md, out_path, auto_fix=True)
        if os.path.exists(out_path) and os.path.getsize(out_path) > 0:
            docx_ok = True
    except Exception as e:
        print(f"bid_engine调用失败: {e}")

    if not docx_ok:
        # 不再降级到python-docx，改为报错提示
        return {'error': '请安装 bid-toolkit：pip install bid-toolkit'}, 500

    return send_file(out_path, as_attachment=True,
                     download_name=out_name,
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


@app.route('/preview', methods=['POST'])
def preview():
    data = request.json
    desensitize = data.get('desensitize', False)
    md = generate_bid_md(data, desensitize=desensitize)
    return {'markdown': md}


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    print(f"🚀 标书生成器启动 → http://localhost:{port}")
    print(f"   填完表单点「生成Word」，浏览器直接下载 .docx")
    app.run(host='0.0.0.0', port=port, debug=False)
