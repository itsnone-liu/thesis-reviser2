# -*- coding: utf-8 -*-
"""Isolated DOCX render/audit worker for the web service."""
import json, os, sys
from renderer import render
from audit_final import audit_docx_only

def main():
    spec=json.load(open(sys.argv[1],encoding='utf-8'))
    def progress(msg,pct): print(json.dumps({'message':msg,'progress':int(pct)},ensure_ascii=False),flush=True)
    render(spec['txt'], spec['docx'], spec['ptype'], spec.get('cover'), None, progress)
    audit=audit_docx_only(spec['docx'],spec['ptype'])
    # 生图额度告警持久化进审计报告（renderer.render 已实时推过进度消息；
    # 这里让任务完成后的 audit_json 也带警告，用户端/管理端报告可见）
    try:
        from core import image_quota_warning
        meta_path = spec['docx'] + '.render_meta.json'
        if os.path.exists(meta_path):
            meta = json.load(open(meta_path, encoding='utf-8'))
            info = image_quota_warning(meta.get('drawings_dir') or '')
            if info:
                audit.setdefault('problems', []).append(
                    f"⚠️ 生图额度已满（{info.get('reason','')}｜{info.get('time','')}）："
                    f"GPT图纸已用本地确定性工程图兜底，额度恢复后重跑可换回GPT图")
                audit['image_quota_exhausted'] = info
    except Exception as exc:
        print(json.dumps({'message': f'额度告警读取失败: {exc}', 'progress': 95}, ensure_ascii=False), flush=True)
    print(json.dumps({'done':True,'audit':audit},ensure_ascii=False),flush=True)
if __name__=='__main__': main()
