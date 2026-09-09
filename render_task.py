# -*- coding: utf-8 -*-
"""Isolated DOCX render/audit worker for the web service."""
import json, sys
from renderer import render
from audit_final import audit_docx_only

def main():
    spec=json.load(open(sys.argv[1],encoding='utf-8'))
    def progress(msg,pct): print(json.dumps({'message':msg,'progress':int(pct)},ensure_ascii=False),flush=True)
    render(spec['txt'], spec['docx'], spec['ptype'], spec.get('cover'), None, progress)
    audit=audit_docx_only(spec['docx'],spec['ptype'])
    print(json.dumps({'done':True,'audit':audit},ensure_ascii=False),flush=True)
if __name__=='__main__': main()
