#!/usr/bin/env python3
from __future__ import annotations
import json, os, re
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.sax.saxutils import escape

# Deterministic SCOUT_PL renderer: the only writer of the binary snapshot.
ROOT=Path(__file__).resolve().parents[1]
INPUT=Path(os.environ.get("SCOUT_PL_PAYLOAD", ROOT/"scout_runs/current.json"))
OUTPUT=Path(os.environ.get("SCOUT_PL_OUTPUT", ROOT/"database_festivals.xlsx"))
MAIN="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RELS="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG="http://schemas.openxmlformats.org/package/2006/relationships"

def col(n):
    s=""
    while n:
        n,r=divmod(n-1,26); s=chr(65+r)+s
    return s

def cell(r,c,v):
    ref=f"{col(c)}{r}"
    if v is None: return f'<c r="{ref}"/>'
    if isinstance(v,bool): return f'<c r="{ref}" t="b"><v>{1 if v else 0}</v></c>'
    if isinstance(v,(int,float)) and not isinstance(v,bool): return f'<c r="{ref}" t="n"><v>{v}</v></c>'
    return f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{escape(str(v))}</t></is></c>'

def sheet_xml(headers,rows):
    matrix=[headers,*rows]; maxc=max((len(r) for r in matrix),default=1); maxr=max(len(matrix),1); dim=f'A1:{col(maxc)}{maxr}'
    out=[f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{MAIN}" xmlns:r="{RELS}"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><dimension ref="{dim}"/><sheetData>']
    for rn,row in enumerate(matrix,1):
        out.append(f'<row r="{rn}">')
        for cn in range(1,maxc+1): out.append(cell(rn,cn,row[cn-1] if cn<=len(row) else None))
        out.append("</row>")
    out.append(f'</sheetData><autoFilter ref="{dim}"/></worksheet>')
    return ''.join(out)

def styles_xml():
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="{MAIN}"><numFmts count="0"/><fonts count="1"><font><sz val="11"/><name val="Aptos"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0"/></styleSheet>'

def main():
    payload=json.loads(INPUT.read_text(encoding="utf-8"))
    if payload.get("source")!="SCOUT_PL": raise SystemExit("SCOUT_PL_RENDER_FAIL: source")
    if payload.get("snapshot_mode")!="current_run_only": raise SystemExit("SCOUT_PL_RENDER_FAIL: snapshot_mode")
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}",str(payload.get("run_date",""))): raise SystemExit("SCOUT_PL_RENDER_FAIL: run_date")
    sheets=payload.get('sheets') or {}
    required={'RUN_INFO','OPPORTUNITIES','ORGANIZERS','CONTACTS','CANONICAL_APPEND'}
    if not required.issubset(sheets): raise SystemExit(f"SCOUT_PL_RENDER_FAIL: missing {sorted(required-set(sheets))}")
    names=list(sheets); tmp=OUTPUT.with_suffix('.tmp.xlsx'); tmp.parent.mkdir(parents=True,exist_ok=True)
    with ZipFile(tmp,'w',ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml',content_types(len(names)))
        z.writestr('_rels/.rels',f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        z.writestr('xl/workbook.xml',workbook_xml(names))
        z.writestr('xl/_rels/workbook.xml.rels',rels_xml(names))
        z.writestr('xl/styles.xml',styles_xml())
        for i,name in enumerate(names,1): z.writestr(f'xl/worksheets/sheet{i}.xml',sheet_xml(sheets[name]['headers'],sheets[name]['rows']))
    tmp.replace(OUTPUT); print('SCOUT_PL_RENDER_OK',payload['run_date'])

def workbook_xml(names):
    ss="".join(f'<sheet name="{escape(n)}" sheetId="{i}" r:id="rId{i}"/>' for i,n in enumerate(names,1))
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="{MAIN}" xmlns:r="{RELS}"><sheets>{ss}</sheets></workbook>'

def rels_xml(names):
    rs="".join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/sheet{i}.xml"/>' for i,_ in enumerate(names,1))
    rs+=f'<Relationship Id="rId{len(names)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="/xl/styles.xml"/>'
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{PKG}">{rs}</Relationships>'

def content_types(count):
    osn='<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    osn+= "".join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(1,count+1))
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>{osn}</Types>'

if __name__=='__main__': main()