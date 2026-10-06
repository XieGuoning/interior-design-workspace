#!/bin/bash
set -euo pipefail
umask 077
cd /workspace/interior-design-workspace
mkdir -p /workspace/design-tools/{apt/lists/partial,apt/archives/partial,root,cache,config,matplotlib,bin}
cat > /workspace/design-tools/apt/sources.list <<'ONBOARDING_FILE_EOF'
deb [signed-by=/usr/share/keyrings/debian-archive-keyring.gpg] https://deb.debian.org/debian trixie main
deb [signed-by=/usr/share/keyrings/debian-archive-keyring.gpg] https://security.debian.org/debian-security trixie-security main

ONBOARDING_FILE_EOF
cat > /workspace/design-tools/apt/apt.conf <<'ONBOARDING_FILE_EOF'
Dir::Etc::parts "-";
Dir::Etc::main "-";
Dir::Etc::sourcelist "/workspace/design-tools/apt/sources.list";
Dir::Etc::sourceparts "-";
Dir::State::lists "/workspace/design-tools/apt/lists";
Dir::Cache::archives "/workspace/design-tools/apt/archives";
APT::Sandbox::User "agent";

ONBOARDING_FILE_EOF
cat > /workspace/design-tools/activate.sh <<'ONBOARDING_FILE_EOF'
export XDG_CACHE_HOME=/workspace/design-tools/cache
export XDG_CONFIG_HOME=/workspace/design-tools/config
export MPLCONFIGDIR=/workspace/design-tools/matplotlib
export PIP_CACHE_DIR=/workspace/design-tools/cache/pip
export PATH=/workspace/design-tools/bin:$PATH

ONBOARDING_FILE_EOF
cat > /workspace/design-tools/bin/libreoffice <<'ONBOARDING_FILE_EOF'
#!/bin/bash
set -euo pipefail
source /workspace/design-tools/activate.sh
export LD_LIBRARY_PATH=/workspace/design-tools/root/usr/lib/libreoffice/program:/workspace/design-tools/root/usr/lib/x86_64-linux-gnu${LD_LIBRARY_PATH:+:$LD_LIBRARY_PATH}
export URE_BOOTSTRAP=file:///workspace/design-tools/root/usr/lib/libreoffice/program/fundamentalrc
exec /workspace/design-tools/root/usr/lib/libreoffice/program/soffice -env:BRAND_BASE_DIR=file:///workspace/design-tools/root/usr/lib/libreoffice -env:UserInstallation=file:///workspace/design-tools/lo-profile "$@"

ONBOARDING_FILE_EOF
cat > /workspace/design-tools/relocate.py <<'ONBOARDING_FILE_EOF'
from pathlib import Path
root=Path('/workspace/design-tools/root')
for p in root.rglob('*'):
    if p.is_symlink():
        target=p.readlink()
        if target.is_absolute() and (root/str(target).lstrip('/')).exists():
            p.unlink(); p.symlink_to(root/str(target).lstrip('/'))
f=root/'usr/lib/libreoffice/program/fundamentalrc'
s=f.read_text().replace('BRAND_BASE_DIR=file:///usr/lib/libreoffice','BRAND_BASE_DIR=file:///workspace/design-tools/root/usr/lib/libreoffice')
f.write_text(s)

ONBOARDING_FILE_EOF
cat > /workspace/design-tools/validate.py <<'ONBOARDING_FILE_EOF'
from pathlib import Path
import importlib.metadata as md
import ezdxf, openpyxl, docx, pptx, pypdf, PIL, matplotlib
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Inches
from pptx.util import Inches as PI
p=Path('work/onboarding'); p.mkdir(parents=True, exist_ok=True)
d=ezdxf.new('R2007'); d.units=4
m=d.modelspace(); m.add_lwpolyline([(0,0),(1000,0),(1000,1000),(0,1000)],close=True)
m.add_text('中文测试',dxfattribs={'height':100})
d.saveas(p/'r2007.dxf'); r=ezdxf.readfile(p/'r2007.dxf')
assert r.dxfversion=='AC1021' and r.units==4
poly=r.modelspace().query('LWPOLYLINE')[0]; assert poly.closed and len(poly)==4
assert r.modelspace().query('TEXT')[0].dxf.text=='中文测试'
assert not r.audit().has_errors
D=docx.Document(); D.sections[0].page_width=Inches(8.27); D.sections[0].page_height=Inches(11.69)
style=D.styles['Normal']; style.font.name='Noto Sans CJK SC'; style.font.size=Pt(12)
style.element.get_or_add_rPr().rFonts.set(qn('w:eastAsia'),'Noto Sans CJK SC')
for text in ['室内设计辅助环境验证','中文字体预览：客厅、厨房、卧室。','测试图纸：R2007 DXF（AC1021），单位：毫米。','面积示例：1 平方米；仅用于合成测试。']:
    D.add_paragraph(text)
t=D.add_table(rows=1, cols=3); t.style='Table Grid'
for c,s in zip(t.rows[0].cells,['材料名称','数量','计量单位']): c.text=s
for c,s in zip(t.add_row().cells,['地面材料','1','平方米']): c.text=s
D.save(p/'chinese.docx'); assert docx.Document(p/'chinese.docx').paragraphs[0].text=='室内设计辅助环境验证'
w=openpyxl.Workbook(); s=w.active; s.title='材料清单'; s.append(['材料名称','数量','计量单位']); s.append(['地面材料',1,'平方米'])
for row in s:
    for c in row: c.font=openpyxl.styles.Font(name='Noto Sans CJK SC',size=12)
s.column_dimensions['A'].width=24; s.column_dimensions['C'].width=20
s.print_options.horizontalCentered=True; s.sheet_properties.pageSetUpPr.fitToPage=True; s.page_setup.fitToWidth=1; s.page_setup.fitToHeight=1
w.save(p/'chinese-sheet.xlsx'); assert openpyxl.load_workbook(p/'chinese-sheet.xlsx').active['A2'].value=='地面材料'
P=pptx.Presentation(); slide=P.slides.add_slide(P.slide_layouts[6]); box=slide.shapes.add_textbox(PI(1),PI(1),PI(8),PI(3))
for i,text in enumerate(['室内设计辅助环境验证','中文提案预览：客厅、厨房、卧室。']):
    par=box.text_frame.paragraphs[0] if i==0 else box.text_frame.add_paragraph(); par.text=text
    for run in par.runs:
        run.font.name='Noto Sans CJK SC'; run.font.size=pptx.util.Pt(24)
        ea=OxmlElement('a:ea'); ea.set('typeface','Noto Sans CJK SC'); run._r.get_or_add_rPr().append(ea)
P.save(p/'chinese-slides.pptx'); assert len(pptx.Presentation(p/'chinese-slides.pptx').slides)==1
matplotlib.use('Agg')
from matplotlib import pyplot as plt
from matplotlib.font_manager import FontProperties
font=FontProperties(fname='/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc')
fig,ax=plt.subplots(); ax.set_title('室内设计中文图表验证',fontproperties=font); ax.plot([0,1],[0,1]); fig.savefig(p/'chinese-chart.pdf'); plt.close(fig)
(p/'versions.txt').write_text('\n'.join(f'{x}=={md.version(x)}' for x in ['ezdxf','openpyxl','python-docx','python-pptx','pypdf','Pillow','matplotlib'])+'\n')
print('Imports, DXF AC1021 units/geometry/Chinese/audit, DOCX/XLSX/PPTX round-trips, Chinese chart passed')

ONBOARDING_FILE_EOF
cat > /workspace/design-tools/check-pdf.py <<'ONBOARDING_FILE_EOF'
from pathlib import Path
from pypdf import PdfReader
from PIL import Image
import subprocess
p=Path('work/onboarding')
for name, expected in [('chinese','中文字体预览'),('chinese-sheet','地面材料'),('chinese-slides','中文提案预览'),('chinese-chart','室内设计中文图表验证')]:
    f=p/(name+'.pdf'); r=PdfReader(f); assert len(r.pages)==1
    text=''.join(page.extract_text() for page in r.pages)
    assert expected in text and '\ufffd' not in text, (name,text)
    subprocess.run(['pdftotext',str(f),str(p/(name+'.txt'))],check=True)
    fonts=subprocess.check_output(['pdffonts',str(f)],text=True); assert 'NotoSansCJK' in fonts
    subprocess.run(['pdftoppm','-scale-to','1200','-singlefile','-png',str(f),str(p/name)],check=True)
    im=Image.open(p/(name+'.png')); assert im.width>500 and im.height>500
    print(name, 'one page; expected Chinese text; embedded Noto; PNG render passed')

ONBOARDING_FILE_EOF
cat > /workspace/design-tools/check.sh <<'ONBOARDING_FILE_EOF'
#!/bin/bash
set -euo pipefail
cd /workspace/interior-design-workspace
source /workspace/design-tools/activate.sh
.venv/bin/python /workspace/design-tools/validate.py
.venv/bin/python -m pip check
libreoffice --version
fc-match 'Noto Sans CJK SC'
# Remove only these synthetic outputs so stale files cannot mask conversion failures.
rm -f work/onboarding/chinese.pdf work/onboarding/chinese-sheet.pdf work/onboarding/chinese-slides.pdf
libreoffice --headless --convert-to pdf --outdir work/onboarding work/onboarding/chinese.docx work/onboarding/chinese-sheet.xlsx work/onboarding/chinese-slides.pptx
.venv/bin/python /workspace/design-tools/check-pdf.py

ONBOARDING_FILE_EOF
chmod +x /workspace/design-tools/bin/libreoffice /workspace/design-tools/check.sh
source /workspace/design-tools/activate.sh
bash setup.sh
APT_CONFIG=/workspace/design-tools/apt/apt.conf /usr/bin/apt-get update
APT_CONFIG=/workspace/design-tools/apt/apt.conf /usr/bin/apt-get -o Debug::NoLocking=true --download-only -y --no-install-recommends install libreoffice-writer libreoffice-impress libreoffice-calc fonts-noto-cjk poppler-utils
for pkg in /workspace/design-tools/apt/archives/*.deb; do dpkg-deb -x "$pkg" /workspace/design-tools/root; done
mkdir -p /workspace/design-tools/root/etc/libreoffice/registry
cp /workspace/design-tools/root/usr/lib/libreoffice/share/.registry/main.xcd /workspace/design-tools/root/etc/libreoffice/registry/main.xcd
.venv/bin/python /workspace/design-tools/relocate.py
bash /workspace/design-tools/check.sh
.venv/bin/python -m pip freeze > work/installed-packages.txt
