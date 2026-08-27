import pdfplumber
p=r'C:\Users\20159\Desktop\实习\百度「文心千帆 + 全场景 AI」垂类 AI 产品经理岗位线上实习项目全案(1).pdf'
with pdfplumber.open(p) as pdf:
 print('pages',len(pdf.pages))
 for i,page in enumerate(pdf.pages):
  text=(page.extract_text() or '').replace('\u00a0',' ')
  keys=['项目背景','项目目标','岗位','培养','核心任务','阶段','能力','周','实习','产品经理','职责','交付']
  if any(k in text for k in keys):
   print(f'\n---PAGE {i+1}---\n{text[:5000]}')
