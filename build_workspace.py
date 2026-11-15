"""Builds workspace.html (public demo) and apps-script/Index.html (runs inside the sheet) from src/workspace.src.html."""
import json, base64
src = open('src/workspace.src.html', encoding='utf-8').read()
data = json.load(open('src/workspace-data.json', encoding='utf-8'))
bar = json.load(open('src/barriers.json', encoding='utf-8'))
logo = 'data:image/png;base64,' + base64.b64encode(open('images/grps-logo.png', 'rb').read()).decode()
site = 'https://melatkins2-ops.github.io/grps-fss-kssn-integration-test/'
out = (src.replace('/*DATA*/null', json.dumps({'schools': data['schools'], 'resources': data['resources'], 'district': data.get('district', [])}, ensure_ascii=False))
          .replace('/*BARRIERS*/null', json.dumps(bar, ensure_ascii=False)))
open('workspace.html', 'w', encoding='utf-8').write(out.replace('__LOGO__', 'images/grps-logo.png').replace('__SITE__', 'index.html'))
open('apps-script/Index.html', 'w', encoding='utf-8').write(out.replace('__LOGO__', logo).replace('__SITE__', site))
print('built', len(out))
