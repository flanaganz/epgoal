import json,re
from .paths import CONSOLE_ROOT
FILE=CONSOLE_ROOT/'config'/'title_rules.json'
def load_rules(): return json.loads(FILE.read_text(encoding='utf-8')).get('rules',[])
def derive_title(title):
 for rule in load_rules():
  if not rule.get('enabled',True):continue
  if re.search(rule['pattern'],title,re.I):return re.sub(rule['pattern'],rule['replacement'],title,flags=re.I).strip(),rule['name']
 return title,'Ingen regel'
