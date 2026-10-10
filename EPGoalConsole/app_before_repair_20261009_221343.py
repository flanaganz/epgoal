from nicegui import ui
import json
from pathlib import Path
from core.paths import paths,CONSOLE_ROOT,settings
from core.data import load_channels,epgoal_stats,merge_stats,current_snapshot
from core.history import init_db,runs,add_run
from core.rules import derive_title
from core.runner import run_script

init_db(); dark=ui.dark_mode(); dark.enable()
COLORS={'ok':'positive','warn':'warning','bad':'negative'}

def metric(title,value,sub=''):
 with ui.card().classes('w-56'):
  ui.label(title).classes('text-sm text-grey-5');ui.label(str(value)).classes('text-3xl font-bold');ui.label(sub).classes('text-xs text-grey-5')

def dashboard():
 ui.label('EPGoal Console').classes('text-3xl font-bold')
 e=epgoal_stats();m=merge_stats()
 with ui.row().classes('gap-4 wrap'):
  metric('Kanaler',e['channels']);metric('Programmer',f"{e['programmes']:,}");metric('Kanaler uden EPG',e['empty']);metric('Fallback-kanaler',m['supplied']);metric('Uden fallback',m['missing'])
 ui.separator();ui.label('Kildefordeling').classes('text-xl')
 data=[{'name':k,'value':v} for k,v in m['sources'].items()]
 ui.echart({'tooltip':{'trigger':'item'},'series':[{'type':'pie','radius':['45%','72%'],'data':data}]}).classes('w-full h-80')
 hist=list(reversed(runs(60)))
 if hist:
  ui.label('Historik').classes('text-xl')
  ui.echart({'tooltip':{'trigger':'axis'},'legend':{'data':['Programmer','Problemkanaler']},'xAxis':{'type':'category','data':[x['created_at'][5:16] for x in hist]},'yAxis':{'type':'value'},'series':[{'name':'Programmer','type':'line','data':[x['programmes'] for x in hist]},{'name':'Problemkanaler','type':'line','data':[x['empty_channels']+x['missing_channels'] for x in hist]}]}).classes('w-full h-96')

def channels_page():
 ui.label('Kanaler').classes('text-3xl font-bold');chs=load_channels();e=epgoal_stats();overfile=CONSOLE_ROOT/'config'/'source_overrides.json';over=json.loads(overfile.read_text(encoding='utf-8'))
 rows=[]
 for i,c in enumerate(chs):
  count=e['counts'].get(c['output'],0);rows.append({'id':i,**c,'programmes':count,'status':'OK' if count else 'INGEN EPG'})
 cols=[{'name':'channel','label':'Kanal','field':'channel','sortable':True},{'name':'source_mode','label':'Kilde','field':'source_mode','sortable':True},{'name':'source_id','label':'Kilde-ID','field':'source_id'},{'name':'output','label':'Output-ID','field':'output'},{'name':'programmes','label':'Programmer','field':'programmes','sortable':True},{'name':'status','label':'Status','field':'status'}]
 table=ui.table(columns=cols,rows=rows,row_key='id',pagination=20).classes('w-full').props('dense')
 with table.add_slot('body-cell-source_mode'):
  ui.html("""<q-td :props=\"props\"><q-select dense borderless emit-value map-options :options=\"['auto','epgshare','openepg','bss','exclude']\" v-model=\"props.row.source_mode\" @update:model-value=\"() => $parent.$emit('source_change', props.row)\" /></q-td>""")
 def changed(msg):
  r=msg.args;name=r['channel'];mode=r['source_mode']
  if mode=='auto':over.pop(name,None)
  else:over[name]={'mode':'forced','source':mode,'source_channel_id':r.get('source_id',''),'output_id':r.get('output',''),'enabled':True}
  overfile.write_text(json.dumps(over,ensure_ascii=False,indent=2),encoding='utf-8');ui.notify(f'Gemt: {name} → {mode}',type='positive')
 table.on('source_change',changed)

def run_page():
 ui.label('Kørsel').classes('text-3xl font-bold');state={'running':False};log=ui.log(max_lines=2500).classes('w-full h-[620px] bg-black text-green-4')
 def start(which):
  if state['running']:ui.notify('En kørsel er allerede i gang',type='warning');return
  state['running']=True;log.clear();log.push('Starter '+which+' ...')
  def line(x):ui.timer(0,lambda:log.push(x),once=True)
  def done(rc,path):
   def finish():state['running']=False;log.push(f'FÆRDIG, exit={rc}, log={path}');ui.notify('Kørsel færdig' if rc==0 else 'Kørsel fejlede',type='positive' if rc==0 else 'negative')
   ui.timer(0,finish,once=True)
  run_script(which,line,done)
 with ui.row():
  ui.button('Byg frisk EPG',on_click=lambda:start('pipeline'),icon='play_arrow').props('color=primary')
  ui.button('Gem artwork-valg',on_click=lambda:start('save_choices'),icon='save').props('color=secondary')
  ui.button('Registrér snapshot',on_click=lambda:(add_run(current_snapshot(),'manual',''),ui.notify('Snapshot gemt',type='positive')),icon='camera')

def rules_page():
 ui.label('Titelregler').classes('text-3xl font-bold');title=ui.input('Testtitel',value='Folk og Fæ: (6:7) Frank fister Margrethe med piskeris.').classes('w-full');result=ui.label().classes('text-xl')
 def test():
  cleaned,rule=derive_title(title.value);result.set_text(f'TMDb-søgetitel: {cleaned} | Regel: {rule}')
 ui.button('Test titelrensning',on_click=test);test()
 ui.markdown('Originaltitlen i XML ændres ikke. Reglerne bruges kun til TMDb-opslag.')

def history_page():
 ui.label('Kørselshistorik').classes('text-3xl font-bold');r=runs(500)
 cols=[{'name':x,'label':x.replace('_',' ').title(),'field':x,'sortable':True} for x in ['created_at','status','channels','programmes','empty_channels','fallback_channels','missing_channels','log_file']]
 ui.table(columns=cols,rows=r,row_key='id',pagination=25).classes('w-full').props('dense')

with ui.header().classes('items-center justify-between'):
 ui.label('EPGoal Console').classes('text-xl font-bold');ui.label(str(paths()['root'])).classes('text-xs')
with ui.left_drawer(value=True).classes('bg-grey-10'):
 ui.button('Dashboard',on_click=lambda:ui.navigate.to('/'),icon='dashboard').props('flat').classes('w-full')
 ui.button('Kanaler',on_click=lambda:ui.navigate.to('/channels'),icon='tv').props('flat').classes('w-full')
 ui.button('Kørsel',on_click=lambda:ui.navigate.to('/run'),icon='play_circle').props('flat').classes('w-full')
 ui.button('Titelregler',on_click=lambda:ui.navigate.to('/rules'),icon='auto_fix_high').props('flat').classes('w-full')
 ui.button('Historik',on_click=lambda:ui.navigate.to('/history'),icon='insights').props('flat').classes('w-full')
@ui.page('/')
def p1():dashboard()
@ui.page('/channels')
def p2():channels_page()
@ui.page('/run')
def p3():run_page()
@ui.page('/rules')
def p4():rules_page()
@ui.page('/history')
def p5():history_page()

s=settings();ui.run(title='EPGoal Console',host=s['host'],port=s['port'],reload=False,show=False,favicon='📺')
