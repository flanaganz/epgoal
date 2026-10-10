import subprocess,sys,threading
from datetime import datetime
from .paths import paths,CONSOLE_ROOT
from .data import current_snapshot
from .history import add_run

def run_script(which,on_line,on_done):
 target=paths()[which]
 log_dir=CONSOLE_ROOT/'logs';log_dir.mkdir(exist_ok=True)
 log=log_dir/f"{which}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
 def work():
  rc=1
  try:
   with log.open('w',encoding='utf-8') as f:
    p=subprocess.Popen([sys.executable,str(target)],cwd=paths()['root'],stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True,encoding='utf-8',errors='replace',bufsize=1)
    for line in p.stdout:
     f.write(line);f.flush();on_line(line.rstrip())
    rc=p.wait()
  except Exception as e:on_line('FEJL: '+str(e))
  snap=current_snapshot();add_run(snap,'success' if rc==0 else 'failed',log);on_done(rc,log)
 threading.Thread(target=work,daemon=True).start()
