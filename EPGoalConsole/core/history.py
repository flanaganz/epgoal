import sqlite3, json
from .paths import CONSOLE_ROOT
DB=CONSOLE_ROOT/'database'/'epgoal_history.sqlite'
def init_db():
 DB.parent.mkdir(parents=True,exist_ok=True)
 with sqlite3.connect(DB) as c:
  c.execute("CREATE TABLE IF NOT EXISTS runs(id INTEGER PRIMARY KEY, created_at TEXT, status TEXT, channels INTEGER, programmes INTEGER, empty_channels INTEGER, kept_epgshare INTEGER, fallback_channels INTEGER, missing_channels INTEGER, source_counts TEXT, log_file TEXT)")
def add_run(snapshot,status='success',log_file=''):
 init_db()
 with sqlite3.connect(DB) as c:
  c.execute('INSERT INTO runs(created_at,status,channels,programmes,empty_channels,kept_epgshare,fallback_channels,missing_channels,source_counts,log_file) VALUES(?,?,?,?,?,?,?,?,?,?)',(snapshot['created_at'],status,snapshot['channels'],snapshot['programmes'],snapshot['empty_channels'],snapshot['kept_epgshare'],snapshot['fallback_channels'],snapshot['missing_channels'],json.dumps(snapshot['source_counts']),str(log_file)))
def runs(limit=100):
 init_db()
 with sqlite3.connect(DB) as c:
  c.row_factory=sqlite3.Row
  return [dict(x) for x in c.execute('SELECT * FROM runs ORDER BY id DESC LIMIT ?',(limit,))]
