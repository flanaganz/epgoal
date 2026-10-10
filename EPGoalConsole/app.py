from nicegui import ui
import json
from core.paths import paths, CONSOLE_ROOT, settings
from core.data import load_channels, epgoal_stats, merge_stats, current_snapshot
from core.history import init_db, runs, add_run
from core.rules import derive_title
from core.runner import run_script

init_db()


def navigation():
    with ui.header().classes('items-center justify-between bg-blue-10'):
        ui.label('EPGoal Console').classes('text-xl font-bold')
        ui.label(str(paths()['root'])).classes('text-xs')
    with ui.left_drawer(value=True).classes('bg-grey-10'):
        ui.button('Dashboard', on_click=lambda: ui.navigate.to('/'), icon='dashboard').props('flat').classes('w-full')
        ui.button('Kanaler', on_click=lambda: ui.navigate.to('/channels'), icon='tv').props('flat').classes('w-full')
        ui.button('Kørsel', on_click=lambda: ui.navigate.to('/run'), icon='play_circle').props('flat').classes('w-full')
        ui.button('Titelregler', on_click=lambda: ui.navigate.to('/rules'), icon='auto_fix_high').props('flat').classes('w-full')
        ui.button('Historik', on_click=lambda: ui.navigate.to('/history'), icon='insights').props('flat').classes('w-full')


def metric(title, value, sub=''):
    with ui.card().classes('w-56'):
        ui.label(title).classes('text-sm text-grey-5')
        ui.label(str(value)).classes('text-3xl font-bold')
        if sub:
            ui.label(sub).classes('text-xs text-grey-5')


@ui.page('/')
def dashboard_page():
    navigation()
    ui.label('Dashboard').classes('text-3xl font-bold')
    e = epgoal_stats()
    m = merge_stats()
    with ui.row().classes('gap-4 flex-wrap'):
        metric('Kanaler', e['channels'])
        metric('Programmer', f"{e['programmes']:,}")
        metric('Kanaler uden EPG', e['empty'])
        metric('Fallback-kanaler', m['supplied'])
        metric('Uden fallback', m['missing'])
    ui.separator()
    ui.label('Kildefordeling').classes('text-xl')
    chart_data = [{'name': k, 'value': v} for k, v in m['sources'].items()]
    if chart_data:
        ui.echart({'tooltip': {'trigger': 'item'}, 'series': [{'type': 'pie', 'radius': ['45%', '72%'], 'data': chart_data}]}).classes('w-full h-80')
    else:
        ui.label('Ingen merge-statistik fundet endnu.').classes('text-orange-5')
    history = list(reversed(runs(60)))
    if history:
        ui.label('Historik').classes('text-xl')
        ui.echart({
            'tooltip': {'trigger': 'axis'},
            'legend': {'data': ['Programmer', 'Problemkanaler']},
            'xAxis': {'type': 'category', 'data': [x['created_at'][5:16] for x in history]},
            'yAxis': {'type': 'value'},
            'series': [
                {'name': 'Programmer', 'type': 'line', 'data': [x['programmes'] for x in history]},
                {'name': 'Problemkanaler', 'type': 'line', 'data': [x['empty_channels'] + x['missing_channels'] for x in history]},
            ],
        }).classes('w-full h-96')


@ui.page('/channels')
def channels_page():
    navigation()
    ui.label('Kanaler').classes('text-3xl font-bold')
    channels = load_channels()
    epg = epgoal_stats()
    override_file = CONSOLE_ROOT / 'config' / 'source_overrides.json'
    overrides = json.loads(override_file.read_text(encoding='utf-8'))
    rows = []
    for index, channel in enumerate(channels):
        count = epg['counts'].get(channel['output'], 0)
        rows.append({'id': index, **channel, 'programmes': count, 'status': 'OK' if count else 'INGEN EPG'})
    columns = [
        {'name': 'channel', 'label': 'Kanal', 'field': 'channel', 'sortable': True},
        {'name': 'source_mode', 'label': 'Kilde', 'field': 'source_mode', 'sortable': True},
        {'name': 'source_id', 'label': 'Kilde-ID', 'field': 'source_id'},
        {'name': 'output', 'label': 'Output-ID', 'field': 'output'},
        {'name': 'programmes', 'label': 'Programmer', 'field': 'programmes', 'sortable': True},
        {'name': 'status', 'label': 'Status', 'field': 'status', 'sortable': True},
    ]
    table = ui.table(columns=columns, rows=rows, row_key='id', pagination=20).classes('w-full').props('dense')
    table.add_slot('body-cell-source_mode', r"""
        <q-td :props="props">
          <q-select dense borderless emit-value map-options
            :options="['auto','epgshare','openepg','bss','exclude']"
            v-model="props.row.source_mode"
            @update:model-value="() => $parent.$emit('source_change', props.row)" />
        </q-td>
    """)

    def changed(message):
        row = message.args
        name = row['channel']
        mode = row['source_mode']
        if mode == 'auto':
            overrides.pop(name, None)
        else:
            overrides[name] = {
                'mode': 'forced', 'source': mode,
                'source_channel_id': row.get('source_id', ''),
                'output_id': row.get('output', ''), 'enabled': True,
            }
        override_file.write_text(json.dumps(overrides, ensure_ascii=False, indent=2), encoding='utf-8')
        ui.notify(f'Gemt: {name} -> {mode}', type='positive')

    table.on('source_change', changed)


@ui.page('/run')
def run_page():
    navigation()
    ui.label('Kørsel').classes('text-3xl font-bold')
    state = {'running': False}
    log = ui.log(max_lines=2500).classes('w-full h-[620px] bg-black text-green-4')

    def start(which):
        if state['running']:
            ui.notify('En kørsel er allerede i gang', type='warning')
            return
        state['running'] = True
        log.clear()
        log.push('Starter ' + which + ' ...')

        def line(value):
            ui.timer(0, lambda: log.push(value), once=True)

        def done(code, path):
            def finish():
                state['running'] = False
                log.push(f'FÆRDIG, exit={code}, log={path}')
                ui.notify('Kørsel færdig' if code == 0 else 'Kørsel fejlede', type='positive' if code == 0 else 'negative')
            ui.timer(0, finish, once=True)

        run_script(which, line, done)

    with ui.row():
        ui.button('Byg frisk EPG', on_click=lambda: start('pipeline'), icon='play_arrow').props('color=primary')
        ui.button('Gem artwork-valg', on_click=lambda: start('save_choices'), icon='save').props('color=secondary')
        ui.button('Registrér snapshot', on_click=lambda: (add_run(current_snapshot(), 'manual', ''), ui.notify('Snapshot gemt', type='positive')), icon='camera')


@ui.page('/rules')
def rules_page():
    navigation()
    ui.label('Titelregler').classes('text-3xl font-bold')
    title = ui.input('Testtitel', value='Folk og Fæ: (6:7) Frank fister Margrethe med piskeris.').classes('w-full')
    result = ui.label().classes('text-xl')

    def test():
        cleaned, rule = derive_title(title.value)
        result.set_text(f'TMDb-søgetitel: {cleaned} | Regel: {rule}')

    ui.button('Test titelrensning', on_click=test)
    ui.markdown('Originaltitlen i XML ændres ikke. Reglerne bruges kun til TMDb-opslag.')
    test()


@ui.page('/history')
def history_page():
    navigation()
    ui.label('Kørselshistorik').classes('text-3xl font-bold')
    history = runs(500)
    columns = [{'name': x, 'label': x.replace('_', ' ').title(), 'field': x, 'sortable': True} for x in ['created_at', 'status', 'channels', 'programmes', 'empty_channels', 'fallback_channels', 'missing_channels', 'log_file']]
    ui.table(columns=columns, rows=history, row_key='id', pagination=25).classes('w-full').props('dense')


settings_data = settings()
ui.run(title='EPGoal Console', host=settings_data['host'], port=settings_data['port'], reload=False, show=False, favicon='📺', dark=True)
