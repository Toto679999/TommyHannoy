import sys
import threading
import time
import keyboard
import tkinter as tk
from tkinter import simpledialog, messagebox
from datetime import datetime
from collections import Counter
import pandas as pd

# Fenêtre de configuration pour demander le nom de base
root = tk.Tk()
root.withdraw()
messagebox.showinfo(
    "Configuration du logger",
    "Veuillez saisir le nom du log txt et du rapport Excel\nFormat attendu : Nom du client_Partie numéro x"
)
base_name = simpledialog.askstring(
    "Nom du fichier",
    "Entrez le nom de base (Nom du client_Partie numéro x) :"
)
if not base_name:
    messagebox.showerror(
        "Erreur",
        "Un nom de fichier est requis. Le script va se fermer."
    )
    sys.exit(1)

# Code d'arrêt secret
STOP_CODE = "PipiProut666"

# Fichiers selon le nom de base fourni
LOG_FILE = f"{base_name}.txt"
timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
OUTPUT_FILE = f"{base_name}_{timestamp}.xlsx"

# Paramètres
HEARTBEAT_INTERVAL = 120  # secondes
ACTIVE_THRESHOLD = 6      # actions par intervalle pour être actif

# Tampon pour suppressions consécutives
delete_buffer = 0
delete_start = None
buffer_lock = threading.Lock()

# Ouverture du log brut
log = open(LOG_FILE, 'a', encoding='utf-8', buffering=1)

def flush_deletions():
    global delete_buffer, delete_start
    with buffer_lock:
        if delete_buffer > 0:
            ts = delete_start.isoformat()
            log.write(f"{ts}|DEL|{delete_buffer}\n")
            delete_buffer = 0
            delete_start = None

def on_key(event):
    global delete_buffer, delete_start
    if event.event_type != 'down': return
    name = event.name.lower()
    ts = datetime.now().isoformat()
    if name in ('backspace', 'delete'):
        with buffer_lock:
            if delete_buffer == 0:
                delete_start = datetime.now()
            delete_buffer += 1
        return
    flush_deletions()
    if len(name) == 1 and name.isprintable():
        log.write(f"{ts}|INS|{name}\n")
    else:
        log.write(f"{ts}|KEY|{name}\n")

def on_hotkey(name):
    flush_deletions()
    ts = datetime.now().isoformat()
    log.write(f"{ts}|CMD|{name}\n")

def heartbeat_loop():
    while True:
        time.sleep(HEARTBEAT_INTERVAL)
        flush_deletions()
        ts = datetime.now().isoformat()
        log.write(f"{ts}|HEARTBEAT|\n")

def parse_log(path):
    events = []
    with open(path, 'r', encoding='utf-8') as f:
        for line in f:
            parts = line.strip().split('|')
            if len(parts) >= 2:
                ts = datetime.fromisoformat(parts[0])
                etype = parts[1]
                val = parts[2] if len(parts) >= 3 else ''
                events.append((ts, etype, val))
    if not events:
        print("Aucun évènement.")
        return
    df = pd.DataFrame(events, columns=['ts', 'etype', 'val'])
    df.set_index(pd.to_datetime(df.ts), inplace=True)

    # Summary metrics
    total_ins = (df.etype == 'INS').sum()
    total_del = df[df.etype == 'DEL'].val.astype(int).sum()
    cmds = df[df.etype == 'CMD'].val.value_counts().to_dict()
    hbs = df[df.etype == 'HEARTBEAT'].index.sort_values()
    active_cnt = sum(
        1 for i in range(len(hbs)-1)
        if len(df.loc[hbs[i]:hbs[i+1]]) > ACTIVE_THRESHOLD
    )
    active_min = active_cnt * (HEARTBEAT_INTERVAL/60)
    total_min = (df.index[-1] - df.index[0]).total_seconds() / 60
    metrics = {
        'total_duration_min': round(total_min, 2),
        'active_time_min': round(active_min, 2),
        'insertions': int(total_ins),
        'deletions': int(total_del)
    }
    for k, v in cmds.items():
        metrics[k.lower()] = int(v)

    # Resample in 10-minute bins
    ins10 = df[df.etype=='INS'].resample('10min').size()
    del10 = df[df.etype=='DEL'].resample('10min').size()
    act10 = df[df.etype!='HEARTBEAT'].resample('10min').size()
    rate10 = (ins10 / 10).rename('ins_per_min')
    hist_df = pd.concat([ins10.rename('insertions'), del10.rename('deletions')], axis=1).fillna(0).astype(int)
    evo_df = act10.rename('actions').to_frame()
    rate_df = rate10.to_frame().fillna(0)

    # Occurrence des mots
    words, curr = [], []
    for _, r in df.iterrows():
        if r.etype == 'INS':
            curr.append(r.val)
        elif r.etype == 'KEY' and r.val == 'space':
            if curr:
                words.append(''.join(curr)); curr = []
    if curr:
        words.append(''.join(curr))
    wc = Counter(words)
    occ_df = pd.DataFrame([{'word': w, 'count': c} for w, c in wc.items() if c > 1])

    # Écriture Excel avec charts et style
    with pd.ExcelWriter(OUTPUT_FILE, engine='xlsxwriter') as writer:
        pd.DataFrame([metrics]).to_excel(writer, sheet_name='Summary', index=False)
        hist_df.to_excel(writer, sheet_name='Histogram')
        evo_df.to_excel(writer, sheet_name='Evolution')
        rate_df.to_excel(writer, sheet_name='TypingRate')
        if not occ_df.empty:
            occ_df.to_excel(writer, sheet_name='Occurrences', index=False)

        wb = writer.book
        # Format horaire colonne A
        time_fmt = wb.add_format({'num_format': 'hh:mm:ss'})
        for sheet in ['Histogram', 'Evolution', 'TypingRate']:
            ws = writer.sheets[sheet]
            ws.set_column('A:A', 12, time_fmt)
            ws.set_column('B:C', 12)

        # Histogram chart (bar)
        ws1 = writer.sheets['Histogram']
        chart1 = wb.add_chart({'type': 'column'})
        chart1.add_series({
            'name': 'Insertions',
            'categories': f'=Histogram!$A$2:$A${len(hist_df)+1}',
            'values': f'=Histogram!$B$2:$B${len(hist_df)+1}',
            'marker': {'type': 'circle', 'size': 6}
        })
        chart1.add_series({
            'name': 'Deletions',
            'values': f'=Histogram!$C$2:$C${len(hist_df)+1}',
            'marker': {'type': 'circle', 'size': 6}
        })
        chart1.set_title({'name': 'Histogramme Insertions vs Suppressions'})
        chart1.set_x_axis({'name': 'Heure', 'num_format': 'hh:mm:ss', 'major_gridlines': {'visible': True}})
        chart1.set_y_axis({'name': 'Nombre', 'major_gridlines': {'visible': True}})
        chart1.set_legend({'position': 'bottom'})
        chart1.set_style(11)
        ws1.insert_chart('E2', chart1)

        # Evolution chart (line)
        ws2 = writer.sheets['Evolution']
        chart2 = wb.add_chart({'type': 'line'})
        chart2.add_series({
            'name': 'Actions',
            'categories': f'=Evolution!$A$2:$A${len(evo_df)+1}',
            'values': f'=Evolution!$B$2:$B${len(evo_df)+1}',
            'marker': {'type': 'circle', 'size': 6}
        })
        chart2.set_title({'name': "Évolution des actions"})
        chart2.set_x_axis({'name': 'Heure', 'num_format': 'hh:mm:ss', 'major_gridlines': {'visible': True}})
        chart2.set_y_axis({'name': 'Nombre d’actions', 'major_gridlines': {'visible': True}})
        chart2.set_legend({'position': 'bottom'})
        chart2.set_style(11)
        ws2.insert_chart('D2', chart2)

        # Typing rate chart (line)
        ws3 = writer.sheets['TypingRate']
        chart3 = wb.add_chart({'type': 'line'})
        chart3.add_series({
            'name': 'Vitesse de frappe',
            'categories': f'=TypingRate!$A$2:$A${len(rate_df)+1}',
            'values': f'=TypingRate!$B$2:$B${len(rate_df)+1}',
            'marker': {'type': 'circle', 'size': 6}
        })
        chart3.set_title({'name': 'Vitesse de frappe'})
        chart3.set_x_axis({'name': 'Heure', 'num_format': 'hh:mm:ss', 'major_gridlines': {'visible': True}})
        chart3.set_y_axis({'name': 'Caractères/min', 'major_gridlines': {'visible': True}})
        chart3.set_legend({'position': 'bottom'})
        chart3.set_style(11)
        ws3.insert_chart('D2', chart3)

        writer.save()
    print(f"Excel report with styled charts: {OUTPUT_FILE}")


def main():
    keyboard.hook(on_key)
    for combo, name in [('ctrl+c','COPY'),('ctrl+v','PASTE'),('ctrl+x','CUT'),
                        ('ctrl+b','BOLD'),('ctrl+i','ITALIC'),('ctrl+u','UNDERLINE')]:
        keyboard.add_hotkey(combo, lambda n=name: on_hotkey(n))
    threading.Thread(target=heartbeat_loop, daemon=True).start()
    try:
        while input().strip() != STOP_CODE:
            pass
        flush_deletions()
        log.write(f"{datetime.now().isoformat()}|STOP|\n")
    except KeyboardInterrupt:
        flush_deletions()
    finally:
        log.close()
        parse_log(LOG_FILE)

if __name__ == '__main__':
    print("Started. Enter code to stop.")
    main()
