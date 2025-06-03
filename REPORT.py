import os
import sys
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from datetime import datetime

# 1. Sélection du dossier contenant les rapports Excel
def select_folder():
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo(
        "Agrégation des rapports",
        "Sélectionnez le dossier contenant les rapports Excel à agréger."
    )
    folder = filedialog.askdirectory(title="Dossier de rapports Excel")
    if not folder:
        messagebox.showerror("Erreur", "Aucun dossier sélectionné.")
        sys.exit(1)
    return folder

# 2. Saisie des métadonnées et évaluation globale
def ask_metadata_and_evaluation():
    client_code = simpledialog.askstring("Code client", "Entrez le CODE CLIENT :")
    if not client_code:
        messagebox.showerror("Erreur", "CODE CLIENT requis.")
        sys.exit(1)
    date_reunion = simpledialog.askstring("Date réunion", "Entrez la DATE RÉUNION (JJ/MM/AAAA) :")
    if not date_reunion:
        messagebox.showerror("Erreur", "DATE RÉUNION requise.")
        sys.exit(1)
    duree_audio = simpledialog.askstring("Durée audio", "Entrez la DURÉE AUDIO (HH:MM) :")
    if not duree_audio:
        messagebox.showerror("Erreur", "DURÉE AUDIO requise.")
        sys.exit(1)
    format_ = simpledialog.askstring("Format", "Entrez le FORMAT (SYB/SYD/CRS) :")
    if not format_:
        messagebox.showerror("Erreur", "FORMAT requis.")
        sys.exit(1)

    def ask_score(prompt):
        val = simpledialog.askinteger("Évaluation", f"{prompt} (0-10) :", minvalue=0, maxvalue=10)
        if val is None:
            messagebox.showerror("Erreur", "Note requise.")
            sys.exit(1)
        return val

    raw_scores = {
        'Bruit de fond': ask_score("Bruit de fond : 0=aucun,10=très présent"),
        'Interruptions': ask_score("Interruptions : 0=jamais,10=constamment"),
        'Complexité lexicale': ask_score("Complexité lexicale : 0=basique,10=très complexe")
    }

    # Calcul de la durée en minutes
    try:
        h, m = map(int, duree_audio.split(':'))
        dur_min = h * 60 + m
    except:
        dur_min = 0

    # Facteur de modulation selon durée entre 30min et 600min
    min_dur, max_dur = 30, 600
    ratio = max(0, min((dur_min - min_dur) / (max_dur - min_dur), 1))
    factor = 1 + ratio

    # Score du format de réunion
    fmt_map = {'SYB': 10, 'SYD': 5, 'CRS': 0}
    raw_scores['Format'] = fmt_map.get(format_.upper(), 0)

    # Application du facteur, clamp à 10
    adjusted_scores = {k: min(v * factor, 10) for k, v in raw_scores.items()}

    # Poids des critères (somme = 1)
    base_weights = {
        'Bruit de fond': 0.15,
        'Interruptions': 0.15,
        'Format': 0.20,
        'Complexité lexicale': 0.30
    }
    total_base = sum(base_weights.values())
    weights = {k: w / total_base for k, w in base_weights.items()}

    # Note globale pondérée
    global_note = sum(adjusted_scores[k] * weights[k] for k in adjusted_scores)
    return client_code, date_reunion, duree_audio, format_, adjusted_scores, global_note, dur_min

# 3. Lecture et agrégation des feuilles Summary
def aggregate_summaries(folder, client_code, date_reunion, duree_audio, format_):
    files = [f for f in os.listdir(folder) if f.lower().endswith('.xlsx')]
    rows = []
    for fname in files:
        path = os.path.join(folder, fname)
        try:
            df_sum = pd.read_excel(path, sheet_name='Summary')
        except:
            continue
        if df_sum.empty:
            continue
        row = df_sum.iloc[0].to_dict()
        row.update({
            'report': fname,
            'CODE CLIENT': client_code,
            'DATE REUNION': date_reunion,
            'DUREE AUDIO': duree_audio,
            'FORMAT': format_
        })
        rows.append(row)
    df = pd.DataFrame(rows)

    # Ajout calculs d'actions
    action_cols = [c for c in df.columns if c not in (
        'report','CODE CLIENT','DATE REUNION','DUREE AUDIO','FORMAT',
        'total_duration_min','active_time_min'
    )]
    df['total_actions'] = df[action_cols].sum(axis=1)
    df['actions_per_min'] = df['total_actions'] / df['total_duration_min']

    # Organisation finale
    id_cols = ['report','CODE CLIENT','DATE REUNION','DUREE AUDIO','FORMAT']
    time_cols = ['total_duration_min','active_time_min','actions_per_min']
    df = df[id_cols + time_cols + action_cols]
    return df, time_cols + action_cols

# 4. Génération du rapport agrégé (Summary + Pivot)
def write_aggregated_report(df, metrics_cols, scores, global_note, folder, dur_min):
    pivot_df = df.set_index('report')[metrics_cols].T
    timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
    outfile = os.path.join(folder, f"aggregated_metrics_{timestamp}.xlsx")

    with pd.ExcelWriter(outfile, engine='xlsxwriter') as writer:
        # Feuille Summary
        df.to_excel(writer, sheet_name='Summary', index=False)
        ws_sum = writer.sheets['Summary']
        max_rep = max(len(r) for r in df['report']) + 2
        ws_sum.set_column('A:A', max_rep)
        ws_sum.set_column('B:E', 20)
        ws_sum.set_column('F:H', 16)
        ws_sum.set_column('I:Z', 14)

        # Feuille Pivot
        wb = writer.book
        ws = wb.add_worksheet('Pivot')
        bold = wb.add_format({'bold': True})
        header_fmt = wb.add_format({'bold': True, 'bg_color': '#D9D9D9', 'align':'center'})
        num_fmt = wb.add_format({'num_format': '0.00', 'align':'center'})
        green_fmt = wb.add_format({'font_color':'green', 'num_format': '+0.00;-0.00;0.00', 'align':'center'})
        red_fmt   = wb.add_format({'font_color':'red',   'num_format': '+0.00;-0.00;0.00', 'align':'center'})
        text_fmt  = wb.add_format({'align':'center'})

        # Bloc identité
        meta = [
            ('CODE CLIENT', df.at[0,'CODE CLIENT']),
            ('DATE REUNION', df.at[0,'DATE REUNION']),
            ('DUREE AUDIO', df.at[0,'DUREE AUDIO']),
            ('FORMAT', df.at[0,'FORMAT'])
        ]
        for i, (lbl, val) in enumerate(meta):
            ws.write(i, 0, lbl, bold)
            ws.write(i, 1, val)

        # Bloc Évaluation TURBO
        start_eval = len(meta) + 1
        ws.write(start_eval, 0, 'Évaluation TURBO', bold)
        for off, (crit, val) in enumerate(scores.items(), start=1):
            ws.write(start_eval+off, 0, crit, bold)
            ws.write(start_eval+off, 1, val, num_fmt)
        ws.write(start_eval+len(scores)+1, 0, 'Note globale', bold)
        ws.write(start_eval+len(scores)+1, 1, global_note, num_fmt)

        # Bloc Sessions de travail
        start_data = start_eval + len(scores) + 3
        ws.write(start_data, 0, 'Sessions de travail', bold)
        for c, rpt in enumerate(pivot_df.columns, start=1):
            ws.write(start_data, c, rpt, header_fmt)
        total_col = len(pivot_df.columns) + 1
        ws.write(start_data, total_col, 'Total', header_fmt)
        for r, metric in enumerate(pivot_df.index, start=start_data+1):
            ws.write(r, 0, metric, bold)
            for c, rpt in enumerate(pivot_df.columns, start=1):
                v = pivot_df.at[metric, rpt]
                ws.write(r, c, '' if pd.isna(v) else v, num_fmt)
            tot = pivot_df.loc[metric].dropna().sum()
            ws.write(r, total_col, tot, num_fmt)

        # Différence Prév/Réel (min)
        diff_row = start_data + len(pivot_df.index) + 2
        ws.write(diff_row, 0, 'Diff Prév/Réel (min)', bold)
        rate_map = {'CRS':3.29, 'SYB':1.82, 'SYD':2.5}
        rate = rate_map.get(df.at[0,'FORMAT'].upper(), 1)
        predicted_total = dur_min * rate
        actual_total = df['total_duration_min'].sum()
        diff_total = predicted_total - actual_total
        fmt = green_fmt if diff_total > 0 else red_fmt
        ws.write(diff_row, total_col, diff_total, fmt)

        # Ajustement largeur des colonnes
        ws.set_column('A:A', 30)
        for col in range(1, total_col+1):
            ws.set_column(col, col, 20)

    print(f"Rapport agrégé généré : {outfile}")

# MAIN
if __name__ == '__main__':
    folder = select_folder()
    client_code, date_reunion, duree_audio, format_, scores, global_note, dur_min = ask_metadata_and_evaluation()
    df, metrics_cols = aggregate_summaries(folder, client_code, date_reunion, duree_audio, format_)
    write_aggregated_report(df, metrics_cols, scores, global_note, folder, dur_min)
