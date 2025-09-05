# -*- coding: utf-8 -*-
import os
import re
from collections import Counter
import pandas as pd
from flask import Flask, request, jsonify, session, send_from_directory, send_file
from werkzeug.utils import secure_filename
import io
from unidecode import unidecode

# --- Configuração Inicial ---
app = Flask(__name__, static_url_path='/static')
app.secret_key = 'chave-secreta-para-manter-a-sessao'
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# --- Funções de Validação e Tratamento ---
def standardize_text(text):
    if not isinstance(text, str): return ""
    text = unidecode(text.lower())
    text = re.sub(r'[^\w\s]', '', text)
    return ' '.join(text.split())

def is_potentially_invalid_code(text):
    if not isinstance(text, str): return False
    has_digits = bool(re.search(r'\d', text))
    has_letters = bool(re.search(r'[a-zA-Z]', text))
    if has_digits and has_letters: return True
    if text.isdigit() and len(text) >= 8: return True
    return False

# --- Rotas da Aplicação ---
@app.route('/')
def index():
    session.clear()
    return send_from_directory('.', 'index.html')

# =====================================================================
# FLUXO 1: ANÁLISE DE INCONSISTÊNCIAS
# =====================================================================
@app.route('/api/upload_for_analysis', methods=['POST'])
def upload_for_analysis():
    session.clear()
    if 'file' not in request.files: return jsonify({"error": "Nenhum ficheiro enviado"}), 400
    file = request.files['file']
    filename = secure_filename(file.filename)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)
    session['analysis_filepath'] = filepath
    session['analysis_filename'] = filename
    session['staged_removals'] = {}
    xls = pd.ExcelFile(filepath)
    return jsonify({"fileName": filename, "sheets": xls.sheet_names, "columns": list(pd.read_excel(xls, sheet_name=xls.sheet_names[0]).columns)})

@app.route('/api/get_columns_for_analysis', methods=['POST'])
def get_columns_for_analysis():
    sheet_name = request.json.get('sheetName')
    filepath = session.get('analysis_filepath')
    df = pd.read_excel(filepath, sheet_name=sheet_name)
    return jsonify({"columns": list(df.columns)})

@app.route('/api/analyze', methods=['POST'])
def analyze():
    data = request.json
    sheet_title, columns_to_check, filepath = data.get('sheetTitle'), data.get('columns'), session.get('analysis_filepath')
    min_repetitions = 3
    if not all([sheet_title, columns_to_check, filepath]): return jsonify({"error": "Dados insuficientes"}), 400

    try:
        df = pd.read_excel(filepath, sheet_name=sheet_title)
        frequent_values = {col: set(df[col].astype(str).apply(standardize_text).value_counts()[lambda x: x >= min_repetitions].index) for col in columns_to_check if col in df.columns}
        
        inconsistent_rows_data = []
        for index, row in df.iterrows():
            is_row_inconsistent = False
            issue_reason = ""
            for col in columns_to_check:
                cell_value = row.get(col, '')
                if is_potentially_invalid_code(cell_value):
                    is_row_inconsistent = True
                    issue_reason = f"Formato inválido em '{col}'"
                    break
                standardized_value = standardize_text(cell_value)
                if standardized_value not in frequent_values.get(col, set()):
                    is_row_inconsistent = True
                    issue_reason = f"Valor infrequente em '{col}'"
                    break
            if is_row_inconsistent:
                row_data = {col_header: row.get(col_header, '') for col_header in columns_to_check}
                row_data['rowIndex'] = index + 2
                row_data['issue'] = issue_reason
                inconsistent_rows_data.append(row_data)

        return jsonify({"headers": columns_to_check, "rows": inconsistent_rows_data})
    except Exception as e:
        return jsonify({'error': f'Erro ao analisar dados: {e}'}), 500

@app.route('/api/stage_tab_changes', methods=['POST'])
def stage_tab_changes():
    data = request.json
    sheet_title, rows_to_remove = data.get('sheetTitle'), data.get('rowsToRemove')
    if 'staged_removals' not in session: session['staged_removals'] = {}
    staged = session['staged_removals']
    staged[sheet_title] = rows_to_remove
    session['staged_removals'] = staged
    return jsonify({"success": True, "message": f"Decisões para '{sheet_title}' foram guardadas."})

@app.route('/api/process_all_staged_changes', methods=['POST'])
def process_all_staged_changes():
    staged_removals = session.get('staged_removals', {})
    filepath = session.get('analysis_filepath')
    if not os.path.exists(filepath): return jsonify({"error": "Ficheiro original não encontrado"}), 404
        
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        xls = pd.ExcelFile(filepath)
        for name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=name)
            if name in staged_removals and staged_removals[name]:
                df.drop(index=[idx - 2 for idx in staged_removals[name]], inplace=True)
            df.to_excel(writer, sheet_name=name, index=False)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name=f"processado_final_{session.get('analysis_filename', 'f.xlsx')}", mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# =====================================================================
# FLUXO 2: ATUALIZAÇÃO DE PLANILHA
# =====================================================================
def upload_and_store(file, session_prefix):
    filename = secure_filename(file.filename)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)
    session[f'{session_prefix}_filepath'] = filepath
    session[f'{session_prefix}_filename'] = filename
    return {"fileName": filename, "sheets": pd.ExcelFile(filepath).sheet_names}

@app.route('/api/upload_base', methods=['POST'])
def upload_base_file():
    session.clear()
    if 'file' not in request.files: return jsonify({"error": "Nenhum ficheiro enviado"}), 400
    return jsonify(upload_and_store(request.files['file'], 'base'))

@app.route('/api/upload_update', methods=['POST'])
def upload_update_file():
    if 'file' not in request.files: return jsonify({"error": "Nenhum ficheiro enviado"}), 400
    return jsonify(upload_and_store(request.files['file'], 'update'))

@app.route('/api/get_common_columns', methods=['POST'])
def get_common_columns():
    data = request.json
    base_sheet, update_sheet = data.get('baseSheet'), data.get('updateSheet')
    base_fp, update_fp = session.get('base_filepath'), session.get('update_filepath')
    df_base = pd.read_excel(base_fp, sheet_name=base_sheet)
    df_update = pd.read_excel(update_fp, sheet_name=update_sheet)
    return jsonify({"commonColumns": list(set(df_base.columns) & set(df_update.columns))})

@app.route('/api/process_multi_tab_update', methods=['POST'])
def process_multi_tab_update():
    instructions = request.json.get('instructions')
    base_fp, update_fp = session.get('base_filepath'), session.get('update_filepath')
    if not all([instructions, base_fp, update_fp]):
        return jsonify({"error": "Instruções de fusão ou ficheiros em falta."}), 400
    
    try:
        dfs_base = pd.read_excel(base_fp, sheet_name=None)
        dfs_update = pd.read_excel(update_fp, sheet_name=None)
        
        for inst in instructions:
            b_tab, u_tab, key = inst['baseTab'], inst['updateTab'], inst['keyColumn']
            if b_tab in dfs_base and u_tab in dfs_update:
                df_b = dfs_base[b_tab]
                df_u = dfs_update[u_tab]
                if key in df_b.columns and key in df_u.columns:
                    # **LÓGICA DE FUSÃO CORRIGIDA AQUI**
                    # 1. Definir o índice em ambos os dataframes
                    df_b = df_b.set_index(key)
                    df_u = df_u.set_index(key)
                    
                    # 2. Filtrar o dataframe base para manter apenas as chaves que existem no de atualização
                    # Isto remove as linhas do base que foram excluídas no de atualização.
                    df_b_filtered = df_b[df_b.index.isin(df_u.index)]
                    
                    # 3. Atualizar os valores no dataframe base filtrado com os valores do de atualização
                    df_b_filtered.update(df_u)
                    
                    # 4. Substituir o dataframe antigo pelo novo, já processado
                    dfs_base[b_tab] = df_b_filtered.reset_index()

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for name, df in dfs_base.items():
                df.to_excel(writer, sheet_name=name, index=False)
        output.seek(0)
        return send_file(output, as_attachment=True, download_name=f"atualizado_{session.get('base_filename', 'b.xlsx')}", mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as e:
        return jsonify({'error': f'Erro durante o processamento: {e}'}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)