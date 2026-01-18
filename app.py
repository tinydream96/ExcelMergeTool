import os
import uuid
import pandas as pd
from flask import Flask, render_template, request, jsonify, send_file

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def get_excel_sheets(file_path):
    if file_path.endswith('.csv'):
        return ['CSV Default']
    try:
        xl = pd.ExcelFile(file_path)
        return xl.sheet_names
    except Exception as e:
        return [str(e)]

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    file = request.files['file']
    ext = os.path.splitext(file.filename)[1]
    temp_filename = f"{uuid.uuid4()}{ext}"
    temp_path = os.path.join(UPLOAD_FOLDER, temp_filename)
    file.save(temp_path)
    
    sheets = get_excel_sheets(temp_path)
    return jsonify({
        'filename': file.filename,
        'temp_path': temp_path,
        'sheets': sheets
    })

@app.route('/get_columns', methods=['POST'])
def get_columns():
    data = request.json
    try:
        # Load Target Headers
        if data['target_file'].endswith('.csv'):
            df_t = pd.read_csv(data['target_file'], nrows=0)
        else:
            df_t = pd.read_excel(data['target_file'], sheet_name=data['target_sheet'], nrows=0)
        
        # Load Source Headers
        if data['source_file'].endswith('.csv'):
            df_s = pd.read_csv(data['source_file'], nrows=0)
        else:
            df_s = pd.read_excel(data['source_file'], sheet_name=data['source_sheet'], nrows=0)
            
        return jsonify({
            'target_columns': df_t.columns.tolist(),
            'source_columns': df_s.columns.tolist()
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/preview', methods=['POST'])
def preview():
    data = request.json
    try:
        # Load ID columns only for performance
        if data['target_file'].endswith('.csv'):
            df_t = pd.read_csv(data['target_file'], usecols=[data['target_key']])
        else:
            df_t = pd.read_excel(data['target_file'], sheet_name=data['target_sheet'], usecols=[data['target_key']])
            
        if data['source_file'].endswith('.csv'):
            df_s = pd.read_csv(data['source_file'], usecols=[data['source_key']])
        else:
            df_s = pd.read_excel(data['source_file'], sheet_name=data['source_sheet'], usecols=[data['source_key']])
            
        t_keys = set(df_t[data['target_key']].dropna().astype(str))
        s_keys = set(df_s[data['source_key']].dropna().astype(str))
        
        matches = len(t_keys.intersection(s_keys))
        total = len(t_keys)
        
        return jsonify({
            'total': total,
            'matches': matches,
            'misses': total - matches
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/merge', methods=['POST'])
def merge():
    data = request.json
    try:
        # Load dataframes
        if data['target_file'].endswith('.csv'):
            df_t = pd.read_csv(data['target_file'])
        else:
            df_t = pd.read_excel(data['target_file'], sheet_name=data['target_sheet'])
            
        if data['source_file'].endswith('.csv'):
            df_s = pd.read_csv(data['source_file'])
        else:
            df_s = pd.read_excel(data['source_file'], sheet_name=data['source_sheet'])

        # Prepare mapping
        # Convert keys to string for reliable matching
        df_t['_key_tmp_'] = df_t[data['target_key']].astype(str).str.strip()
        df_s['_key_tmp_'] = df_s[data['source_key']].astype(str).str.strip()
        
        # Apply each column mapping
        for m in data['mapping']:
            src_col = m['src']
            tgt_col = m['tgt']
            
            # Create map from source
            mapping_dict = df_s.set_index('_key_tmp_')[src_col].to_dict()
            
            # Map values to target, only update if key exists in mapping_dict
            def update_val(row):
                key = row['_key_tmp_']
                if key in mapping_dict:
                    return mapping_dict[key]
                return row[tgt_col]
            
            df_t[tgt_col] = df_t.apply(update_val, axis=1)

        # Drop temp key
        df_t = df_t.drop(columns=['_key_tmp_'])
        
        output_path = os.path.join(UPLOAD_FOLDER, f"result_{uuid.uuid4()}.xlsx")
        df_t.to_excel(output_path, index=False)
        
        return send_file(output_path, as_attachment=True)
    except Exception as e:
        print(e)
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':

    # Open browser automatically
    import webbrowser
    from threading import Timer
    
    def open_browser():
        webbrowser.open_new("http://127.0.0.1:5001")

    Timer(1.5, open_browser).start()
    app.run(debug=False, port=5001)
