import pandas as pd
import json
import os

def process_excel(file_path, output_json='data.json'):
    print(f"Lendo planilha: {file_path}...")
    
    try:
        df = pd.read_excel(file_path, sheet_name='Worksheet', header=1)
        
        # --- 1. SEPARAÇÃO DE MÉTRICAS ---
        import unicodedata
        import re

        def normalize_key(val):
            s = str(val or '')
            s = unicodedata.normalize('NFD', s)
            s = re.sub(r'[\u0300-\u036f]', '', s)
            s = s.replace('\u00A0', ' ')
            s = re.sub(r'[\r\n]+', ' ', s)
            s = re.sub(r'\s*/\s*', '/', s)
            s = re.sub(r'\s+', ' ', s)
            return s.strip().lower()

        def find_col(candidates):
            for c in candidates:
                if c in df.columns:
                    return c
            cols_norm = {normalize_key(col): col for col in df.columns}
            for c in candidates:
                n = normalize_key(c)
                if n in cols_norm:
                    return cols_norm[n]
            return None

        col_alunos = find_col(['Alunos', 'Total de Alunos', 'Total Alunos'])
        col_aptos = find_col([
            'Alunos Aptos/Agendados',
            'Alunos Aptos / Agendados',
            'Aptos/Agendados',
            'Aptos / Agendados',
            'Alunos Aptos',
            'Aptos'
        ])
        col_inaptos = find_col([
            'Alunos Inaptos/Não Agendados',
            'Alunos Inaptos / Não Agendados',
            'Alunos Inaptos/Nao Agendados',
            'Alunos Inaptos / Nao Agendados',
            'Inaptos/Não Agendados',
            'Inaptos / Não Agendados',
            'Inaptos/Nao Agendados',
            'Inaptos / Nao Agendados',
            'Alunos Inaptos',
            'Inaptos'
        ])
        col_presentes = find_col(['Presentes', 'Presente'])
        col_ausentes = find_col(['Ausentes', 'Ausente'])

        df['_total_alunos'] = pd.to_numeric(df[col_alunos], errors='coerce').fillna(0).astype(int) if col_alunos else 0
        df['_alunos_aptos'] = pd.to_numeric(df[col_aptos], errors='coerce').fillna(0).astype(int) if col_aptos else 0
        df['_alunos_inaptos'] = pd.to_numeric(df[col_inaptos], errors='coerce').fillna(0).astype(int) if col_inaptos else 0
        df['_presentes'] = pd.to_numeric(df[col_presentes], errors='coerce').fillna(0).astype(int) if col_presentes else 0
        df['_ausentes'] = pd.to_numeric(df[col_ausentes], errors='coerce').fillna(0).astype(int) if col_ausentes else 0

        # --- 2. LIMPEZA E EXPORTAÇÃO ---
        
        if 'Categoria' in df.columns:
            df['Categoria'] = df['Categoria'].fillna('-').astype(str)
        else:
            df['Categoria'] = '-'

        cols_to_keep = [
            'Unidade operacional (Escola)',
            'Curso',
            'Modalidade',
            'Categoria',
            '_total_alunos',
            '_alunos_aptos',
            '_alunos_inaptos',
            '_presentes',
            '_ausentes'
        ]
        
        data = df[cols_to_keep].to_dict(orient='records')
        
        with open(output_json, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
            
        print(f"Sucesso! {len(data)} registros processados.")
        
    except Exception as e:
        print(f"Erro crítico: {e}")

if __name__ == "__main__":
    FILE = 'PlanilhaAvaliaçãoObjetiva.xlsx'
    if os.path.exists(FILE):
        process_excel(FILE)
    else:
        print(f"Erro: Arquivo {FILE} não encontrado.")
