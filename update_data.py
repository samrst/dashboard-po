import pandas as pd
import json
import os

def process_excel(file_path, output_json='data.json'):
    print(f"Lendo planilha: {file_path}...")
    
    try:
        df = pd.read_excel(file_path, sheet_name='Worksheet', header=1)
        
        # --- 1. SEPARAÇÃO DE MÉTRICAS ---
        # Mapeamento direto das colunas da planilha modelo
        df['_total_alunos'] = pd.to_numeric(df['Alunos '], errors='coerce').fillna(0).astype(int)
        df['_alunos_aptos'] = pd.to_numeric(df['Alunos Aptos'], errors='coerce').fillna(0).astype(int)
        df['_alunos_inaptos'] = pd.to_numeric(df['Alunos Inaptos'], errors='coerce').fillna(0).astype(int)
        df['_agendados'] = pd.to_numeric(df['Alunos Agendados'], errors='coerce').fillna(0).astype(int)
        df['_nao_agendados'] = pd.to_numeric(df['Alunos Não Agendados'], errors='coerce').fillna(0).astype(int)
        df['_presentes'] = pd.to_numeric(df['Presentes'], errors='coerce').fillna(0).astype(int)
        df['_ausentes'] = pd.to_numeric(df['Ausentes'], errors='coerce').fillna(0).astype(int)

        # --- 2. LIMPEZA E EXPORTAÇÃO ---
        
        cols_to_keep = [
            'Unidade operacional (Escola)',
            'Curso',
            'Modalidade',
            '_total_alunos',
            '_alunos_aptos',
            '_alunos_inaptos',
            '_agendados',
            '_nao_agendados',
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
