from flask import Flask, jsonify, request
from flask_cors import CORS
import pandas as pd
import openpyxl

app = Flask(__name__)
CORS(app)  # Permite que o HTML consulte o servidor local

NOME_ARQUIVO = "faturas_energia.xlsx"

@app.route('/api/dados', methods=['GET'])
def obter_dados():
    try:
        xls = pd.ExcelFile(NOME_ARQUIVO)
        
        # Filtra todas as abas que não sejam a 'TOTAL'
        abas_disponiveis = [sheet for sheet in xls.sheet_names if sheet.upper().strip() != 'TOTAL']
        
        if not abas_disponiveis:
            return jsonify({'status': 'erro', 'mensagem': 'Nenhuma aba válida encontrada na planilha.'}), 400

        # --- NOVO: Cálculo da soma de todos os valores provisionados de todas as abas ---
        total_provisionado_geral = 0.0
        for aba in abas_disponiveis:
            try:
                # Leitura rápida de cada aba para obter o valor na célula correspondente (Linha 4, Coluna D / índice 3)
                df_temp = pd.read_excel(NOME_ARQUIVO, sheet_name=aba, skiprows=2, nrows=1)
                val = float(df_temp.iloc[0, 3]) if not df_temp.empty and pd.notnull(df_temp.iloc[0, 3]) else 0.0
                total_provisionado_geral += val
            except Exception:
                continue
        # ----------------------------------------------------------------------------------

        # Obtém o mês selecionado via parâmetro da requisição (ex: ?mes=JULHO)
        mes_solicitado = request.args.get('mes')
        
        if mes_solicitado and mes_solicitado in abas_disponiveis:
            nome_aba = mes_solicitado
        else:
            nome_aba = abas_disponiveis[0]

        # Lê os dados ignorando o cabeçalho superior
        df_aba = pd.read_excel(NOME_ARQUIVO, sheet_name=nome_aba, skiprows=2)
        
        # Leitura dos resumos das colunas superiores (Linha 4 do Excel)
        val_provisionado = float(df_aba.iloc[0, 3]) if pd.notnull(df_aba.iloc[0, 3]) else 0.0
        val_outubro      = float(df_aba.iloc[0, 5]) if pd.notnull(df_aba.iloc[0, 5]) else 0.0
        val_reversao     = float(df_aba.iloc[0, 7]) if pd.notnull(df_aba.iloc[0, 7]) else 0.0
        
        lista_faturas = []
        total_bruto = 0.0
        total_liquido = 0.0
        
        # Percorre as notas fiscais
        for _, row in df_aba.iterrows():
            bruto = row.iloc[0]
            liquido = row.iloc[1]
            raw_nota = row.iloc[2]
            
            # Formata o número da nota para remover o .0
            if pd.notnull(raw_nota):
                if isinstance(raw_nota, float):
                    nota = str(int(raw_nota)).strip()
                else:
                    nota = str(raw_nota).replace('.0', '').strip()
            else:
                nota = ""
            
            if "TOTAL" in nota.upper() or pd.isna(bruto) or bruto == "-":
                break
                
            try:
                v_bruto = float(bruto)
                v_liq = float(liquido) if pd.notnull(liquido) and liquido != "-" else 0.0
                
                lista_faturas.append({
                    'nota': nota,
                    'bruto': v_bruto,
                    'liquido': v_liq
                })
                
                total_bruto += v_bruto
                total_liquido += v_liq
            except ValueError:
                continue
                
        return jsonify({
            'status': 'sucesso',
            'aba_atual': nome_aba,
            'abas_disponiveis': abas_disponiveis,
            'kpis': {
                'total_bruto': total_bruto,
                'total_liquido': total_liquido,
                'provisionado': val_provisionado,
                'provisionado_geral': total_provisionado_geral,  # <-- Retornado para o frontend
                'outubro_total': val_outubro,
                'reversao': val_reversao
            },
            'faturas': lista_faturas
        })
    except Exception as e:
        return jsonify({'status': 'erro', 'mensagem': str(e)}), 500

if __name__ == '__main__':
    app.run(port=5000)