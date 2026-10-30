from flask import Flask, jsonify, request
from flask_cors import CORS
import openpyxl

app = Flask(__name__)
CORS(app)

NOME_ARQUIVO = "faturas_energia.xlsx"

@app.route('/api/dados', methods=['GET'])
def obter_dados():
    try:
        wb = openpyxl.load_workbook(NOME_ARQUIVO, data_only=True)
        
        # Filtra abas indesejadas (ignorando 'TOTAL' e 'STATUS')
        abas_ignoradas = ['TOTAL', 'STATUS']
        abas_disponiveis = [sheet for sheet in wb.sheetnames if sheet.upper().strip() not in abas_ignoradas]
        
        if not abas_disponiveis:
            return jsonify({'status': 'erro', 'mensagem': 'Nenhuma aba válida encontrada na planilha.'}), 400

        # Cálculo da soma de todos os valores provisionados de todas as abas (Linha 4, Coluna D)
        total_provisionado_geral = 0.0
        for aba_nome in abas_disponiveis:
            sheet_temp = wb[aba_nome]
            val = sheet_temp.cell(row=4, column=4).value
            if val is not None and isinstance(val, (int, float)):
                total_provisionado_geral += float(val)

        # Mês selecionado
        mes_solicitado = request.args.get('mes')
        if mes_solicitado and mes_solicitado in abas_disponiveis:
            nome_aba = mes_solicitado
        else:
            nome_aba = abas_disponiveis[0]

        sheet = wb[nome_aba]

        # Resumos das colunas superiores (Linha 4 do Excel)
        v_prov = sheet.cell(row=4, column=4).value
        v_out  = sheet.cell(row=4, column=6).value
        v_rev  = sheet.cell(row=4, column=8).value

        val_provisionado = float(v_prov) if isinstance(v_prov, (int, float)) else 0.0
        val_outubro      = float(v_out)  if isinstance(v_out,  (int, float)) else 0.0
        val_reversao     = float(v_rev)  if isinstance(v_rev,  (int, float)) else 0.0

        lista_faturas = []
        total_bruto = 0.0
        total_liquido = 0.0

        # Percorre as linhas a partir da linha 4
        for row_idx in range(4, sheet.max_row + 1):
            bruto   = sheet.cell(row=row_idx, column=1).value
            liquido = sheet.cell(row=row_idx, column=2).value
            raw_nota = sheet.cell(row=row_idx, column=3).value

            # Formata a nota fiscal
            if raw_nota is not None:
                if isinstance(raw_nota, float):
                    nota = str(int(raw_nota)).strip()
                else:
                    nota = str(raw_nota).replace('.0', '').strip()
            else:
                nota = ""

            # Para se for linha TOTAL ou vazia
            if "TOTAL" in str(nota).upper() or bruto is None or bruto == "-":
                break

            try:
                v_bruto = float(bruto)
                v_liq   = float(liquido) if liquido is not None and liquido != "-" else 0.0

                lista_faturas.append({
                    'nota': nota,
                    'bruto': v_bruto,
                    'liquido': v_liq
                })

                total_bruto += v_bruto
                total_liquido += v_liq
            except (ValueError, TypeError):
                continue

        wb.close()

        return jsonify({
            'status': 'sucesso',
            'aba_atual': nome_aba,
            'abas_disponiveis': abas_disponiveis,
            'kpis': {
                'total_bruto': total_bruto,
                'total_liquido': total_liquido,
                'provisionado': val_provisionado,
                'provisionado_geral': total_provisionado_geral,
                'outubro_total': val_outubro,
                'reversao': val_reversao
            },
            'faturas': lista_faturas
        })

    except Exception as e:
        return jsonify({'status': 'erro', 'mensagem': str(e)}), 500

if __name__ == '__main__':
    app.run(port=5000)