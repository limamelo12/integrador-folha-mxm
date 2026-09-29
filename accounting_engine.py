import json
import os
import re
import openpyxl
from datetime import datetime

def load_mapping(config_path):
    """Loads the De-Para mapping from a JSON file."""
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_mapping(mapping, config_path):
    """Saves the De-Para mapping to a JSON file."""
    try:
        os.makedirs(os.path.dirname(config_path), exist_ok=True)
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(mapping, f, ensure_ascii=False, indent=4)
        return True
    except Exception as e:
        print(f"Error saving mapping: {str(e)}")
        return False

def generate_entries(parsed_data, mapping, batch_number="1", entry_date=None, doc_number="FOLHA", process_type=None, *args, **kwargs):
    """
    Generates accounting entry rows based on parsed events and user mapping.
    
    mapping structure:
    {
        "event_code": {
            "debit_account": str,
            "credit_account": str,
            "cost_center": str
        }
    }
    """
    events = list(parsed_data.get('events') or [])
    meta = parsed_data.get('metadata') or {}
    empresa_code = meta.get('empresa_codigo', '0001')
    
    # Process social charges as virtual events if present
    sc = parsed_data.get('social_charges') or {}
    if sc:
        func_val = float(sc.get('gps_empresa_func') or 0.0)
        socios_val = float(sc.get('gps_empresa_socios') or 0.0)
        auton_val = float(sc.get('gps_empresa_auton') or 0.0)
        rat_val = float(sc.get('gps_rat') or 0.0)
        terc_val = float(sc.get('gps_terceiros') or 0.0)
        fgts_val = float(sc.get('fgts_total') or 0.0)

        patronal_total = func_val + socios_val + auton_val
        gps_total = patronal_total + rat_val + terc_val
        
        # If mapping specifically maps consolidated GPS instead of detailed GPS_PATRONAL:
        if 'GPS' in mapping and 'GPS_PATRONAL' not in mapping:
            if abs(gps_total) >= 0.01:
                events.append({
                    'code': 'GPS',
                    'description': 'INSS',
                    'total': gps_total,
                    'section': 'social_charges',
                    'type': 'encargo'
                })
        else:
            if abs(patronal_total) >= 0.01:
                events.append({
                    'code': 'GPS_PATRONAL',
                    'description': 'INSS Patronal (Empresa)',
                    'total': patronal_total,
                    'section': 'social_charges',
                    'type': 'encargo'
                })
            if abs(rat_val) >= 0.01:
                events.append({
                    'code': 'GPS_RAT',
                    'description': 'INSS RAT/FAP',
                    'total': rat_val,
                    'section': 'social_charges',
                    'type': 'encargo'
                })
            if abs(terc_val) >= 0.01:
                events.append({
                    'code': 'GPS_TERCEIROS',
                    'description': 'INSS Terceiros',
                    'total': terc_val,
                    'section': 'social_charges',
                    'type': 'encargo'
                })
        if abs(fgts_val) >= 0.01:
            events.append({
                'code': 'FGTS',
                'description': 'FGTS',
                'total': fgts_val,
                'section': 'social_charges',
                'type': 'encargo'
            })
    
    # Normalizar código da empresa (0001 -> SVP, 0002 -> 2DIR / Holding, 0003 -> MCI)
    raw_code = str(empresa_code or meta.get('empresa_codigo', '0001')).strip()
    emp_nome_upper = str(meta.get('empresa_nome', '')).upper()
    cnpj_clean = re.sub(r"\D", "", str(meta.get('cnpj', '')))

    if "JASPER" in emp_nome_upper or "PARTICIPA" in emp_nome_upper or "HOLDING" in emp_nome_upper or cnpj_clean.startswith("09436824") or raw_code in ('2', '0002') or '2000' in raw_code:
        clean_empresa = '0002'
    elif '3000' in raw_code or raw_code in ('3', '0003') or 'INDUSTRIA' in emp_nome_upper:
        clean_empresa = '0003'
    elif '1000' in raw_code or raw_code in ('1', '0001') or 'VIA PARQUE' in emp_nome_upper:
        clean_empresa = '0001'
    else:
        clean_empresa = raw_code.zfill(4)

    # Definir centro de custo padrão fixo por empresa (0001 -> 1SVP, 0002 -> 2DIR, 0003 -> 3FAB)
    STANDARD_COMPANY_CC = {
        "0001": "1SVP",
        "0002": "2DIR",
        "0003": "3FAB"
    }
    COMPANY_HIST_TAG = {
        "0001": "1SVP",
        "0002": "2DIR",
        "0003": "MCI"
    }
    company_fixed_cc = STANDARD_COMPANY_CC.get(clean_empresa, "1SVP")
    hist_company_tag = COMPANY_HIST_TAG.get(clean_empresa, "1SVP")
        
    # Format month/year for history (e.g. 05/2026 -> MAI/26)
    mes_ano = ""
    mes_digits = ""
    if meta.get('periodo_referencia'):
        try:
            parts = meta['periodo_referencia'].split('/')
            if len(parts) == 2:
                mes, ano = parts[0], parts[1]
                mes_digits = mes.zfill(2)
                meses_pt = {
                    "01": "JAN", "02": "FEV", "03": "MAR", "04": "ABR",
                    "05": "MAI", "06": "JUN", "07": "JUL", "08": "AGO",
                    "09": "SET", "10": "OUT", "11": "NOV", "12": "DEZ"
                }
                mes_abrev = meses_pt.get(mes, mes)
                ano_abrev = ano[-2:]
                mes_ano = f"{mes_abrev}/{ano_abrev}"
        except Exception:
            pass
    if not mes_ano:
        mes_ano = datetime.now().strftime("%b/%y").upper()
        mes_digits = datetime.now().strftime("%m")
        
    # Determinar tipo de processo (folha, rescisao, ferias)
    if not process_type:
        process_type = meta.get('tipo_processo', 'folha')
    proc_clean = str(process_type).lower().strip()
    
    if 'rescis' in proc_clean:
        hist_proc_tag = "RESC"
        default_batch_prefix = "RESC"
    elif 'feria' in proc_clean:
        hist_proc_tag = "FERIAS"
        default_batch_prefix = "FER"
    else:
        hist_proc_tag = "FOLHA"
        default_batch_prefix = "FOLH"

    # Default batch and doc number if generic
    if batch_number in ["1", "", None] or str(batch_number).startswith("FOLH") or str(batch_number).startswith("RESC"):
        batch_number = f"{default_batch_prefix}{mes_digits}" if mes_digits else default_batch_prefix
    if doc_number in ["FOLHA", "", None] or str(doc_number).startswith("FOLH") or str(doc_number).startswith("RESC"):
        doc_number = f"{default_batch_prefix}{mes_digits}" if mes_digits else hist_proc_tag

    if not entry_date:
        # Default to the end date of the period or today
        if meta.get('periodo_fim'):
            # Assume DD/MM/YYYY
            try:
                date_obj = datetime.strptime(meta['periodo_fim'], "%d/%m/%Y")
                entry_date = date_obj.strftime("%d%m%Y")
            except ValueError:
                entry_date = datetime.now().strftime("%d%m%Y")
        else:
            entry_date = datetime.now().strftime("%d%m%Y")
    else:
        # Standardize date format to DDMMAAAA
        entry_date = entry_date.replace("/", "").replace("-", "")
        
    rows = []
    sequence = 1
    unmapped_events = []
    
    total_debit = 0.0
    total_credit = 0.0
    
    # Process each event
    for ev in events:
        code = str(ev.get('code', '')).strip()
        desc = str(ev.get('description', '')).strip()
        try:
            amount = float(ev.get('total') or 0.0)
        except (ValueError, TypeError):
            amount = 0.0
        
        # Skip events with zero value
        if abs(amount) < 0.01:
            continue
            
        ev_map = mapping.get(code, {})
        debit_acc = ev_map.get('debit_account', '').strip()
        credit_acc = ev_map.get('credit_account', '').strip()
        cc = ev_map.get('cost_center', '').strip()
        
        # Track unmapped events
        if not debit_acc or not credit_acc:
            unmapped_events.append({
                'code': code,
                'description': desc,
                'amount': amount,
                'debit_missing': not debit_acc,
                'credit_missing': not credit_acc
            })
            
        # Format history according to company historical pattern
        # For social charges in Empresa 0003: "INSS" or "FGTS"
        if str(code).startswith('GPS_'):
            event_name_hist = "INSS" if clean_empresa == "0003" else desc.upper().strip()
        elif str(code) == 'FGTS':
            event_name_hist = "FGTS"
        else:
            event_name_hist = desc.upper().strip()

        # Check if custom history is defined in mapping or specific to Empresa 0002 (Holding)
        custom_hist = str(ev_map.get('historico', '')).strip()
        if custom_hist:
            hist_entry = custom_hist[:200]
        elif clean_empresa == "0002":
            if "INSS" in desc.upper() or str(code).startswith('GPS_'):
                hist_entry = "PELO VALOR DE INSS S/ PRO LABORE CONF FOLHA"
            elif "IRRF" in desc.upper():
                hist_entry = "PELO VALOR DE IRRF S/ PRO LABORE CONF FOLHA"
            else:
                hist_entry = "PELO VALOR DE PRO LABORE CONF FOLHA"
        else:
            # Pattern: <EVENTO>-<PROC> <TAG> - <MES>/<ANO> (ex: SALDO DE SALÁRIO-RESC MCI - AGO/26 ou SALÁRIO BASE-FOLHA MCI - AGO/26)
            hist_entry = f"{event_name_hist}-{hist_proc_tag} {hist_company_tag} - {mes_ano}".upper()[:200]
            
        if clean_empresa == "0002":
            num_titulo = ""
        else:
            num_titulo = f"{hist_proc_tag} {hist_company_tag} - {mes_ano}".upper()[:50]
            
        # Leg 1: Debit Row
        if debit_acc:
            clean_debit = debit_acc.replace(".", "").strip()
            use_cc_debit = (company_fixed_cc if company_fixed_cc else cc) if (clean_debit.startswith('3') or clean_debit.startswith('4') or clean_debit.startswith('5')) else None
            
            rows.append({
                'empresa': clean_empresa,
                'lote': batch_number,
                'data': entry_date,
                'documento': doc_number,
                'conta': debit_acc,
                'cc': use_cc_debit if use_cc_debit else None,
                'tipo': 'D',
                'historico': hist_entry,
                'valor': amount,
                'sequencia': sequence,
                'numero_titulo': num_titulo
            })
            total_debit += amount
            sequence += 1
            
        # Leg 2: Credit Row
        if credit_acc:
            clean_credit = credit_acc.replace(".", "").strip()
            use_cc_credit = (company_fixed_cc if company_fixed_cc else cc) if (clean_credit.startswith('3') or clean_credit.startswith('4') or clean_credit.startswith('5')) else None
            
            rows.append({
                'empresa': clean_empresa,
                'lote': batch_number,
                'data': entry_date,
                'documento': doc_number,
                'conta': credit_acc,
                'cc': use_cc_credit if use_cc_credit else None,
                'tipo': 'C',
                'historico': hist_entry,
                'valor': amount,
                'sequencia': sequence,
                'numero_titulo': num_titulo
            })
            total_credit += amount
            sequence += 1
            
    is_balanced = abs(total_debit - total_credit) < 0.01
    
    return {
        'rows': rows,
        'total_debit': total_debit,
        'total_credit': total_credit,
        'difference': abs(total_debit - total_credit),
        'is_balanced': is_balanced,
        'unmapped_events': unmapped_events
    }

def write_to_excel_template(template_path, output_path, entries_data):
    """
    Loads the MXM template Excel file, populates the 'Dados' sheet with generated entries,
    and saves the workbook to output_path.
    """
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template Excel file not found: {template_path}")
        
    wb = openpyxl.load_workbook(template_path)
    
    # Ensure sheet 'Dados' or 'Preencher' exists
    if 'Dados' in wb.sheetnames:
        sheet = wb['Dados']
    elif 'Preencher' in wb.sheetnames:
        sheet = wb['Preencher']
    else:
        sheet = wb.active
    
    # Clear existing data rows under header
    # Header is on row 1. Let's delete all rows from row 2 onwards.
    if sheet.max_row > 1:
        sheet.delete_rows(2, sheet.max_row)
        
    # Columns mapping in the sheet 'Dados' based on the template:
    # A (1): Codigo da empresa
    # B (2): Lote
    # C (3): Data do lancamento
    # D (4): Documento
    # E (5): Conta contabil
    # F (6): Centro de Custo
    # G (7): Indicador de Debito ou Credito
    # H (8): Historico do Lancamento
    # I (9): Blank/Unnamed: 8
    # J (10): Valor do Lancamento na 2a Moeda (blank)
    # K (11): Codigo da Segunda Moeda (blank)
    # L (12): Valor do Lancamento (numeric)
    # M (13): Sequencia do Lancamento (numeric)
    # N (14): Codigo do projeto (blank)
    # O (15): Codigo do Fornecedor (blank)
    # P (16): Codigo do Cliente (blank)
    # Q (17): Numero do Titulo (blank)
    # R (18): Converter Moeda (blank)
    # S (19): Historico Padrao (blank)
    # T (20): Hist. Padrao - Complemento 1 (blank)
    # U (21): Hist. Padrao - Complemento 2 (blank)
    # V (22): Hist. Padrao - Complemento 3 (blank)
    # W (23): Excluir Lancamentos (blank)

    for entry in entries_data:
        row_idx = sheet.max_row + 1
        
        emp_val = str(entry['empresa']).zfill(4)
        sheet.cell(row=row_idx, column=1, value=emp_val)
        sheet.cell(row=row_idx, column=2, value=str(entry['lote'])[:6])
        sheet.cell(row=row_idx, column=3, value=str(entry['data'])[:8])
        sheet.cell(row=row_idx, column=4, value=str(entry['documento'])[:6] if entry['documento'] else "")
        sheet.cell(row=row_idx, column=5, value=str(entry['conta'])[:15])
        sheet.cell(row=row_idx, column=6, value=str(entry['cc'])[:15] if entry['cc'] else "")
        sheet.cell(row=row_idx, column=7, value=str(entry['tipo'])[:1])
        sheet.cell(row=row_idx, column=8, value=str(entry['historico'])[:200])
        
        # Column I (9) is blank
        sheet.cell(row=row_idx, column=9, value="")
        # Column J (10) 2nd currency is blank
        sheet.cell(row=row_idx, column=10, value="")
        # Column K (11) 2nd currency code is blank
        sheet.cell(row=row_idx, column=11, value="")
        
        # Column L (12) is Valor do Lancamento
        sheet.cell(row=row_idx, column=12, value=float(entry['valor']))
        
        # Column M (13) is Sequencia
        sheet.cell(row=row_idx, column=13, value=int(entry['sequencia']))
        
        # Fill non-title columns as blank
        for c in [14, 15, 16]:
            sheet.cell(row=row_idx, column=c, value="")
            
        # Column Q (17) is Numero do Titulo
        sheet.cell(row=row_idx, column=17, value=str(entry.get('numero_titulo', ''))[:50])
        
        # Remaining columns (18 to 23) as blank
        for c in range(18, 24):
            sheet.cell(row=row_idx, column=c, value="")
            
    wb.save(output_path)
    return True
