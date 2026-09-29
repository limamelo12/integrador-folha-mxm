import re
import pdfplumber

def clean_value(val_str):
    """Converts a Brazilian format number string (e.g. 1.233.265,90 or -1.200,00) to float."""
    if not val_str:
        return 0.0
    val_str = val_str.strip()
    val_str = val_str.replace(".", "").replace(",", ".")
    try:
        return float(val_str)
    except ValueError:
        return 0.0

def parse_payroll_pdf(pdf_path):
    """
    Parses payroll PDF summary containing summaries of multiple filiais or a consolidated summary.
    Supports:
      - Format A: Individual filiais with code, desc, value, quantity (e.g. Report.pdf).
      - Format B: 'Resumo Geral Consolidado' (Alterdata) with 4-column breakdown
        (Ativos, Demitidos, Afastados, Total) and 2-column breakdown for Sócios/Autônomos,
        plus page 4 social charges and bases.
    Returns:
        dict: {
            'filiais': {
                '03050': {
                    'metadata': {
                        'empresa_codigo': str,
                        'empresa_nome': str,
                        'cnpj': str,
                        'periodo_inicio': str,
                        'periodo_fim': str,
                        'periodo_referencia': str
                    },
                    'events': list of dicts (each having code, description, type, total, qtd, section),
                    'bases': {
                        'inss': float,
                        'irrf': float,
                        'fgts': float
                    },
                    'social_charges': {
                        'gps_empresa_func': float,
                        'gps_empresa_socios': float,
                        'gps_empresa_auton': float,
                        'gps_rat': float,
                        'gps_terceiros': float,
                        'fgts_total': float
                    }
                },
                ...
            }
        }
    """
    filiais = {}
    current_code = None
    
    # Header regexes
    re_empresa_format_a = re.compile(r"Empresa\s*:\s*(.*?)\s*\(\s*(\d+)\s*\)", re.IGNORECASE)
    re_empresa_format_b = re.compile(r"Empresa\s*:\s*(?:(\d+)\s*[-–]?\s*)?(.*?)(?:\s*\(\s*(\d+)\s*\)|\s+CNPJ|\s*$)", re.IGNORECASE)
    re_empresa_simple = re.compile(r"Empresa\s*:\s*(.*)", re.IGNORECASE)
    re_cnpj = re.compile(r"(?:CNPJ|PAJ|CEI)(?:\/CEI)?\s*:\s*([\w\.\-\/]+)")
    re_periodo = re.compile(r"(?:Ref\.|Período)\s*:\s*([\d\/]+)\s*a\s*([\d\/]+)")
    
    # Event regexes
    # Format A (1 value + 1 quantity): 001 SALARIO BASE 12.345,67 15
    re_event_1col_qtd = re.compile(r"^\s*(\d{3})\s+(.+?)\s+([\d\.\-]+,[\d\-]+)\s+(\d+)\s*$")
    
    # Format B (4 values: Ativos, Demitidos, Afastados, Total): 001 Salário Base 654.511,48 0,00 0,00 654.511,48
    re_event_4col = re.compile(r"^\s*(\d{3})\s+(.+?)\s+([\d\.\-]+,[\d\-]+)\s+([\d\.\-]+,[\d\-]+)\s+([\d\.\-]+,[\d\-]+)\s+([\d\.\-]+,[\d\-]+)\s*$")
    
    # Format B Sócios/Autônomos (2 values: Sócios, Autônomos): 002 REMUNERAÇÃO AUTONOMO 0,00 17.030,61
    re_event_2col = re.compile(r"^\s*(\d{3})\s+(.+?)\s+([\d\.\-]+,[\d\-]+)\s+([\d\.\-]+,[\d\-]+)\s*$")
    
    # Section trackers for Format B
    current_person_type = "funcionarios" # "funcionarios" or "socios"
    current_event_type = "provento"       # "provento" or "desconto"
    
    # Bases regex Format A
    re_bases_a = re.compile(r"Totais das Bases[\s\.]*:\s*([\d\.\-]+,[\d\-]+)\s+([\d\.\-]+,[\d\-]+)\s+([\d\.\-]+,[\d\-]+)", re.IGNORECASE)
    
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            text = page.extract_text()
            if not text:
                continue
                
            lines = text.split("\n")
            for line in lines:
                line_str = line.strip()
                if not line_str:
                    continue
                    
                # Check for company header line
                if line_str.startswith("Empresa :") or line_str.startswith("Empresa:"):
                    m_a = re_empresa_format_a.search(line_str)
                    if m_a:
                        name, code = m_a.groups()
                        code = code.strip()
                        name = name.strip()
                    else:
                        m_b = re_empresa_format_b.search(line_str)
                        if m_b and (m_b.group(1) or m_b.group(3)):
                            g1, g2, g3 = m_b.groups()
                            code = (g1 or g3 or "0001").strip()
                            name = (g2 or "").strip()
                        else:
                            m_simple = re_empresa_simple.search(line_str)
                            if m_simple:
                                name = m_simple.group(1).strip()
                                code = current_code if current_code else "0001"
                            else:
                                code = current_code if current_code else "0001"
                                name = "Empresa"
                                
                    name = re.sub(r"\s+Página\s*:\s*\d+", "", name, flags=re.IGNORECASE).strip()
                    name = re.sub(r"\s+CNPJ\s*:.*", "", name, flags=re.IGNORECASE).strip()
                    
                    if current_code != code:
                        current_code = code
                        # Reset sections on new company
                        current_person_type = "funcionarios"
                        current_event_type = "provento"
                        
                    if current_code not in filiais:
                        filiais[current_code] = {
                            'metadata': {
                                'empresa_codigo': current_code,
                                'empresa_nome': name,
                                'cnpj': '',
                                'periodo_inicio': '',
                                'periodo_fim': '',
                                'periodo_referencia': ''
                            },
                            'events': [],
                            'bases': {
                                'inss': 0.0,
                                'irrf': 0.0,
                                'fgts': 0.0
                            },
                            'social_charges': {
                                'gps_empresa_func': 0.0,
                                'gps_empresa_socios': 0.0,
                                'gps_empresa_auton': 0.0,
                                'gps_rat': 0.0,
                                'gps_terceiros': 0.0,
                                'fgts_total': 0.0
                            }
                        }
                        
                if current_code is None:
                    continue
                    
                # Extract CNPJ
                m_cnpj = re_cnpj.search(line_str)
                if m_cnpj and not filiais[current_code]['metadata']['cnpj']:
                    filiais[current_code]['metadata']['cnpj'] = m_cnpj.group(1).strip()
                    
                # Extract Period
                m_per = re_periodo.search(line_str)
                if m_per and not filiais[current_code]['metadata']['periodo_inicio']:
                    filiais[current_code]['metadata']['periodo_inicio'] = m_per.group(1).strip()
                    filiais[current_code]['metadata']['periodo_fim'] = m_per.group(2).strip()
                    parts = m_per.group(1).split('/')
                    if len(parts) == 3:
                        filiais[current_code]['metadata']['periodo_referencia'] = f"{parts[1]}/{parts[2]}"
                        
                # Section triggers (Format B)
                if "Valores pagos aos Funcionários" in line_str or "Valores pagos aos Funcionarios" in line_str:
                    current_person_type = "funcionarios"
                elif "Valores pagos aos Sócios" in line_str or "Valores pagos aos Socios" in line_str:
                    current_person_type = "socios"
                    current_event_type = "provento"
                elif "TOTAL DE ADICIONAIS" in line_str:
                    current_event_type = "desconto"
                elif "TOTAL DE DESCONTOS" in line_str:
                    current_event_type = "pos_desconto"
                    
                # Match event lines - Format B (4 monetary columns)
                m_4col = re_event_4col.match(line_str)
                if m_4col and not line_str.startswith("TOTAL") and not line_str.startswith("ADICIONAIS"):
                    ev_code, ev_desc, val_atv, val_dem, val_afa, val_tot = m_4col.groups()
                    tot_num = clean_value(val_tot)
                    filiais[current_code]['events'].append({
                        'code': ev_code,
                        'description': ev_desc.strip(),
                        'type': current_event_type,
                        'total': tot_num,
                        'qtd': 0,
                        'section': 'funcionarios'
                    })
                    continue
                    
                # Match event lines - Format B Sócios/Autônomos (2 monetary columns)
                m_2col = re_event_2col.match(line_str)
                if m_2col and current_person_type == "socios" and not line_str.startswith("TOTAL") and not line_str.startswith("ADICIONAIS"):
                    ev_code, ev_desc, val_soc, val_aut = m_2col.groups()
                    tot_num = clean_value(val_soc) + clean_value(val_aut)
                    filiais[current_code]['events'].append({
                        'code': ev_code,
                        'description': ev_desc.strip(),
                        'type': current_event_type,
                        'total': tot_num,
                        'qtd': 0,
                        'section': 'socios_autonomos'
                    })
                    continue
                    
                # Match event lines - Format A (1 value + 1 quantity)
                m_1col = re_event_1col_qtd.match(line_str)
                if m_1col and not line_str.startswith("TOTAL"):
                    ev_code, ev_desc, val_str, qtd_str = m_1col.groups()
                    ev_code_int = int(ev_code)
                    ev_type = "provento" if (ev_code_int < 600 or ev_code_int == 998) else "desconto"
                    filiais[current_code]['events'].append({
                        'code': ev_code,
                        'description': ev_desc.strip(),
                        'type': ev_type,
                        'total': clean_value(val_str),
                        'qtd': int(qtd_str),
                        'section': 'funcionarios'
                    })
                    continue
                    
                # Match bases - Format A
                m_bases = re_bases_a.search(line_str)
                if m_bases:
                    inss, irrf, fgts = m_bases.groups()
                    filiais[current_code]['bases']['inss'] = clean_value(inss)
                    filiais[current_code]['bases']['irrf'] = clean_value(irrf)
                    filiais[current_code]['bases']['fgts'] = clean_value(fgts)
                    continue
                    
                # Match charges & bases - Format B (Page 4)
                if "Empresa Funcionários:" in line_str or "Empresa Funcionarios:" in line_str:
                    m = re.search(r"Empresa Funcion[áa]rios:\s*([\d\.\-]+,[\d\-]+)", line_str, re.IGNORECASE)
                    if m:
                        filiais[current_code]['social_charges']['gps_empresa_func'] = clean_value(m.group(1))
                        
                if "Empresa Autônomos:" in line_str or "Empresa Autonomos:" in line_str:
                    m = re.search(r"Empresa Aut[ôo]nomos:\s*([\d\.\-]+,[\d\-]+)", line_str, re.IGNORECASE)
                    if m:
                        filiais[current_code]['social_charges']['gps_empresa_auton'] = clean_value(m.group(1))
                        
                if "Empresa Sócios:" in line_str or "Empresa Socios:" in line_str:
                    m = re.search(r"Empresa S[óo]cios:\s*([\d\.\-]+,[\d\-]+)", line_str, re.IGNORECASE)
                    if m:
                        filiais[current_code]['social_charges']['gps_empresa_socios'] = clean_value(m.group(1))
                        
                if "RAT Emp" in line_str:
                    m = re.search(r"RAT Emp.*?:\s*([\d\.\-]+,[\d\-]+)", line_str, re.IGNORECASE)
                    if m:
                        filiais[current_code]['social_charges']['gps_rat'] = clean_value(m.group(1))
                        
                if "Terceiros 5,80 %:" in line_str or "Terceiros 5.80 %:" in line_str or "Terceiros:" in line_str:
                    m = re.search(r"Terceiros(?:\s+[\d\.\,]+\s*%)?:\s*([\d\.\-]+,[\d\-]+)", line_str, re.IGNORECASE)
                    if m:
                        filiais[current_code]['social_charges']['gps_terceiros'] = clean_value(m.group(1))
                        
                if "Total FGTS apurado" in line_str:
                    m = re.search(r"Total FGTS apurado[^\:]*:\s*([\d\.\-]+,[\d\-]+)", line_str, re.IGNORECASE)
                    if m:
                        filiais[current_code]['social_charges']['fgts_total'] = clean_value(m.group(1))
                        
                # Bases - Format B
                if "Base Empregados:" in line_str:
                    m = re.search(r"Base Empregados:\s*([\d\.\-]+,[\d\-]+)", line_str, re.IGNORECASE)
                    if m:
                        filiais[current_code]['bases']['inss'] += clean_value(m.group(1))
                        
                if "Base Autônomos:" in line_str or "Base Autonomos:" in line_str:
                    m = re.search(r"Base Aut[ôo]nomos:\s*([\d\.\-]+,[\d\-]+)", line_str, re.IGNORECASE)
                    if m:
                        filiais[current_code]['bases']['inss'] += clean_value(m.group(1))
                        
                if "Base Sócios:" in line_str or "Base Socios:" in line_str:
                    m = re.search(r"Base S[óo]cios:\s*([\d\.\-]+,[\d\-]+)", line_str, re.IGNORECASE)
                    if m:
                        filiais[current_code]['bases']['inss'] += clean_value(m.group(1))
                        
                if "Base de calc. FGTS sem 13" in line_str:
                    m = re.search(r"Base de calc\. FGTS sem 13[^\:]*:\s*([\d\.\-]+,[\d\-]+)", line_str, re.IGNORECASE)
                    if m:
                        filiais[current_code]['bases']['fgts'] += clean_value(m.group(1))
                        
                if "Base de calc. FGTS 13" in line_str:
                    m = re.search(r"Base de calc\. FGTS 13[^\:]*:\s*([\d\.\-]+,[\d\-]+)", line_str, re.IGNORECASE)
                    if m:
                        filiais[current_code]['bases']['fgts'] += clean_value(m.group(1))
                        
    return {
        'filiais': filiais
    }
