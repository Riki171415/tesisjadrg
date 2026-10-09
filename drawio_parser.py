import xml.etree.ElementTree as ET
import urllib.parse
import re

import html

def clean_html(raw_html):
    if not raw_html: return ''
    cleanr = re.compile('<.*?>')
    raw_html = raw_html.replace('&nbsp;', ' ')
    cleantext = re.sub(cleanr, ' ', raw_html)
    cleantext = html.unescape(cleantext)
    return ' '.join(cleantext.split())

def parse_drawio_xml(xml_string):
    root = ET.fromstring(xml_string)
    
    nodes = {}
    edges = []
    
    for cell in root.iter('mxCell'):
        edge = cell.get('edge')
        cid = cell.get('id')
        if not cid: continue
        
        if edge == '1':
            src = cell.get('source')
            tgt = cell.get('target')
            if src and tgt:
                edges.append((src, tgt))
        else:
            val = clean_html(cell.get('value', ''))
            geo = cell.find('mxGeometry')
            y = float(geo.get('y', 0)) if geo is not None else 0
            if val:
                nodes[cid] = {'val': val, 'y': y}

    # Build adjacency list
    adj = {n: [] for n in nodes}
    for src, tgt in edges:
        if src in adj and tgt in adj:
            adj[src].append(tgt)

    decision_tree = {}
    flat_mapping = {}

    # Extract PDC codes
    def extract_pdc(text):
        match = re.search(r'PDC\s+([A-Z0-9\s]+(?:or\s+[A-Z0-9\s]+)*)', text, re.IGNORECASE)
        if match:
            raw = match.group(1).upper()
            return [x.strip() for x in raw.split(' OR ') if x.strip()]
        return []

    # Extract AX codes (could be multiple, e.g., AX 11PEX & 11PFX)
    def extract_ax(text):
        match = re.search(r'AX\s+([A-Z0-9\s&]+)', text)
        if match:
            raw = match.group(1).replace(' ', '')
            return [x for x in raw.split('&') if x]
        return []

    # Extract DC (5 digits)
    def is_dc(text):
        text = text.strip()
        if re.fullmatch(r'\d{5}', text): return True
        if re.match(r'Go\s*to\s*MDC\s*\d+', text, re.IGNORECASE): return True
        return False

    for node_id, node_data in nodes.items():
        text = node_data['val']
        pdcs = extract_pdc(text)
        
        for pdc in pdcs:
            # Fungsi BFS untuk mencari target signifikan (DC atau AX) dengan melompati node perantara
            def get_significant_nodes(start_ids):
                visited = set()
                queue = list(start_ids)
                significant = []
                while queue:
                    curr = queue.pop(0)
                    if curr in visited: continue
                    visited.add(curr)
                    
                    if curr not in nodes: continue
                    t_text = nodes[curr]['val']
                    
                    # Jika itu DC atau AX, ini signifikan
                    if is_dc(t_text) or extract_ax(t_text):
                        significant.append(curr)
                    # Jika itu PDC lain, JANGAN diteruskan (itu ranah PDC lain)
                    elif extract_pdc(t_text):
                        continue
                    else:
                        # Node perantara (misal "Go to MDC 34"), teruskan BFS
                        queue.extend(adj.get(curr, []))
                return significant

            significant_targets = get_significant_nodes(adj.get(node_id, []))
            
            for t_id in significant_targets:
                t_text = nodes[t_id]['val']
                
                # If target is a direct DC
                if is_dc(t_text):
                    flat_mapping[pdc] = t_text.strip()
                
                # If target is an AX node
                ax_list = extract_ax(t_text)
                if ax_list:
                    # Find where this AX node points to (again using BFS)
                    ax_dc_targets = get_significant_nodes(adj.get(t_id, []))
                    
                    dc_targets = []
                    for ax_t_id in ax_dc_targets:
                        ax_t_text = nodes[ax_t_id]['val']
                        if is_dc(ax_t_text):
                            dc_targets.append((nodes[ax_t_id]['y'], ax_t_text.strip()))
                    
                    if len(dc_targets) >= 2:
                        # Sort by Y descending. The one with larger Y is 'Yes' (AX satisfied)
                        dc_targets.sort(key=lambda x: x[0], reverse=True)
                        yes_dc = dc_targets[0][1]
                        no_dc = dc_targets[1][1]
                        
                        if pdc not in decision_tree:
                            decision_tree[pdc] = {'default': no_dc}
                        
                        for ax in ax_list:
                            decision_tree[pdc][ax] = yes_dc

    return decision_tree, flat_mapping

def extract_page_names(xml_string):
    root = ET.fromstring(xml_string)
    names = []
    for diagram in root.iter('diagram'):
        name = diagram.get('name')
        if name:
            names.append(name)
    return names

def extract_xml_from_pdf(pdf_file):
    import PyPDF2
    reader = PyPDF2.PdfReader(pdf_file)
    # The Drawio XML is usually inside the first page or the metadata
    subject = reader.metadata.get('/Subject', '')
    if isinstance(subject, str) and subject.startswith('%3Cmxfile'):
        return urllib.parse.unquote(subject)
    return None
