import os
import re
import tkinter as tk
from tkinter import filedialog
from collections import defaultdict
from itertools import combinations
from urllib.parse import quote
from decimal import Decimal, InvalidOperation

import pandas as pd
from lxml import etree
from openpyxl import load_workbook

from html_report import generate_html_reports


MAX_ID_FIELDS = 6

# ============================================================
# VALUE HELPERS
# ============================================================

def clean(value):
    return '' if value is None else str(value).strip()

def strip_namespace(text):
    return re.sub('\\{[^}]*\\}', '', str(text))

def readable_path(path):
    return strip_namespace(path)

def canonical_number(value):
    value = clean(value)
    try:
        number = Decimal(value)
    except (InvalidOperation, ValueError):
        return None
    if not number.is_finite():
        return None
    if number == 0:
        return '0'
    result = format(number.normalize(), 'f')
    if '.' in result:
        result = result.rstrip('0').rstrip('.')
    return result

def comparison_token(value):
    value = clean(value)
    number = canonical_number(value)
    if number is not None:
        return ('NUMBER', number)
    return ('TEXT', value)

def values_equal(a, b):
    return comparison_token(a) == comparison_token(b)

# ============================================================
# FOLDER SELECTION
# ============================================================

def choose_folders(root):
    old_folder = filedialog.askdirectory(parent=root, title='Select OLD folder')
    if not old_folder:
        return None
    new_folder = filedialog.askdirectory(parent=root, title='Select NEW folder')
    if not new_folder:
        return None
    output_folder = filedialog.askdirectory(parent=root, title='Select output folder')
    if not output_folder:
        return None
    return (old_folder, new_folder, output_folder)

# ============================================================
# FILE COLLECTION
# ============================================================

def collect_xml_files(folder):
    files = {}
    for current, _, filenames in os.walk(folder):
        for filename in filenames:
            if not filename.lower().endswith('.xml'):
                continue
            full_path = os.path.join(current, filename)
            relative_path = os.path.relpath(full_path, folder)
            files[relative_path] = full_path
    return files

def element_children(element):
    return [child for child in element if isinstance(child.tag, str)]

# ============================================================
# PRETTY XML
# ============================================================

def parse_pretty_xml(path):
    """
    Pretty-print XML IN MEMORY.

    Original XML is not changed.

    We then reparse the pretty XML so that element.sourceline
    corresponds to the lines displayed in the HTML report.
    """
    parser = etree.XMLParser(
        remove_blank_text=True,
        remove_comments=False,
        huge_tree=True,
    )
    tree = etree.parse(path, parser)
    pretty_bytes = etree.tostring(
        tree,
        pretty_print=True,
        encoding="utf-8",
        xml_declaration=True,
    )
    pretty_text = pretty_bytes.decode('utf-8')
    display_parser = etree.XMLParser(
        remove_blank_text=False,
        remove_comments=False,
        huge_tree=True,
    )
    display_root = etree.fromstring(pretty_bytes, display_parser)
    return (display_root, pretty_text)

# ============================================================
# IDENTITY FACTS
# ============================================================

def identity_facts(element):
    """
    Descend through singular branches.

    Stop before repeated same-tag collections.
    """
    facts = {}

    def walk(node, prefix=''):
        for attribute, value in node.attrib.items():
            key = f'{prefix}@{attribute}' if prefix else f'@{attribute}'
            facts[key] = clean(value)
        children = element_children(node)
        if not children:
            key = f'{prefix}#text' if prefix else '#text'
            facts[key] = clean(node.text)
            return
        groups = defaultdict(list)
        for child in children:
            groups[child.tag].append(child)
        for tag, siblings in groups.items():
            if len(siblings) > 1:
                continue
            child = siblings[0]
            walk(child, f'{prefix}{tag}/')
    walk(element)
    return facts

# ============================================================
# EXACT IDENTITY MATCHING
# ============================================================

def fact_token(facts, field):
    if field not in facts:
        return ('MISSING',)
    kind, value = comparison_token(facts[field])
    return ('PRESENT', kind, value)

def identity_for(element, fields):
    facts = identity_facts(element)
    return tuple((fact_token(facts, field) for field in fields))

def identities_unique(elements, fields):
    identities = [identity_for(element, fields) for element in elements]
    return len(identities) == len(set(identities))  # set removes duplicates

def overlap_count(old_elements, new_elements, fields):
    old_ids = {identity_for(element, fields) for element in old_elements}
    new_ids = {identity_for(element, fields) for element in new_elements}
    return len(old_ids & new_ids)  # & is set intersection

def choose_identity_fields(old_elements, new_elements, parent_path, tag):
    # For repeated elements, choose the best field(s) to identify each sibling.
    all_elements = list(old_elements) + list(new_elements)
    fact_sets = [identity_facts(element) for element in all_elements]
    candidate_fields = set()
    for facts in fact_sets:
        candidate_fields.update(facts.keys())
    useful_fields = []
    for field in candidate_fields:
        tokens = {fact_token(facts, field) for facts in fact_sets}
        if len(tokens) > 1:  # field varies, so it may help distinguish siblings
            useful_fields.append(field)
    useful_fields.sort()
    if not useful_fields:
        raise ValueError(
            "\nCannot distinguish repeated XML objects.\n\n"
            f"Parent:\n{readable_path(parent_path)}\n\n"
            f"Repeated element:\n{strip_namespace(tag)}"
        )
    maximum_overlap = min(len(old_elements), len(new_elements))
    best_fields = None
    best_overlap = -1
    max_size = min(MAX_ID_FIELDS, len(useful_fields))
    for size in range(1, max_size + 1):
        best_this_size = None
        best_this_overlap = -1
        for fields in combinations(useful_fields, size):
            if not identities_unique(old_elements, fields):
                continue
            if not identities_unique(new_elements, fields):
                continue
            overlap = overlap_count(old_elements, new_elements, fields)
            if overlap > best_this_overlap:
                best_this_overlap = overlap
                best_this_size = fields
            if (
                overlap > best_overlap
                or (
                    overlap == best_overlap
                    and (best_fields is None or len(fields) < len(best_fields))
                )
            ):
                best_overlap = overlap
                best_fields = fields
        if best_this_size is not None and best_this_overlap == maximum_overlap:
            return (best_this_size, best_this_overlap)
    if best_fields is not None:
        return (best_fields, best_overlap)
    raise ValueError(
        "\nCannot uniquely identify repeated XML objects "
        "without positional matching.\n\n"
        f"Parent:\n{readable_path(parent_path)}\n\n"
        f"Repeated element:\n{strip_namespace(tag)}"
    )

# ============================================================
# PATH HELPERS
# ============================================================

def field_label(field):
    field = strip_namespace(field)
    if field == '#text':
        return 'value'
    if field.endswith('/#text'):
        field = field[:-6]
    return field.rstrip('/')

def identity_value(token):  # return the value used in an identity path
    if token[0] == 'MISSING':
        return '<MISSING>'
    return quote(token[2], safe='-_.~ ')

def identity_segment(tag, fields, identity):
    segment = str(tag)
    for field, token in zip(fields, identity):
        segment += f'[{field_label(field)}={identity_value(token)}]'
    return segment  # e.g. Beam[Name=A]

# ============================================================
# BUILD DICTIONARIES
# ============================================================

def add_value(values, lines, key, value, line):
    if key in values:
        raise ValueError(f'\nDuplicate contextual path:\n\n{readable_path(key)}')
    values[key] = clean(value)
    lines[key] = line

def record_element(element, path, values, lines):
    if element is None:
        return
    line = element.sourceline
    for attribute, value in element.attrib.items():
        add_value(values, lines, f'{path}/@{attribute}', value, line)
    children = element_children(element)
    if not children:
        add_value(values, lines, path, element.text, line)
        return
    text = clean(element.text)
    if text:
        add_value(values, lines, f'{path}/#text', text, line)

# ============================================================
# RECURSIVE FLATTENING
# ============================================================

def flatten_pair(
    old_element,
    new_element,
    path,
    old_values,
    new_values,
    old_lines,
    new_lines,
    filename,
    identity_rules,
):
    record_element(old_element, path, old_values, old_lines)
    record_element(new_element, path, new_values, new_lines)
    old_groups = defaultdict(list)
    new_groups = defaultdict(list)
    if old_element is not None:
        for child in element_children(old_element):
            old_groups[child.tag].append(child)
    if new_element is not None:
        for child in element_children(new_element):
            new_groups[child.tag].append(child)
    all_tags = sorted(set(old_groups) | set(new_groups), key=str)
    for tag in all_tags:
        old_children = old_groups.get(tag, [])
        new_children = new_groups.get(tag, [])
        if max(len(old_children), len(new_children)) <= 1:
            old_child = old_children[0] if old_children else None
            new_child = new_children[0] if new_children else None
            flatten_pair(
                old_child,
                new_child,
                f"{path}/{tag}",
                old_values,
                new_values,
                old_lines,
                new_lines,
                filename,
                identity_rules,
            )
            continue
        fields, overlap = choose_identity_fields(
            old_children, new_children, path, tag
        )
        identity_rules.append(
            {
                "File": filename,
                "Parent Path": readable_path(path),
                "Repeated Element": strip_namespace(tag),
                "Identity Fields": " + ".join(
                    field_label(field) for field in fields
                ),
                "OLD Count": len(old_children),
                "NEW Count": len(new_children),
                "Exact Matches": overlap,
            }
        )
        old_map = {}
        for child in old_children:
            identity = identity_for(child, fields)
            if identity in old_map:
                raise ValueError(
                    "\nDuplicate OLD identity:\n\n"
                    f"{readable_path(path)}/{strip_namespace(tag)}"
                )
            old_map[identity] = child
        new_map = {}
        for child in new_children:
            identity = identity_for(child, fields)
            if identity in new_map:
                raise ValueError(
                    "\nDuplicate NEW identity:\n\n"
                    f"{readable_path(path)}/{strip_namespace(tag)}"
                )
            new_map[identity] = child
        identities = sorted(set(old_map) | set(new_map), key=repr)
        for identity in identities:
            segment = identity_segment(tag, fields, identity)
            flatten_pair(
                old_map.get(identity),
                new_map.get(identity),
                f"{path}/{segment}",
                old_values,
                new_values,
                old_lines,
                new_lines,
                filename,
                identity_rules,
            )

# ============================================================
# PARSE OLD / NEW XML PAIR
# ============================================================

def parse_xml_pair(old_path, new_path, filename, identity_rules):
    old_root, old_pretty_text = parse_pretty_xml(old_path)
    new_root, new_pretty_text = parse_pretty_xml(new_path)
    if old_root.tag != new_root.tag:
        raise ValueError(
            "OLD and NEW root elements differ: "
            f"{strip_namespace(old_root.tag)} vs {strip_namespace(new_root.tag)}"
        )
    old_values = {}
    new_values = {}
    old_lines = {}
    new_lines = {}
    root_path = f'/{old_root.tag}'
    flatten_pair(
        old_root,
        new_root,
        root_path,
        old_values,
        new_values,
        old_lines,
        new_lines,
        filename,
        identity_rules,
    )
    return (old_values, new_values, old_lines, new_lines, old_pretty_text, new_pretty_text)

# ============================================================
# DICTIONARY COMPARISON
# ============================================================

def compare_dictionaries(
    old_values, new_values, old_lines, new_lines, filename
):
    rows = []
    keys = sorted(set(old_values) | set(new_values))
    for key in keys:
        in_old = key in old_values
        in_new = key in new_values
        old_value = old_values.get(key, '')
        new_value = new_values.get(key, '')
        if not in_old:
            status = 'Added'
        elif not in_new:
            status = 'Removed'
        elif not values_equal(old_value, new_value):
            status = 'Changed'
        else:
            continue
        rows.append(
            {
                "File": filename,
                "Status": status,
                "Key": readable_path(key),
                "Before": old_value,
                "After": new_value,
                "Old Line": old_lines.get(key),
                "New Line": new_lines.get(key),
            }
        )
    return rows

# ============================================================
# FOLDER COMPARISON
# ============================================================

def compare_folders(old_folder, new_folder):
    old_files = collect_xml_files(old_folder)
    new_files = collect_xml_files(new_folder)
    rows = []
    identity_rules = []
    html_items = []
    all_files = sorted(set(old_files) | set(new_files))
    total = len(all_files)
    for index, relative_path in enumerate(all_files, start=1):
        print(f'[{index}/{total}] {relative_path}', flush=True)
        old_path = old_files.get(relative_path)
        new_path = new_files.get(relative_path)
        if old_path is None:
            row = {
                "File": relative_path,
                "Status": "Added",
                "Key": "/",
                "Before": "",
                "After": "XML file present",
                "Old Line": None,
                "New Line": None,
            }
            rows.append(row)
            html_items.append(
                {
                    "file": relative_path,
                    "old_path": None,
                    "new_path": new_path,
                    "old_text": "",
                    "new_text": read_pretty_for_whole_file(new_path),
                    "differences": [row],
                    "whole_file_status": "Added",
                }
            )
            continue
        if new_path is None:
            row = {
                "File": relative_path,
                "Status": "Removed",
                "Key": "/",
                "Before": "XML file present",
                "After": "",
                "Old Line": None,
                "New Line": None,
            }
            rows.append(row)
            html_items.append(
                {
                    "file": relative_path,
                    "old_path": old_path,
                    "new_path": None,
                    "old_text": read_pretty_for_whole_file(old_path),
                    "new_text": "",
                    "differences": [row],
                    "whole_file_status": "Removed",
                }
            )
            continue
        try:
            (
                old_values,
                new_values,
                old_lines,
                new_lines,
                old_pretty_text,
                new_pretty_text,
            ) = parse_xml_pair(old_path, new_path, relative_path, identity_rules)
            file_rows = compare_dictionaries(
                old_values, new_values, old_lines, new_lines, relative_path
            )
            rows.extend(file_rows)
            if file_rows:
                html_items.append(
                    {
                        "file": relative_path,
                        "old_path": old_path,
                        "new_path": new_path,
                        "old_text": old_pretty_text,
                        "new_text": new_pretty_text,
                        "differences": file_rows,
                        "whole_file_status": None,
                    }
                )
        except Exception as error:
            rows.append(
                {
                    "File": relative_path,
                    "Status": "Error",
                    "Key": "",
                    "Before": "",
                    "After": str(error),
                    "Old Line": None,
                    "New Line": None,
                }
            )
    difference_columns = ['File', 'Status', 'Key', 'Before', 'After']
    differences_df = pd.DataFrame(
        [
            {column: row.get(column, "") for column in difference_columns}
            for row in rows
        ],
        columns=difference_columns,
    )
    identity_df = pd.DataFrame(
        identity_rules,
        columns=[
            "File",
            "Parent Path",
            "Repeated Element",
            "Identity Fields",
            "OLD Count",
            "NEW Count",
            "Exact Matches",
        ],
    )
    return (differences_df, identity_df, html_items)

# ============================================================
# PRETTY DISPLAY FOR WHOLE FILE ADD / REMOVE
# ============================================================

def read_pretty_for_whole_file(path):
    if path is None:
        return ''
    try:
        _, pretty_text = parse_pretty_xml(path)
        return pretty_text
    except Exception:
        with open(path, 'r', encoding='utf-8', errors='replace') as f:
            return f.read()

# ============================================================
# OUTPUT
# ============================================================

def next_report_path(folder):
    base = 'xml_comparison_report'
    path = os.path.join(folder, f'{base}.xlsx')
    number = 1
    while os.path.exists(path):
        path = os.path.join(folder, f'{base}_{number}.xlsx')
        number += 1
    return path

def format_excel(filepath):
    workbook = load_workbook(filepath)
    for worksheet in workbook.worksheets:
        worksheet.freeze_panes = 'A2'
        worksheet.auto_filter.ref = worksheet.dimensions
        for cells in worksheet.columns:
            width = max((len(str(cell.value or '')) for cell in cells))
            worksheet.column_dimensions[cells[0].column_letter].width = min(width + 2, 100)
    workbook.save(filepath)

# ============================================================
# MAIN
# ============================================================

def main():
    root = tk.Tk()
    root.withdraw()
    folders = choose_folders(root)
    root.destroy()
    if not folders:
        print('Cancelled.')
        return
    old_folder, new_folder, output_folder = folders
    print('\nStarting comparison...\n')
    differences, identity_rules, html_items = compare_folders(old_folder, new_folder)
    excel_path = next_report_path(output_folder)
    report_stem = os.path.splitext(os.path.basename(excel_path))[0]
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        differences.to_excel(writer, sheet_name='Differences', index=False)
        identity_rules.to_excel(writer, sheet_name='Identity Rules', index=False)
    format_excel(excel_path)
    html_path = generate_html_reports(html_items, output_folder, report_stem)
    print()
    print(f'Found {len(differences)} reportable differences.')
    print(f'\nExcel:\n{excel_path}')
    print(f'\nHTML:\n{html_path}')


if __name__ == '__main__':
    main()
