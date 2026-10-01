import os
import re
import tkinter as tk
from tkinter import filedialog
from collections import defaultdict
from itertools import combinations
from urllib.parse import quote
from decimal import Decimal, InvalidOperation

import pandas as pd
import xml.etree.ElementTree as ET
from openpyxl import load_workbook


MAX_ID_FIELDS = 6


# ============================================================
# BASIC HELPERS
# ============================================================

def clean(value):
    return "" if value is None else str(value).strip()


def strip_namespace(text):
    return re.sub(r"\{[^}]*\}", "", str(text))


def readable_path(path):
    return strip_namespace(path)


def canonical_number(value):
    """
    Returns canonical numeric string, or None if not numeric.

    28.0000 -> 28
    28      -> 28
    2.8E1   -> 28
    0.5000  -> 0.5
    """
    value = clean(value)

    try:
        number = Decimal(value)
    except (InvalidOperation, ValueError):
        return None

    if not number.is_finite():
        return None

    if number == 0:
        return "0"

    result = format(number.normalize(), "f")

    if "." in result:
        result = result.rstrip("0").rstrip(".")

    return result


def comparison_token(value):
    """
    Numeric-looking values compare numerically.
    Everything else compares as exact text.
    """
    value = clean(value)

    number = canonical_number(value)

    if number is not None:
        return ("NUMBER", number)

    return ("TEXT", value)


def values_equal(a, b):
    return comparison_token(a) == comparison_token(b)


# ============================================================
# FOLDER SELECTION
# ============================================================

def choose_folders(root):

    old_folder = filedialog.askdirectory(
        parent=root,
        title="Select OLD folder"
    )

    if not old_folder:
        return None

    new_folder = filedialog.askdirectory(
        parent=root,
        title="Select NEW folder"
    )

    if not new_folder:
        return None

    output_folder = filedialog.askdirectory(
        parent=root,
        title="Select Folder to Save Output"
    )

    if not output_folder:
        return None

    return old_folder, new_folder, output_folder


# ============================================================
# XML FILE COLLECTION
# ============================================================

def collect_xml_files(folder):

    files = {}

    for current_folder, _, filenames in os.walk(folder):

        for filename in filenames:

            if not filename.lower().endswith(".xml"):
                continue

            full_path = os.path.join(
                current_folder,
                filename
            )

            relative_path = os.path.relpath(
                full_path,
                folder
            )

            files[relative_path] = full_path

    return files


# ============================================================
# IDENTITY FACTS
# ============================================================

def identity_facts(element):
    """
    Collect exact scalar facts from this element.

    We may descend through singular branches:

        Tray
          -> AddOn
              -> AddOnId

    so Tray can use:

        AddOn/AddOnId = CC Mount

    But when a child tag repeats, we stop before that collection:

        AddOnValidation
        AddOnValidation
        AddOnValidation

    Those repeated objects establish their own identities later.
    """

    facts = {}

    def walk(node, prefix=""):

        # Attributes on current node
        for attribute, value in node.attrib.items():

            field = (
                f"{prefix}@{attribute}"
                if prefix
                else f"@{attribute}"
            )

            facts[field] = clean(value)

        children = list(node)

        # Leaf value
        if not children:

            if prefix:
                facts[f"{prefix}#text"] = clean(node.text)
            else:
                facts["#text"] = clean(node.text)

            return

        # Group immediate children by tag
        groups = defaultdict(list)

        for child in children:
            groups[child.tag].append(child)

        for tag, same_tag_children in groups.items():

            # Repeated collection:
            # do NOT use descendants to identify this parent.
            if len(same_tag_children) > 1:
                continue

            child = same_tag_children[0]

            child_prefix = (
                f"{prefix}{tag}/"
                if prefix
                else f"{tag}/"
            )

            walk(
                child,
                child_prefix
            )

    walk(element)

    return facts


# ============================================================
# EXACT IDENTITY MATCHING
# ============================================================

def fact_token(facts, field):

    if field not in facts:
        return ("MISSING", "")

    value_type, value = comparison_token(
        facts[field]
    )

    return (
        "PRESENT",
        value_type,
        value
    )


def identity_for(element, fields):

    facts = identity_facts(element)

    return tuple(
        fact_token(facts, field)
        for field in fields
    )


def identities_unique(elements, fields):

    if len(elements) <= 1:
        return True

    identities = [
        identity_for(element, fields)
        for element in elements
    ]

    return len(identities) == len(set(identities))


def overlap_count(
    old_elements,
    new_elements,
    fields
):

    old_ids = {
        identity_for(element, fields)
        for element in old_elements
    }

    new_ids = {
        identity_for(element, fields)
        for element in new_elements
    }

    return len(old_ids & new_ids)


def choose_identity_fields(
    old_elements,
    new_elements,
    parent_path,
    tag
):
    """
    Find the smallest exact property combination that:

    1. uniquely distinguishes OLD siblings
    2. uniquely distinguishes NEW siblings
    3. preserves as many exact OLD/NEW identities as possible

    No similarity scores.
    No sibling positions.
    """

    all_elements = (
        list(old_elements)
        +
        list(new_elements)
    )

    fact_sets = [
        identity_facts(element)
        for element in all_elements
    ]

    candidate_fields = set()

    for facts in fact_sets:
        candidate_fields.update(facts.keys())

    useful_fields = []

    for field in candidate_fields:

        tokens = {
            fact_token(facts, field)
            for facts in fact_sets
        }

        if len(tokens) > 1:
            useful_fields.append(field)

    useful_fields = sorted(useful_fields)

    if not useful_fields:

        raise ValueError(
            "\nCannot distinguish repeated XML objects.\n\n"
            f"Parent:\n{readable_path(parent_path)}\n\n"
            f"Repeated element:\n{strip_namespace(tag)}\n\n"
            "No distinguishing scalar values were found "
            "through non-repeating descendant branches."
        )

    max_possible_overlap = min(
        len(old_elements),
        len(new_elements)
    )

    best_fields = None
    best_overlap = -1

    max_size = min(
        MAX_ID_FIELDS,
        len(useful_fields)
    )

    for size in range(1, max_size + 1):

        best_this_size = None
        best_this_overlap = -1

        for fields in combinations(
            useful_fields,
            size
        ):

            if not identities_unique(
                old_elements,
                fields
            ):
                continue

            if not identities_unique(
                new_elements,
                fields
            ):
                continue

            overlap = overlap_count(
                old_elements,
                new_elements,
                fields
            )

            if (
                overlap > best_this_overlap
                or (
                    overlap == best_this_overlap
                    and (
                        best_this_size is None
                        or fields < best_this_size
                    )
                )
            ):
                best_this_overlap = overlap
                best_this_size = fields

            if (
                overlap > best_overlap
                or (
                    overlap == best_overlap
                    and (
                        best_fields is None
                        or len(fields) < len(best_fields)
                    )
                )
            ):
                best_overlap = overlap
                best_fields = fields

        # Perfect overlap using smallest possible field count.
        if (
            best_this_size is not None
            and best_this_overlap == max_possible_overlap
        ):
            return (
                best_this_size,
                best_this_overlap
            )

    if best_fields is not None:
        return (
            best_fields,
            best_overlap
        )

    raise ValueError(
        "\nCannot uniquely identify repeated XML objects "
        "without positional matching.\n\n"
        f"Parent:\n{readable_path(parent_path)}\n\n"
        f"Repeated element:\n{strip_namespace(tag)}"
    )


# ============================================================
# PATH CREATION
# ============================================================

def field_label(field):

    field = strip_namespace(field)

    if field == "#text":
        return "value"

    if field.endswith("/#text"):
        field = field[:-6]

    return field.rstrip("/")


def identity_display_value(token):

    if token[0] == "MISSING":
        return "<MISSING>"

    value = token[2]

    return quote(
        value,
        safe="-_.~ "
    )


def identity_segment(
    tag,
    fields,
    identity
):

    segment = str(tag)

    for field, token in zip(
        fields,
        identity
    ):

        segment += (
            f"[{field_label(field)}="
            f"{identity_display_value(token)}]"
        )

    return segment


# ============================================================
# DICTIONARY BUILDING
# ============================================================

def add_value(
    dictionary,
    key,
    value
):

    if key in dictionary:

        raise ValueError(
            "\nDuplicate contextual path generated:\n\n"
            f"{readable_path(key)}"
        )

    dictionary[key] = clean(value)


def record_element(
    element,
    path,
    dictionary
):

    if element is None:
        return

    # Attributes
    for attribute, value in element.attrib.items():

        add_value(
            dictionary,
            f"{path}/@{attribute}",
            value
        )

    children = list(element)

    # Leaf value
    if not children:

        add_value(
            dictionary,
            path,
            element.text
        )

        return

    # Preserve meaningful mixed-content text if present
    text = clean(element.text)

    if text:

        add_value(
            dictionary,
            f"{path}/#text",
            text
        )


def flatten_pair(
    old_element,
    new_element,
    path,
    old_dictionary,
    new_dictionary,
    filename,
    identity_rules
):

    record_element(
        old_element,
        path,
        old_dictionary
    )

    record_element(
        new_element,
        path,
        new_dictionary
    )

    old_groups = defaultdict(list)
    new_groups = defaultdict(list)

    if old_element is not None:

        for child in old_element:
            old_groups[child.tag].append(child)

    if new_element is not None:

        for child in new_element:
            new_groups[child.tag].append(child)

    all_tags = sorted(
        set(old_groups)
        |
        set(new_groups),
        key=str
    )

    for tag in all_tags:

        old_children = old_groups.get(
            tag,
            []
        )

        new_children = new_groups.get(
            tag,
            []
        )

        # ----------------------------------------------------
        # Non-repeated child
        # ----------------------------------------------------

        if max(
            len(old_children),
            len(new_children)
        ) <= 1:

            old_child = (
                old_children[0]
                if old_children
                else None
            )

            new_child = (
                new_children[0]
                if new_children
                else None
            )

            child_path = (
                f"{path}/{tag}"
            )

            flatten_pair(
                old_child,
                new_child,
                child_path,
                old_dictionary,
                new_dictionary,
                filename,
                identity_rules
            )

            continue

        # ----------------------------------------------------
        # Repeated sibling objects
        # ----------------------------------------------------

        identity_fields, overlap = (
            choose_identity_fields(
                old_children,
                new_children,
                path,
                tag
            )
        )

        identity_rules.append({
            "File": filename,
            "Parent Path": readable_path(path),
            "Repeated Element": strip_namespace(tag),
            "Identity Fields": " + ".join(
                field_label(field)
                for field in identity_fields
            ),
            "OLD Count": len(old_children),
            "NEW Count": len(new_children),
            "Exact Matches": overlap
        })

        old_map = {}

        for child in old_children:

            identity = identity_for(
                child,
                identity_fields
            )

            if identity in old_map:
                raise ValueError(
                    "Duplicate OLD identity at "
                    f"{readable_path(path)}/"
                    f"{strip_namespace(tag)}"
                )

            old_map[identity] = child

        new_map = {}

        for child in new_children:

            identity = identity_for(
                child,
                identity_fields
            )

            if identity in new_map:
                raise ValueError(
                    "Duplicate NEW identity at "
                    f"{readable_path(path)}/"
                    f"{strip_namespace(tag)}"
                )

            new_map[identity] = child

        all_identities = sorted(
            set(old_map)
            |
            set(new_map),
            key=repr
        )

        for identity in all_identities:

            segment = identity_segment(
                tag,
                identity_fields,
                identity
            )

            child_path = (
                f"{path}/{segment}"
            )

            flatten_pair(
                old_map.get(identity),
                new_map.get(identity),
                child_path,
                old_dictionary,
                new_dictionary,
                filename,
                identity_rules
            )


# ============================================================
# XML PAIR
# ============================================================

def parse_xml_pair(
    old_path,
    new_path,
    filename,
    identity_rules
):

    old_root = ET.parse(
        old_path
    ).getroot()

    new_root = ET.parse(
        new_path
    ).getroot()

    if old_root.tag != new_root.tag:

        raise ValueError(
            "OLD and NEW root elements differ: "
            f"{strip_namespace(old_root.tag)} vs "
            f"{strip_namespace(new_root.tag)}"
        )

    old_dictionary = {}
    new_dictionary = {}

    root_path = f"/{old_root.tag}"

    flatten_pair(
        old_root,
        new_root,
        root_path,
        old_dictionary,
        new_dictionary,
        filename,
        identity_rules
    )

    return (
        old_dictionary,
        new_dictionary
    )


# ============================================================
# DICTIONARY COMPARISON
# ============================================================

def compare_dictionaries(
    old_dictionary,
    new_dictionary,
    filename
):

    rows = []

    all_keys = sorted(
        set(old_dictionary)
        |
        set(new_dictionary)
    )

    for key in all_keys:

        in_old = key in old_dictionary
        in_new = key in new_dictionary

        old_value = old_dictionary.get(
            key,
            ""
        )

        new_value = new_dictionary.get(
            key,
            ""
        )

        if not in_old:

            status = "Added"

        elif not in_new:

            status = "Removed"

        elif not values_equal(
            old_value,
            new_value
        ):

            status = "Changed"

        else:

            continue

        rows.append({
            "File": filename,
            "Status": status,
            "Key": readable_path(key),
            "Before": old_value,
            "After": new_value
        })

    return rows


# ============================================================
# FOLDER COMPARISON
# ============================================================

def compare_folders(
    old_folder,
    new_folder
):

    old_files = collect_xml_files(
        old_folder
    )

    new_files = collect_xml_files(
        new_folder
    )

    rows = []
    identity_rules = []

    all_files = sorted(
        set(old_files)
        |
        set(new_files)
    )

    total = len(all_files)

    for index, relative_path in enumerate(
        all_files,
        start=1
    ):

        print(
            f"[{index}/{total}] {relative_path}",
            flush=True
        )

        old_path = old_files.get(
            relative_path
        )

        new_path = new_files.get(
            relative_path
        )

        if old_path is None:

            rows.append({
                "File": relative_path,
                "Status": "Added",
                "Key": "/",
                "Before": "",
                "After": "XML file present"
            })

            continue

        if new_path is None:

            rows.append({
                "File": relative_path,
                "Status": "Removed",
                "Key": "/",
                "Before": "XML file present",
                "After": ""
            })

            continue

        try:

            (
                old_dictionary,
                new_dictionary
            ) = parse_xml_pair(
                old_path,
                new_path,
                relative_path,
                identity_rules
            )

            rows.extend(
                compare_dictionaries(
                    old_dictionary,
                    new_dictionary,
                    relative_path
                )
            )

        except Exception as error:

            rows.append({
                "File": relative_path,
                "Status": "Error",
                "Key": "",
                "Before": "",
                "After": str(error)
            })

    differences_df = pd.DataFrame(
        rows,
        columns=[
            "File",
            "Status",
            "Key",
            "Before",
            "After"
        ]
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
            "Exact Matches"
        ]
    )

    return differences_df, identity_df


# ============================================================
# EXCEL OUTPUT
# ============================================================

def next_report_path(folder):

    base = "xml_comparison_report"

    path = os.path.join(
        folder,
        f"{base}.xlsx"
    )

    number = 1

    while os.path.exists(path):

        path = os.path.join(
            folder,
            f"{base}_{number}.xlsx"
        )

        number += 1

    return path


def format_excel(filepath):

    workbook = load_workbook(
        filepath
    )

    for worksheet in workbook.worksheets:

        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = (
            worksheet.dimensions
        )

        for column_cells in worksheet.columns:

            max_length = max(
                len(str(cell.value or ""))
                for cell in column_cells
            )

            worksheet.column_dimensions[
                column_cells[0].column_letter
            ].width = min(
                max_length + 2,
                100
            )

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
        print("Cancelled.")
        return

    (
        old_folder,
        new_folder,
        output_folder
    ) = folders

    print("\nStarting comparison...\n")

    differences, identity_rules = (
        compare_folders(
            old_folder,
            new_folder
        )
    )

    report_path = next_report_path(
        output_folder
    )

    with pd.ExcelWriter(
        report_path,
        engine="openpyxl"
    ) as writer:

        differences.to_excel(
            writer,
            sheet_name="Differences",
            index=False
        )

        identity_rules.to_excel(
            writer,
            sheet_name="Identity Rules",
            index=False
        )

    format_excel(report_path)

    print()
    print(
        f"Found {len(differences)} "
        f"reportable differences."
    )

    print(
        f"Report saved to:\n{report_path}"
    )


if __name__ == "__main__":
    main()