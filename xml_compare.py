import pandas as pd
import xml.etree.ElementTree as ET
import os
import tkinter as tk
from tkinter import filedialog
from openpyxl.styles import PatternFill
from openpyxl import load_workbook


# ========== CONFIG ==========

gray_fill = PatternFill(
    start_color="BFBFBF",
    end_color="BFBFBF",
    fill_type="solid"
)

green_fill = PatternFill(
    start_color="CCFFCC",
    end_color="CCFFCC",
    fill_type="solid"
)

orange_fill = PatternFill(
    start_color="FFA500",
    end_color="FFA500",
    fill_type="solid"
)

row_fill = PatternFill(
    start_color="FFFFCC",
    end_color="FFFFCC",
    fill_type="solid"
)


# ========== STEP 1: Select XML ==========

def _center_window(win, width, height):
    win.update_idletasks()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = (sw - width) // 2
    y = (sh - height) // 2
    win.geometry(f"{width}x{height}+{x}+{y}")


def choose_mode_and_files(root):
    choice = {"mode": None}

    dialog = tk.Toplevel(root)
    dialog.title("XML Comparison Tool")
    dialog.resizable(False, False)

    dialog.attributes("-topmost", True)
    _center_window(dialog, 520, 170)

    def set_mode(mode):
        choice["mode"] = mode
        dialog.destroy()

    def cancel():
        choice["mode"] = None
        dialog.destroy()

    dialog.protocol("WM_DELETE_WINDOW", cancel)
    dialog.bind("<Escape>", lambda _e: cancel())

    tk.Label(dialog, text="Choose mode:").pack(
        padx=16,
        pady=(16, 10)
    )

    frame = tk.Frame(dialog)
    frame.pack(padx=16, pady=(0, 10))

    tk.Button(
        frame,
        text="1 file: Extract",
        width=22,
        command=lambda: set_mode("extract")
    ).grid(row=0, column=0, padx=6, pady=4)

    tk.Button(
        frame,
        text="2 files: Compare",
        width=22,
        command=lambda: set_mode("compare")
    ).grid(row=0, column=1, padx=6, pady=4)

    tk.Button(
        dialog,
        text="Cancel",
        width=12,
        command=cancel
    ).pack(pady=(0, 14))

    dialog.lift()
    dialog.focus_force()
    dialog.after(
        300,
        lambda: dialog.attributes("-topmost", False)
    )

    root.wait_window(dialog)

    mode = choice["mode"]

    if mode is None:
        return []

    if mode == "extract":
        path = filedialog.askopenfilename(
            parent=root,
            title="Select ONE XML file to extract",
            filetypes=[("XML files", "*.xml")]
        )

        return [path] if path else []

    before_path = filedialog.askopenfilename(
        parent=root,
        title="Select BEFORE (older) XML file",
        filetypes=[("XML files", "*.xml")]
    )

    if not before_path:
        return []

    after_path = filedialog.askopenfilename(
        parent=root,
        title="Select AFTER (newer) XML file",
        filetypes=[("XML files", "*.xml")]
    )

    if not after_path:
        return []

    return [before_path, after_path]


def select_output_folder(root):
    return filedialog.askdirectory(
        parent=root,
        title="Select Folder to Save Outputs"
    )


# ========== STEP 2: Parse XML ==========

def parse_privileges(root):
    rows = []

    rows.append({
        "privilegeid": "[FILE]",
        "privilegeversion": "",
        "privilegecategory": "",
        "privilegegroup": "",
        "privilegedisplayname": "",
        "privilegesversion": root.findtext("privilegesversion", default="")
    })

    for privilege in root.findall("privilege"):
        row = {}

        row["privilegeid"] = privilege.findtext("privilegeid")
        row["privilegeversion"] = privilege.findtext("privilegeversion")
        row["privilegecategory"] = privilege.findtext("privilegecategory")
        row["privilegegroup"] = privilege.findtext("privilegegroup")
        row["privilegedisplayname"] = privilege.findtext(
            "privilegedisplayname"
        )
        row["privilegesversion"] = ""

        for group in privilege.findall("usergroups/groupcuid"):
            group_name = group.text
            allow_value = group.get("allow")

            row[group_name] = allow_value

        rows.append(row)

    df = pd.DataFrame(rows).fillna("")

    fixed_cols = [
        "privilegeid",
        "privilegeversion",
        "privilegecategory",
        "privilegegroup",
        "privilegedisplayname",
        "privilegesversion"
    ]

    group_cols = sorted(
        col for col in df.columns
        if col not in fixed_cols
    )

    df = df[fixed_cols + group_cols]
    df = df.sort_values("privilegeid").reset_index(drop=True)

    id_col = "privilegeid"

    return df, id_col, fixed_cols


def parse_users(root):
    rows = []

    for user in root.findall("users/user"):
        row = {}

        row["userid"] = user.findtext("userid")
        row["username"] = user.findtext("username")
        row["isaccountdisabled"] = user.findtext("isaccountdisabled")
        row["groupcuid"] = user.findtext("groupcuid")
        row["provider"] = user.findtext("provider")

        rows.append(row)

    df = pd.DataFrame(rows).fillna("")

    fixed_cols = [
        "userid",
        "username",
        "isaccountdisabled",
        "groupcuid",
        "provider"
    ]

    df = df[fixed_cols]
    df = df.sort_values("userid").reset_index(drop=True)

    id_col = "userid"

    return df, id_col, fixed_cols


def parse_xml_file(xml_path):
    tree = ET.parse(xml_path)
    root = tree.getroot()

    if root.tag == "privileges":
        df, id_col, fixed_cols = parse_privileges(root)
        file_type = "privileges"

    elif root.tag == "OSPAccessImport":
        df, id_col, fixed_cols = parse_users(root)
        file_type = "users"

    else:
        raise ValueError(
            f"Unsupported XML type: {root.tag}"
        )

    return df, file_type, id_col, fixed_cols


# ========== STEP 3: Compare ==========

def df_to_dict(df, id_col):
    result = {}

    for _, row in df.iterrows():
        row_id = row[id_col]
        result[row_id] = row.to_dict()

    return result


def compare_files(df_before, df_after, id_col):
    dict_before = df_to_dict(df_before, id_col)
    dict_after = df_to_dict(df_after, id_col)

    ids_before = set(dict_before.keys())
    ids_after = set(dict_after.keys())

    removed_ids = ids_before - ids_after
    added_ids = ids_after - ids_before
    common_ids = ids_before & ids_after

    removed_rows = []
    added_rows = []
    changed_rows = []

    for row_id in removed_ids:
        removed_rows.append(
            dict_before[row_id]
        )

    for row_id in added_ids:
        added_rows.append(
            dict_after[row_id]
        )

    all_columns = sorted(
        (
            set(df_before.columns)
            | set(df_after.columns)
        )
        - {id_col}
    )

    for row_id in common_ids:
        before_row = dict_before[row_id]
        after_row = dict_after[row_id]

        for column in all_columns:
            before_value = str(
                before_row.get(column, "")
            )

            after_value = str(
                after_row.get(column, "")
            )

            if before_value != after_value:
                changed_rows.append({
                    id_col: row_id,
                    "Parameter": column,
                    "Before": before_value,
                    "After": after_value
                })

    removed_df = pd.DataFrame(removed_rows)
    added_df = pd.DataFrame(added_rows)
    changed_df = pd.DataFrame(changed_rows)

    return removed_df, added_df, changed_df


def align_columns_for_output(
    df_before,
    df_after,
    fixed_cols
):
    extra_cols = sorted(
        (
            set(df_before.columns)
            | set(df_after.columns)
        )
        - set(fixed_cols)
    )

    all_cols = fixed_cols + extra_cols

    df_before = df_before.reindex(
        columns=all_cols,
        fill_value=""
    )

    df_after = df_after.reindex(
        columns=all_cols,
        fill_value=""
    )

    return df_before, df_after


# ========== STEP 4: Highlight ==========

def highlight_rows(
    filepath,
    highlight_df,
    fill_color,
    id_col
):
    if highlight_df.empty:
        return

    wb = load_workbook(filepath)
    ws = wb.active

    df_excel = pd.read_excel(filepath).fillna("")

    highlight_ids = set(
        highlight_df[id_col]
    )

    for index, row in df_excel.iterrows():
        if row[id_col] in highlight_ids:
            excel_row = index + 2

            for cell in ws[excel_row]:
                cell.fill = fill_color

    wb.save(filepath)


def highlight_changes(
    filepath,
    changed_df,
    id_col
):
    if changed_df.empty:
        return

    wb = load_workbook(filepath)
    ws = wb.active

    df_excel = pd.read_excel(filepath).fillna("")

    changes = {}

    for _, row in changed_df.iterrows():
        row_id = row[id_col]
        parameter = row["Parameter"]

        if row_id not in changes:
            changes[row_id] = []

        changes[row_id].append(parameter)

    for index, row in df_excel.iterrows():
        row_id = row[id_col]

        if row_id in changes:
            excel_row = index + 2

            # Highlight changed row
            for cell in ws[excel_row]:
                cell.fill = row_fill

            # Highlight changed cells
            for parameter in changes[row_id]:
                if parameter in df_excel.columns:
                    column_number = (
                        df_excel.columns.get_loc(parameter)
                        + 1
                    )

                    ws.cell(
                        row=excel_row,
                        column=column_number
                    ).fill = orange_fill

    wb.save(filepath)


def autosize_excel_columns(filepath):
    wb = load_workbook(filepath)
    ws = wb.active

    for column_cells in ws.columns:
        length = max(
            len(str(cell.value or ""))
            for cell in column_cells
        )

        ws.column_dimensions[
            column_cells[0].column_letter
        ].width = length + 2

    wb.save(filepath)


# ========== STEP 5: Main ==========

def main():
    root = tk.Tk()
    root.withdraw()

    xml_files = choose_mode_and_files(root)

    if not xml_files:
        print("Cancelled.")
        raise SystemExit

    save_folder = select_output_folder(root)

    if not save_folder:
        print("Cancelled.")
        raise SystemExit

    # Extract one XML file
    if len(xml_files) == 1:
        df, file_type, id_col, fixed_cols = parse_xml_file(
            xml_files[0]
        )

        out_path = os.path.join(
            save_folder,
            f"{file_type}_summary.xlsx"
        )

        df.to_excel(
            out_path,
            index=False
        )

        autosize_excel_columns(
            out_path
        )

        print(
            f"Extracted single file to: {out_path}"
        )

        raise SystemExit

    # Compare two XML files
    before_path = xml_files[0]
    after_path = xml_files[1]

    (
        df_before,
        before_type,
        before_id_col,
        before_fixed_cols
    ) = parse_xml_file(before_path)

    (
        df_after,
        after_type,
        after_id_col,
        after_fixed_cols
    ) = parse_xml_file(after_path)

    if before_type != after_type:
        raise ValueError(
            "The BEFORE and AFTER XML files are different XML types."
        )

    id_col = before_id_col
    fixed_cols = before_fixed_cols

    (
        removed_df,
        added_df,
        changed_df
    ) = compare_files(
        df_before,
        df_after,
        id_col
    )

    # Save comparison report
    report_path = os.path.join(
        save_folder,
        "comparison_report.xlsx"
    )

    with pd.ExcelWriter(
        report_path,
        engine="openpyxl"
    ) as writer:

        removed_df.to_excel(
            writer,
            sheet_name="Removed",
            index=False
        )

        added_df.to_excel(
            writer,
            sheet_name="Added",
            index=False
        )

        changed_df.to_excel(
            writer,
            sheet_name="Changed",
            index=False
        )

        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]

            for column_cells in worksheet.columns:
                length = max(
                    len(str(cell.value or ""))
                    for cell in column_cells
                )

                worksheet.column_dimensions[
                    column_cells[0].column_letter
                ].width = length + 2

    # Create output filenames
    before_base = os.path.splitext(
        os.path.basename(before_path)
    )[0]

    after_base = os.path.splitext(
        os.path.basename(after_path)
    )[0]

    before_out = os.path.join(
        save_folder,
        f"BEFORE_{before_base}_highlighted.xlsx"
    )

    after_out = os.path.join(
        save_folder,
        f"AFTER_{after_base}_highlighted.xlsx"
    )

    # Align columns
    df_before, df_after = align_columns_for_output(
        df_before,
        df_after,
        fixed_cols
    )

    df_before.to_excel(
        before_out,
        index=False
    )

    df_after.to_excel(
        after_out,
        index=False
    )

    # Highlight removed and added rows
    highlight_rows(
        before_out,
        removed_df,
        gray_fill,
        id_col
    )

    highlight_rows(
        after_out,
        added_df,
        green_fill,
        id_col
    )

    # Highlight changed rows and cells
    highlight_changes(
        before_out,
        changed_df,
        id_col
    )

    highlight_changes(
        after_out,
        changed_df,
        id_col
    )

    autosize_excel_columns(
        before_out
    )

    autosize_excel_columns(
        after_out
    )

    print(
        f"Comparison report saved to: {report_path}"
    )

    print(
        f"Highlighted BEFORE saved to: {before_out}"
    )

    print(
        f"Highlighted AFTER saved to: {after_out}"
    )


if __name__ == "__main__":
    main()