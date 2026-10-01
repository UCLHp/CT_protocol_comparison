import os
import re
import html
import json
from string import Template

from lxml import etree


# ============================================================
# XML DISPLAY
# ============================================================

def read_xml_text(path):

    if not path:
        return ""

    try:
        parser = etree.XMLParser(
            remove_blank_text=True,
            remove_comments=False,
            huge_tree=True,
        )

        tree = etree.parse(
            path,
            parser,
        )

        data = etree.tostring(
            tree,
            pretty_print=True,
            encoding="utf-8",
            xml_declaration=True,
        )

        return data.decode("utf-8")

    except Exception:

        with open(path, "rb") as f:
            data = f.read()

        match = re.search(
            br'<\?xml[^>]*encoding=["\']([^"\']+)["\']',
            data[:500],
            re.IGNORECASE,
        )

        encoding = (
            match.group(1).decode("ascii")
            if match
            else "utf-8"
        )

        try:
            return data.decode(encoding)

        except (
            UnicodeDecodeError,
            LookupError,
        ):
            return data.decode(
                "utf-8",
                errors="replace",
            )


def safe_filename(name):

    name = (
        name
        .replace("\\", "__")
        .replace("/", "__")
    )

    return re.sub(
        r"[^A-Za-z0-9._-]",
        "_",
        name,
    )


# ============================================================
# HIGHLIGHT INFORMATION
# ============================================================

def set_status(
    status_map,
    line,
    status,
):

    if not line:
        return

    priority = {
        "Changed": 3,
        "Added": 2,
        "Removed": 2,
    }

    current = status_map.get(line)

    if (
        current is None
        or
        priority.get(status, 0)
        >
        priority.get(current, 0)
    ):
        status_map[line] = status


def make_line_statuses(
    differences,
    old_text,
    new_text,
    whole_file_status=None,
):

    old_status = {}
    new_status = {}

    if whole_file_status == "Removed":

        for number in range(
            1,
            len(old_text.splitlines()) + 1,
        ):
            old_status[number] = "Removed"

    elif whole_file_status == "Added":

        for number in range(
            1,
            len(new_text.splitlines()) + 1,
        ):
            new_status[number] = "Added"

    else:

        for row in differences:

            status = row["Status"]

            if status == "Changed":

                set_status(
                    old_status,
                    row.get("Old Line"),
                    "Changed",
                )

                set_status(
                    new_status,
                    row.get("New Line"),
                    "Changed",
                )

            elif status == "Removed":

                set_status(
                    old_status,
                    row.get("Old Line"),
                    "Removed",
                )

            elif status == "Added":

                set_status(
                    new_status,
                    row.get("New Line"),
                    "Added",
                )

    return (
        old_status,
        new_status,
    )


# ============================================================
# XML LINE RENDERING
# ============================================================

def render_lines(
    text,
    statuses,
    side,
):

    if not text:

        return (
            '<div class="missing-file">'
            'File not present'
            '</div>'
        )

    output = []

    for number, line in enumerate(
        text.splitlines(),
        start=1,
    ):

        status = statuses.get(number)

        css_class = "line"

        if status:
            css_class += f" {status.lower()}"

        escaped = html.escape(
            line.expandtabs(4),
            quote=False,
        )

        output.append(
            f'<div '
            f'class="{css_class}" '
            f'id="{side}-L{number}">'
            f'<span class="line-number">'
            f'{number}'
            f'</span>'
            f'<span class="xml-text">'
            f'{escaped}'
            f'</span>'
            f'</div>'
        )

    return "\n".join(output)


# ============================================================
# FILE REPORT TEMPLATE
# ============================================================

FILE_TEMPLATE = Template(
r"""<!DOCTYPE html>

<html>

<head>

<meta charset="utf-8">

<title>$title</title>

<style>

* {
    box-sizing: border-box;
}

html,
body {
    width: 100%;
    height: 100%;
    margin: 0;
    overflow: hidden;
}

body {
    display: flex;
    flex-direction: column;

    font-family: Arial, sans-serif;

    color: #222;
    background: #f5f5f5;
}


/* =========================================================
   HEADER
   ========================================================= */

header {
    flex: 0 0 auto;

    padding: 10px 16px;

    background: white;

    border-bottom: 1px solid #bbb;

    z-index: 100;
}

.top-row {
    display: flex;
    align-items: center;

    gap: 12px;

    margin-bottom: 7px;
}

h1 {
    margin: 0;

    font-size: 18px;
}

.back-button {
    display: inline-block;

    padding: 6px 10px;

    border: 1px solid #aaa;
    border-radius: 3px;

    background: #eee;

    color: #222;

    text-decoration: none;

    font-size: 13px;
}

.back-button:hover {
    background: #ddd;
}


/* =========================================================
   SUMMARY
   ========================================================= */

.summary {
    font-size: 13px;
}

.legend {
    display: inline-flex;

    gap: 8px;

    margin-left: 15px;
}

.legend span {
    padding: 2px 7px;

    border-radius: 3px;
}

.legend-changed {
    background: #fff3a6;
}

.legend-added {
    background: #c9f5cf;
}

.legend-removed {
    background: #f7c7c7;
}


/* =========================================================
   CONTROLS
   ========================================================= */

.controls {
    display: flex;
    align-items: center;
    flex-wrap: wrap;

    gap: 9px;

    margin-top: 9px;
}

button {
    padding: 5px 10px;

    cursor: pointer;
}

#changeInfo {
    margin-left: 8px;

    max-width: 900px;

    overflow: hidden;

    text-overflow: ellipsis;
    white-space: nowrap;

    font-size: 12px;
}


/* =========================================================
   TWO XML PANES
   ========================================================= */

.viewer {
    flex: 1 1 auto;

    min-height: 0;

    display: grid;

    grid-template-columns:
        minmax(0, 1fr)
        minmax(0, 1fr);

    overflow: hidden;
}

.panel {
    min-width: 0;
    min-height: 0;

    display: flex;
    flex-direction: column;

    overflow: hidden;
}

.panel:first-child {
    border-right: 1px solid #aaa;
}

.panel-title {
    flex: 0 0 auto;

    padding: 7px 10px;

    background: #e8e8e8;

    border-bottom: 1px solid #bbb;

    font-size: 13px;
    font-weight: bold;
}

.code {
    flex: 1 1 auto;

    min-width: 0;
    min-height: 0;

    overflow: auto;

    background: white;

    font-family:
        Consolas,
        "Courier New",
        monospace;

    font-size: 12px;
    line-height: 1.5;

    contain: layout paint;
}


/* =========================================================
   XML LINES
   ========================================================= */

.line {
    display: flex;

    width: max-content;
    min-width: 100%;

    min-height: 18px;

    white-space: pre;
}

.line-number {
    width: 62px;
    min-width: 62px;

    padding-right: 10px;

    text-align: right;

    color: #888;

    border-right: 1px solid #eee;

    user-select: none;
}

.xml-text {
    padding-left: 10px;
    padding-right: 20px;
}


/* =========================================================
   DIFFERENCE COLOURS
   ========================================================= */

.changed {
    background: #fff3a6;
}

.added {
    background: #c9f5cf;
}

.removed {
    background: #f7c7c7;
}

.focused {
    outline: 2px solid #555;

    outline-offset: -2px;
}


/* =========================================================
   HIDE UNCHANGED
   ========================================================= */

.hide-unchanged
.line:not(.changed):not(.added):not(.removed) {
    display: none;
}


.missing-file {
    padding: 20px;

    color: #777;

    font-style: italic;
}

</style>

</head>


<body>


<header>

<div class="top-row">

<a
    class="back-button"
    href="$back_href"
>
← Back to main report
</a>

<h1>
$title
</h1>

</div>


<div class="summary">

Changed: $changed

&nbsp;&nbsp;

Added: $added

&nbsp;&nbsp;

Removed: $removed


<span class="legend">

<span class="legend-changed">
Changed
</span>

<span class="legend-added">
Added
</span>

<span class="legend-removed">
Removed
</span>

</span>

</div>


<div class="controls">

<button
    type="button"
    onclick="previousChange()"
>
Previous change
</button>


<button
    type="button"
    onclick="nextChange()"
>
Next change
</button>


<label>

<input
    type="checkbox"
    id="syncScroll"
    checked
>

Synchronise scrolling

</label>


<label>

<input
    type="checkbox"
    id="showUnchanged"
    checked
>

Show unchanged lines

</label>


<span id="changeInfo"></span>

</div>

</header>


<div class="viewer">


<div class="panel">

<div class="panel-title">
OLD
</div>

<div
    class="code"
    id="oldPane"
>
$old_html
</div>

</div>


<div class="panel">

<div class="panel-title">
NEW
</div>

<div
    class="code"
    id="newPane"
>
$new_html
</div>

</div>


</div>


<script>

const changes = $changes_json;

const oldLineCount = $old_line_count;
const newLineCount = $new_line_count;

let currentChange = -1;
let synchronising = false;

let focusedOld = null;
let focusedNew = null;


const oldPane =
    document.getElementById(
        "oldPane"
    );

const newPane =
    document.getElementById(
        "newPane"
    );

const syncScrollCheckbox =
    document.getElementById(
        "syncScroll"
    );

const showUnchangedCheckbox =
    document.getElementById(
        "showUnchanged"
    );

const changeInfo =
    document.getElementById(
        "changeInfo"
    );


/* =========================================================
   MANUAL SCROLL SYNCHRONISATION
   ========================================================= */

function syncPane(
    source,
    target
) {

    if (
        !syncScrollCheckbox.checked
        ||
        synchronising
    ) {
        return;
    }


    const sourceRange =
        source.scrollHeight
        -
        source.clientHeight;

    const targetRange =
        target.scrollHeight
        -
        target.clientHeight;


    if (
        sourceRange <= 0
        ||
        targetRange <= 0
    ) {
        return;
    }


    synchronising = true;


    target.scrollTop =
        (
            source.scrollTop
            /
            sourceRange
        )
        *
        targetRange;


    requestAnimationFrame(
        () => {
            synchronising = false;
        }
    );
}


oldPane.addEventListener(
    "scroll",
    () => {

        syncPane(
            oldPane,
            newPane
        );

    },
    { passive: true }
);


newPane.addEventListener(
    "scroll",
    () => {

        syncPane(
            newPane,
            oldPane
        );

    },
    { passive: true }
);


syncScrollCheckbox.addEventListener(
    "change",
    () => {

        if (
            syncScrollCheckbox.checked
        ) {

            syncPane(
                oldPane,
                newPane
            );
        }

    }
);


/* =========================================================
   SHOW / HIDE UNCHANGED
   ========================================================= */

showUnchangedCheckbox.addEventListener(
    "change",
    () => {

        document.body.classList.toggle(
            "hide-unchanged",
            !showUnchangedCheckbox.checked
        );

    }
);


/* =========================================================
   FOCUS
   ========================================================= */

function clearFocus() {

    if (focusedOld) {

        focusedOld.classList.remove(
            "focused"
        );

        focusedOld = null;
    }


    if (focusedNew) {

        focusedNew.classList.remove(
            "focused"
        );

        focusedNew = null;
    }
}


/* =========================================================
   SCROLL TO LINE
   ========================================================= */

function centreLine(
    pane,
    element
) {

    if (!element) {
        return;
    }


    const target =
        element.offsetTop
        -
        (
            pane.clientHeight
            /
            2
        )
        +
        (
            element.offsetHeight
            /
            2
        );


    pane.scrollTop =
        Math.max(
            0,
            target
        );
}


/* =========================================================
   APPROXIMATE MATCHING LINE

   Used when an Added line exists only in NEW,
   or Removed line exists only in OLD.

   We map its relative line position into the other file.

   Example:

       NEW line 900 / 1800 lines = 50%

   so OLD is positioned around:

       50% of OLD line count.

   ========================================================= */

function correspondingLine(
    sourceLine,
    sourceCount,
    targetCount
) {

    if (
        !sourceLine
        ||
        sourceCount <= 1
        ||
        targetCount <= 0
    ) {
        return null;
    }


    const fraction =
        (
            sourceLine - 1
        )
        /
        (
            sourceCount - 1
        );


    const targetLine =
        Math.round(
            fraction
            *
            Math.max(
                0,
                targetCount - 1
            )
        )
        +
        1;


    return Math.max(
        1,
        Math.min(
            targetCount,
            targetLine
        )
    );
}


function getLineElement(
    side,
    lineNumber
) {

    if (!lineNumber) {
        return null;
    }


    return document.getElementById(
        side
        +
        "-L"
        +
        lineNumber
    );
}


/* =========================================================
   CHANGE NAVIGATION
   ========================================================= */

function goToChange(index) {

    if (!changes.length) {
        return;
    }


    currentChange =
        (
            index
            +
            changes.length
        )
        %
        changes.length;


    const change =
        changes[
            currentChange
        ];


    clearFocus();


    let oldElement =
        getLineElement(
            "old",
            change.old
        );


    let newElement =
        getLineElement(
            "new",
            change.new
        );


    synchronising = true;


    /* -----------------------------------------------------
       CHANGED
       Exact line exists in both files.
       ----------------------------------------------------- */

    if (
        oldElement
        &&
        newElement
    ) {

        focusedOld = oldElement;
        focusedNew = newElement;


        oldElement.classList.add(
            "focused"
        );

        newElement.classList.add(
            "focused"
        );


        centreLine(
            oldPane,
            oldElement
        );

        centreLine(
            newPane,
            newElement
        );
    }


    /* -----------------------------------------------------
       ADDED
       Exact line exists only in NEW.

       NEW goes to exact line.

       OLD goes to approximately corresponding position,
       but only if sync is enabled.
       ----------------------------------------------------- */

    else if (
        newElement
        &&
        !oldElement
    ) {

        focusedNew = newElement;


        newElement.classList.add(
            "focused"
        );


        centreLine(
            newPane,
            newElement
        );


        if (
            syncScrollCheckbox.checked
        ) {

            const estimatedOldLine =
                correspondingLine(
                    change.new,
                    newLineCount,
                    oldLineCount
                );


            const estimatedOldElement =
                getLineElement(
                    "old",
                    estimatedOldLine
                );


            if (estimatedOldElement) {

                centreLine(
                    oldPane,
                    estimatedOldElement
                );
            }
        }
    }


    /* -----------------------------------------------------
       REMOVED
       Exact line exists only in OLD.

       OLD goes to exact line.

       NEW goes to approximately corresponding position,
       but only if sync is enabled.
       ----------------------------------------------------- */

    else if (
        oldElement
        &&
        !newElement
    ) {

        focusedOld = oldElement;


        oldElement.classList.add(
            "focused"
        );


        centreLine(
            oldPane,
            oldElement
        );


        if (
            syncScrollCheckbox.checked
        ) {

            const estimatedNewLine =
                correspondingLine(
                    change.old,
                    oldLineCount,
                    newLineCount
                );


            const estimatedNewElement =
                getLineElement(
                    "new",
                    estimatedNewLine
                );


            if (estimatedNewElement) {

                centreLine(
                    newPane,
                    estimatedNewElement
                );
            }
        }
    }


    requestAnimationFrame(
        () => {

            synchronising = false;

        }
    );


    changeInfo.textContent =
        (currentChange + 1)
        +
        " / "
        +
        changes.length
        +
        "   "
        +
        change.status
        +
        "   "
        +
        change.key;
}


function nextChange() {

    goToChange(
        currentChange + 1
    );
}


function previousChange() {

    goToChange(
        currentChange - 1
    );
}

</script>


</body>

</html>
"""
)


# ============================================================
# INDIVIDUAL FILE REPORT
# ============================================================

def file_report_html(
    item,
    back_href,
):

    filename = item["file"]

    old_text = item.get(
        "old_text"
    )

    new_text = item.get(
        "new_text"
    )


    if old_text is None:

        old_text = read_xml_text(
            item.get("old_path")
        )


    if new_text is None:

        new_text = read_xml_text(
            item.get("new_path")
        )


    differences = item[
        "differences"
    ]


    old_line_count = len(
        old_text.splitlines()
    )

    new_line_count = len(
        new_text.splitlines()
    )


    (
        old_status,
        new_status,
    ) = make_line_statuses(
        differences,
        old_text,
        new_text,
        item.get(
            "whole_file_status"
        ),
    )


    old_html = render_lines(
        old_text,
        old_status,
        "old",
    )


    new_html = render_lines(
        new_text,
        new_status,
        "new",
    )


    changed = sum(
        row["Status"] == "Changed"
        for row in differences
    )


    added = sum(
        row["Status"] == "Added"
        for row in differences
    )


    removed = sum(
        row["Status"] == "Removed"
        for row in differences
    )


    # --------------------------------------------------------
    # Build navigation entries
    # --------------------------------------------------------

    changes = []

    for row in differences:

        old_line = row.get(
            "Old Line"
        )

        new_line = row.get(
            "New Line"
        )


        # Navigation position:
        #
        # Prefer NEW file position where available.
        # Removed-only items use OLD file position.
        if new_line is not None:

            position = new_line

        elif old_line is not None:

            position = old_line

        else:

            position = float("inf")


        changes.append({
            "status":
                row["Status"],

            "key":
                row.get(
                    "Key",
                    ""
                ),

            "old":
                old_line,

            "new":
                new_line,

            "_position":
                position,
        })


    # --------------------------------------------------------
    # THIS fixes the strange up/down navigation.
    #
    # Next Change now follows XML position top -> bottom.
    # --------------------------------------------------------

    changes.sort(
        key=lambda change: (
            change["_position"],
            change["status"],
            change["key"],
        )
    )


    # Internal sort field is not needed in JavaScript.
    for change in changes:

        change.pop(
            "_position",
            None,
        )


    changes_json = (
        json.dumps(
            changes
        )
        .replace(
            "</",
            "<\\/"
        )
    )


    return FILE_TEMPLATE.substitute(

        title=html.escape(
            filename
        ),

        back_href=html.escape(
            back_href,
            quote=True,
        ),

        changed=changed,

        added=added,

        removed=removed,

        old_html=old_html,

        new_html=new_html,

        old_line_count=old_line_count,

        new_line_count=new_line_count,

        changes_json=changes_json,
    )


# ============================================================
# MASTER REPORT
# ============================================================

MASTER_TEMPLATE = Template(
r"""<!DOCTYPE html>

<html>

<head>

<meta charset="utf-8">

<title>
XML Comparison Report
</title>

<style>

body {
    font-family: Arial, sans-serif;

    margin: 30px;

    color: #222;
}

h1 {
    font-size: 22px;
}

table {
    border-collapse: collapse;

    min-width: 800px;
}

th,
td {
    border: 1px solid #ccc;

    padding: 8px 11px;

    text-align: left;
}

th {
    background: #eee;
}

tr:hover {
    background: #f5f5f5;
}

a {
    text-decoration: none;
}

</style>

</head>


<body>

<h1>
XML Comparison Report
</h1>


<table>

<thead>

<tr>

<th>
File
</th>

<th>
Changed
</th>

<th>
Added
</th>

<th>
Removed
</th>

<th>
Total
</th>

</tr>

</thead>


<tbody>

$rows

</tbody>

</table>

</body>

</html>
"""
)


def master_report_html(
    entries,
):

    rows = []

    for entry in entries:

        rows.append(

            "<tr>"

            f'<td>'
            f'<a href="'
            f'{html.escape(entry["href"], quote=True)}'
            f'">'
            f'{html.escape(entry["file"])}'
            f'</a>'
            f'</td>'

            f'<td>'
            f'{entry["changed"]}'
            f'</td>'

            f'<td>'
            f'{entry["added"]}'
            f'</td>'

            f'<td>'
            f'{entry["removed"]}'
            f'</td>'

            f'<td>'
            f'{entry["total"]}'
            f'</td>'

            "</tr>"
        )


    return MASTER_TEMPLATE.substitute(
        rows="\n".join(
            rows
        )
    )


# ============================================================
# GENERATE ALL HTML REPORTS
# ============================================================

def generate_html_reports(
    items,
    output_folder,
    report_stem,
):

    report_folder_name = (
        f"{report_stem}_files"
    )


    report_folder = os.path.join(
        output_folder,
        report_folder_name,
    )


    os.makedirs(
        report_folder,
        exist_ok=True,
    )


    master_filename = (
        f"{report_stem}.html"
    )


    back_href = (
        f"../{master_filename}"
    )


    entries = []


    for index, item in enumerate(
        items,
        start=1,
    ):

        differences = item[
            "differences"
        ]


        if not differences:
            continue


        filename = (
            f"{index:03d}_"
            f"{safe_filename(item['file'])}"
            ".html"
        )


        filepath = os.path.join(
            report_folder,
            filename,
        )


        with open(
            filepath,
            "w",
            encoding="utf-8",
        ) as file:

            file.write(
                file_report_html(
                    item,
                    back_href,
                )
            )


        changed = sum(
            row["Status"] == "Changed"
            for row in differences
        )


        added = sum(
            row["Status"] == "Added"
            for row in differences
        )


        removed = sum(
            row["Status"] == "Removed"
            for row in differences
        )


        entries.append({

            "file":
                item["file"],

            "changed":
                changed,

            "added":
                added,

            "removed":
                removed,

            "total":
                len(
                    differences
                ),

            "href":
                (
                    f"{report_folder_name}/"
                    f"{filename}"
                ),
        })


    master_path = os.path.join(
        output_folder,
        master_filename,
    )


    with open(
        master_path,
        "w",
        encoding="utf-8",
    ) as file:

        file.write(
            master_report_html(
                entries
            )
        )


    return master_path