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

        tree = etree.parse(path, parser)

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
# LINE STATUS
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

    return old_status, new_status


# ============================================================
# WIDTH ESTIMATION
#
# Keeps horizontal scrollbar width stable even though only a
# small virtual section of the XML is actually rendered.
# ============================================================

def estimate_width_px(lines):
    if not lines:
        return 900

    max_chars = 0

    for line in lines:
        length = len(line.expandtabs(4))

        if length > max_chars:
            max_chars = length

    width = int(
        90
        +
        max_chars * 7.3
        +
        30
    )

    return max(
        900,
        width,
    )


# ============================================================
# FILE TEMPLATE
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
    gap: 10px;
    margin-top: 9px;
}

button {
    padding: 5px 10px;
    cursor: pointer;
}

#changeInfo {
    margin-left: 5px;
    max-width: 850px;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    font-size: 12px;
}

#syncStatus {
    color: #777;
    font-size: 11px;
}


/* =========================================================
   VIEWER
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

.range-note {
    margin-left: 8px;
    color: #666;
    font-weight: normal;
}


/* =========================================================
   VIRTUAL DOCUMENT

   Scrollbar represents the ENTIRE XML file.

   Only the currently visible area plus a buffer is placed
   into the DOM.
   ========================================================= */

.code {
    position: relative;

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
    line-height: 18px;

    contain: layout paint;

    outline: none;
}

.scroll-space {
    position: relative;
    min-width: 100%;
}

.virtual-lines {
    position: absolute;
    top: 0;
    left: 0;
    will-change: transform;
}

.line {
    display: flex;
    height: 18px;
    white-space: pre;
}

.line-number {
    flex: 0 0 64px;
    width: 64px;

    padding-right: 10px;

    text-align: right;

    color: #888;

    border-right: 1px solid #eee;

    user-select: none;
}

.xml-text {
    flex: 1 0 auto;

    padding-left: 10px;
    padding-right: 25px;
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
    outline: 2px solid #444;
    outline-offset: -2px;
}

.empty-message {
    padding: 18px;
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
← Previous change
</button>


<button
    type="button"
    onclick="nextChange()"
>
Next change →
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


<span id="syncStatus">
Sync on
</span>


<span id="changeInfo"></span>

</div>

</header>


<div class="viewer">


<div class="panel">

<div class="panel-title">

OLD

<span
    class="range-note"
    id="oldRange"
></span>

</div>


<div
    class="code"
    id="oldPane"
    tabindex="0"
>

<div
    class="scroll-space"
    id="oldSpace"
>

<div
    class="virtual-lines"
    id="oldVirtual"
></div>

</div>

</div>

</div>


<div class="panel">

<div class="panel-title">

NEW

<span
    class="range-note"
    id="newRange"
></span>

</div>


<div
    class="code"
    id="newPane"
    tabindex="0"
>

<div
    class="scroll-space"
    id="newSpace"
>

<div
    class="virtual-lines"
    id="newVirtual"
></div>

</div>

</div>

</div>


</div>


<script>


/* =========================================================
   CONSTANTS
   ========================================================= */

const LINE_HEIGHT = 18;

/*
Only this many lines above/below the viewport are rendered.
Even a 50,000-line XML therefore remains lightweight.
*/
const BUFFER_LINES = 150;


/* =========================================================
   SOURCE DATA
   ========================================================= */

const changes =
    $changes_json;

const oldLines =
    $old_lines_json;

const newLines =
    $new_lines_json;

const oldStatuses =
    $old_status_json;

const newStatuses =
    $new_status_json;

const oldWidth =
    $old_width;

const newWidth =
    $new_width;


/* =========================================================
   GLOBAL STATE
   ========================================================= */

let currentChange = -1;

let selectingText = false;

let pendingSyncFrame = null;

let pendingSyncSource = null;

let pendingSyncTarget = null;


/* =========================================================
   ELEMENTS
   ========================================================= */

const syncScrollCheckbox =
    document.getElementById(
        "syncScroll"
    );

const showUnchangedCheckbox =
    document.getElementById(
        "showUnchanged"
    );

const syncStatus =
    document.getElementById(
        "syncStatus"
    );

const changeInfo =
    document.getElementById(
        "changeInfo"
    );


/* =========================================================
   CHANGED LINE LISTS

   Used when Show unchanged lines is OFF.
   ========================================================= */

function sortedStatusLines(
    statuses
) {

    return Object.keys(
        statuses
    )
    .map(
        Number
    )
    .sort(
        function(a, b) {
            return a - b;
        }
    );
}


const oldChangedLines =
    sortedStatusLines(
        oldStatuses
    );

const newChangedLines =
    sortedStatusLines(
        newStatuses
    );


/* =========================================================
   PANE STATE
   ========================================================= */

const oldState = {

    side: "old",

    pane:
        document.getElementById(
            "oldPane"
        ),

    space:
        document.getElementById(
            "oldSpace"
        ),

    virtual:
        document.getElementById(
            "oldVirtual"
        ),

    range:
        document.getElementById(
            "oldRange"
        ),

    lines:
        oldLines,

    statuses:
        oldStatuses,

    changedLines:
        oldChangedLines,

    width:
        oldWidth,

    renderStart:
        -1,

    renderEnd:
        -1,

    lastFocus:
        null,

    suppressSync:
        false,
};


const newState = {

    side: "new",

    pane:
        document.getElementById(
            "newPane"
        ),

    space:
        document.getElementById(
            "newSpace"
        ),

    virtual:
        document.getElementById(
            "newVirtual"
        ),

    range:
        document.getElementById(
            "newRange"
        ),

    lines:
        newLines,

    statuses:
        newStatuses,

    changedLines:
        newChangedLines,

    width:
        newWidth,

    renderStart:
        -1,

    renderEnd:
        -1,

    lastFocus:
        null,

    suppressSync:
        false,
};


/* =========================================================
   DISPLAY MODE
   ========================================================= */

function showingUnchanged() {

    return (
        showUnchangedCheckbox.checked
    );
}


function displayCount(
    state
) {

    if (
        showingUnchanged()
    ) {

        return state.lines.length;
    }

    return state.changedLines.length;
}


function actualLineAtDisplayIndex(
    state,
    index
) {

    if (
        showingUnchanged()
    ) {

        return index + 1;
    }

    return state.changedLines[
        index
    ];
}


/* =========================================================
   BINARY SEARCH

   Used when unchanged lines are hidden.
   ========================================================= */

function nearestChangedIndex(
    changedLines,
    actualLine
) {

    if (
        !changedLines.length
    ) {

        return 0;
    }


    let low = 0;

    let high =
        changedLines.length - 1;


    while (
        low <= high
    ) {

        const middle =
            Math.floor(
                (
                    low
                    +
                    high
                )
                /
                2
            );


        const value =
            changedLines[
                middle
            ];


        if (
            value === actualLine
        ) {

            return middle;
        }


        if (
            value < actualLine
        ) {

            low =
                middle + 1;
        }

        else {

            high =
                middle - 1;
        }
    }


    if (
        low <= 0
    ) {

        return 0;
    }


    if (
        low >=
        changedLines.length
    ) {

        return (
            changedLines.length - 1
        );
    }


    const before =
        changedLines[
            low - 1
        ];


    const after =
        changedLines[
            low
        ];


    if (
        Math.abs(
            actualLine
            -
            before
        )
        <=
        Math.abs(
            after
            -
            actualLine
        )
    ) {

        return low - 1;
    }


    return low;
}


function displayIndexForActualLine(
    state,
    actualLine
) {

    if (
        !actualLine
    ) {

        return 0;
    }


    if (
        showingUnchanged()
    ) {

        return Math.max(
            0,

            Math.min(
                state.lines.length - 1,
                actualLine - 1
            )
        );
    }


    return nearestChangedIndex(
        state.changedLines,
        actualLine
    );
}


/* =========================================================
   CURRENT FOCUS LINE
   ========================================================= */

function focusLineForState(
    state
) {

    if (
        currentChange < 0
        ||
        currentChange >= changes.length
    ) {

        return null;
    }


    const change =
        changes[
            currentChange
        ];


    if (
        state.side === "old"
    ) {

        return change.old;
    }


    return change.new;
}


/* =========================================================
   SCROLL SPACE SIZE
   ========================================================= */

function configureScrollSpace(
    state
) {

    const count =
        displayCount(
            state
        );


    const height =
        Math.max(
            state.pane.clientHeight,
            count * LINE_HEIGHT
        );


    const width =
        Math.max(
            state.pane.clientWidth,
            state.width
        );


    state.space.style.height =
        height
        +
        "px";


    state.space.style.width =
        width
        +
        "px";


    state.virtual.style.width =
        width
        +
        "px";
}


/* =========================================================
   RANGE LABEL
   ========================================================= */

function updateRangeLabel(
    state
) {

    const count =
        displayCount(
            state
        );


    if (
        count <= 0
    ) {

        state.range.textContent =
            "no reportable lines";

        return;
    }


    const firstIndex =
        Math.max(
            0,

            Math.min(
                count - 1,

                Math.floor(
                    state.pane.scrollTop
                    /
                    LINE_HEIGHT
                )
            )
        );


    const visibleRows =
        Math.max(
            1,

            Math.ceil(
                state.pane.clientHeight
                /
                LINE_HEIGHT
            )
        );


    const lastIndex =
        Math.min(
            count - 1,
            firstIndex + visibleRows - 1
        );


    const firstActual =
        actualLineAtDisplayIndex(
            state,
            firstIndex
        );


    const lastActual =
        actualLineAtDisplayIndex(
            state,
            lastIndex
        );


    if (
        showingUnchanged()
    ) {

        state.range.textContent =
            "lines "
            +
            firstActual
            +
            "–"
            +
            lastActual
            +
            " of "
            +
            state.lines.length;
    }

    else {

        state.range.textContent =
            "changes only · "
            +
            count
            +
            " lines";
    }
}


/* =========================================================
   VIRTUAL RENDER
   ========================================================= */

function renderVirtual(
    state,
    force
) {

    /*
    Do not replace DOM while the user is actively selecting
    XML text.
    */
    if (
        selectingText
    ) {

        return;
    }


    configureScrollSpace(
        state
    );


    const count =
        displayCount(
            state
        );


    updateRangeLabel(
        state
    );


    if (
        count <= 0
    ) {

        state.virtual.replaceChildren();


        const message =
            document.createElement(
                "div"
            );


        message.className =
            "empty-message";


        message.textContent =
            "No reportable lines in this pane.";


        state.virtual.appendChild(
            message
        );


        state.virtual.style.transform =
            "translateY(0px)";


        state.renderStart =
            0;

        state.renderEnd =
            0;

        return;
    }


    const firstVisible =
        Math.max(
            0,

            Math.floor(
                state.pane.scrollTop
                /
                LINE_HEIGHT
            )
        );


    const visibleRows =
        Math.max(
            1,

            Math.ceil(
                state.pane.clientHeight
                /
                LINE_HEIGHT
            )
        );


    const start =
        Math.max(
            0,
            firstVisible - BUFFER_LINES
        );


    const end =
        Math.min(
            count,
            firstVisible
            +
            visibleRows
            +
            BUFFER_LINES
        );


    const focusLine =
        focusLineForState(
            state
        );


    if (
        !force
        &&
        start === state.renderStart
        &&
        end === state.renderEnd
        &&
        focusLine === state.lastFocus
    ) {

        return;
    }


    state.renderStart =
        start;

    state.renderEnd =
        end;

    state.lastFocus =
        focusLine;


    const fragment =
        document.createDocumentFragment();


    for (
        let displayIndex = start;
        displayIndex < end;
        displayIndex++
    ) {

        const actualLine =
            actualLineAtDisplayIndex(
                state,
                displayIndex
            );


        if (
            !actualLine
        ) {

            continue;
        }


        const row =
            document.createElement(
                "div"
            );


        row.className =
            "line";


        row.id =
            state.side
            +
            "-L"
            +
            actualLine;


        const status =
            state.statuses[
                String(actualLine)
            ];


        if (
            status
        ) {

            row.classList.add(
                status.toLowerCase()
            );
        }


        if (
            actualLine === focusLine
        ) {

            row.classList.add(
                "focused"
            );
        }


        const number =
            document.createElement(
                "span"
            );


        number.className =
            "line-number";


        number.textContent =
            actualLine;


        const text =
            document.createElement(
                "span"
            );


        text.className =
            "xml-text";


        text.textContent =
            state.lines[
                actualLine - 1
            ];


        row.appendChild(
            number
        );


        row.appendChild(
            text
        );


        fragment.appendChild(
            row
        );
    }


    state.virtual.replaceChildren(
        fragment
    );


    state.virtual.style.transform =
        "translateY("
        +
        (
            start
            *
            LINE_HEIGHT
        )
        +
        "px)";
}


/* =========================================================
   RENDER SCHEDULING
   ========================================================= */

let oldRenderFrame = null;

let newRenderFrame = null;


function scheduleRender(
    state
) {

    if (
        selectingText
    ) {

        return;
    }


    if (
        state.side === "old"
    ) {

        if (
            oldRenderFrame !== null
        ) {

            return;
        }


        oldRenderFrame =
            requestAnimationFrame(
                function() {

                    oldRenderFrame =
                        null;


                    renderVirtual(
                        oldState,
                        false
                    );
                }
            );
    }

    else {

        if (
            newRenderFrame !== null
        ) {

            return;
        }


        newRenderFrame =
            requestAnimationFrame(
                function() {

                    newRenderFrame =
                        null;


                    renderVirtual(
                        newState,
                        false
                    );
                }
            );
    }
}


/* =========================================================
   SCROLL SYNCHRONISATION

   Handles:
       mouse wheel
       touchpad
       scrollbar drag
       scrollbar track click
       PageUp/PageDown
       keyboard scrolling
   ========================================================= */

function performPendingSync() {

    pendingSyncFrame =
        null;


    const source =
        pendingSyncSource;


    const target =
        pendingSyncTarget;


    pendingSyncSource =
        null;


    pendingSyncTarget =
        null;


    if (
        !source
        ||
        !target
        ||
        !syncScrollCheckbox.checked
        ||
        selectingText
    ) {

        return;
    }


    const sourceVerticalRange =
        Math.max(
            0,
            source.pane.scrollHeight
            -
            source.pane.clientHeight
        );


    const targetVerticalRange =
        Math.max(
            0,
            target.pane.scrollHeight
            -
            target.pane.clientHeight
        );


    const sourceHorizontalRange =
        Math.max(
            0,
            source.pane.scrollWidth
            -
            source.pane.clientWidth
        );


    const targetHorizontalRange =
        Math.max(
            0,
            target.pane.scrollWidth
            -
            target.pane.clientWidth
        );


    let targetTop = 0;


    if (
        sourceVerticalRange > 0
        &&
        targetVerticalRange > 0
    ) {

        targetTop =
            (
                source.pane.scrollTop
                /
                sourceVerticalRange
            )
            *
            targetVerticalRange;
    }


    let targetLeft = 0;


    if (
        sourceHorizontalRange > 0
        &&
        targetHorizontalRange > 0
    ) {

        targetLeft =
            (
                source.pane.scrollLeft
                /
                sourceHorizontalRange
            )
            *
            targetHorizontalRange;
    }


    target.suppressSync =
        true;


    target.pane.scrollTop =
        targetTop;


    target.pane.scrollLeft =
        targetLeft;


    scheduleRender(
        target
    );


    requestAnimationFrame(
        function() {

            target.suppressSync =
                false;
        }
    );
}


function scheduleSync(
    source,
    target
) {

    if (
        !syncScrollCheckbox.checked
        ||
        selectingText
    ) {

        return;
    }


    pendingSyncSource =
        source;


    pendingSyncTarget =
        target;


    if (
        pendingSyncFrame !== null
    ) {

        return;
    }


    pendingSyncFrame =
        requestAnimationFrame(
            performPendingSync
        );
}


/* =========================================================
   SCROLL EVENTS
   ========================================================= */

function handlePaneScroll(
    source,
    target
) {

    scheduleRender(
        source
    );


    if (
        source.suppressSync
    ) {

        return;
    }


    scheduleSync(
        source,
        target
    );
}


oldState.pane.addEventListener(
    "scroll",
    function() {

        handlePaneScroll(
            oldState,
            newState
        );
    },
    {
        passive: true
    }
);


newState.pane.addEventListener(
    "scroll",
    function() {

        handlePaneScroll(
            newState,
            oldState
        );
    },
    {
        passive: true
    }
);


/* =========================================================
   TEXT SELECTION

   Only clicking directly on XML text pauses sync.

   Clicking/dragging scrollbar DOES NOT pause sync.
   ========================================================= */

function startPossibleSelection(
    event
) {

    if (
        event.button !== 0
    ) {

        return;
    }


    const target =
        event.target;


    if (
        target
        &&
        target.classList
        &&
        target.classList.contains(
            "xml-text"
        )
    ) {

        selectingText =
            true;


        syncStatus.textContent =
            "Sync paused while selecting";
    }
}


function stopSelection() {

    if (
        !selectingText
    ) {

        return;
    }


    selectingText =
        false;


    syncStatus.textContent =
        (
            syncScrollCheckbox.checked
            ?
            "Sync on"
            :
            "Sync off"
        );


    renderVirtual(
        oldState,
        true
    );


    renderVirtual(
        newState,
        true
    );
}


oldState.pane.addEventListener(
    "mousedown",
    startPossibleSelection
);


newState.pane.addEventListener(
    "mousedown",
    startPossibleSelection
);


document.addEventListener(
    "mouseup",
    stopSelection
);


window.addEventListener(
    "blur",
    stopSelection
);


/* =========================================================
   PANE FOCUS
   ========================================================= */

oldState.pane.addEventListener(
    "mousedown",
    function() {

        oldState.pane.focus(
            {
                preventScroll: true
            }
        );
    }
);


newState.pane.addEventListener(
    "mousedown",
    function() {

        newState.pane.focus(
            {
                preventScroll: true
            }
        );
    }
);


/* =========================================================
   CORRESPONDING LINE

   Converts a line's relative position from one file into
   the other file.

   Used when a line exists only in OLD or only in NEW.
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


    return Math.max(
        1,

        Math.min(
            targetCount,

            Math.round(
                fraction
                *
                (
                    targetCount - 1
                )
            )
            +
            1
        )
    );
}


/* =========================================================
   SCROLL TO ACTUAL XML LINE
   ========================================================= */

function scrollToActualLine(
    state,
    actualLine
) {

    if (
        !actualLine
    ) {

        return;
    }


    configureScrollSpace(
        state
    );


    const count =
        displayCount(
            state
        );


    if (
        count <= 0
    ) {

        return;
    }


    const displayIndex =
        displayIndexForActualLine(
            state,
            actualLine
        );


    let top =
        (
            displayIndex
            *
            LINE_HEIGHT
        )
        -
        (
            state.pane.clientHeight
            /
            2
        )
        +
        (
            LINE_HEIGHT
            /
            2
        );


    const maxTop =
        Math.max(
            0,
            state.pane.scrollHeight
            -
            state.pane.clientHeight
        );


    top =
        Math.max(
            0,

            Math.min(
                maxTop,
                top
            )
        );


    state.suppressSync =
        true;


    state.pane.scrollTop =
        top;


    requestAnimationFrame(
        function() {

            state.suppressSync =
                false;
        }
    );
}


/* =========================================================
   CHANGE NAVIGATION

   IMPORTANT:

   The changes array is generated by Python in NEW-document
   top-to-bottom order.

   Therefore Right Arrow / Next Change moves consistently
   downward through NEW.

   Removed items have no NEW line, so Python gives them a
   position corresponding to their OLD relative location.
   ========================================================= */

function goToChange(
    index
) {

    if (
        !changes.length
    ) {

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


    let oldTarget =
        change.old;


    let newTarget =
        change.new;


    /*
    Added item:
    exact line exists only in NEW.
    */

    if (
        !oldTarget
        &&
        newTarget
    ) {

        oldTarget =
            correspondingLine(
                newTarget,
                newLines.length,
                oldLines.length
            );
    }


    /*
    Removed item:
    exact line exists only in OLD.

    Use the same NEW-side navigation location that Python
    used when sorting the change list.
    */

    if (
        !newTarget
        &&
        change.navNew
    ) {

        newTarget =
            change.navNew;
    }


    /*
    Fallback in case an older report does not contain navNew.
    */

    if (
        !newTarget
        &&
        oldTarget
    ) {

        newTarget =
            correspondingLine(
                oldTarget,
                oldLines.length,
                newLines.length
            );
    }


    scrollToActualLine(
        oldState,
        oldTarget
    );


    scrollToActualLine(
        newState,
        newTarget
    );


    renderVirtual(
        oldState,
        true
    );


    renderVirtual(
        newState,
        true
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


/* =========================================================
   SHOW / HIDE UNCHANGED
   ========================================================= */

function actualLineAtPaneCentre(
    state
) {

    const count =
        displayCount(
            state
        );


    if (
        count <= 0
    ) {

        return 1;
    }


    const centreIndex =
        Math.max(
            0,

            Math.min(
                count - 1,

                Math.floor(
                    (
                        state.pane.scrollTop
                        +
                        state.pane.clientHeight
                        /
                        2
                    )
                    /
                    LINE_HEIGHT
                )
            )
        );


    return actualLineAtDisplayIndex(
        state,
        centreIndex
    );
}


showUnchangedCheckbox.addEventListener(
    "change",
    function() {

        /*
        Preserve approximately the same XML location when
        changing between full-document and changes-only mode.
        */

        const oldCentre =
            actualLineAtPaneCentre(
                oldState
            );


        const newCentre =
            actualLineAtPaneCentre(
                newState
            );


        oldState.renderStart =
            -1;

        oldState.renderEnd =
            -1;

        newState.renderStart =
            -1;

        newState.renderEnd =
            -1;


        configureScrollSpace(
            oldState
        );


        configureScrollSpace(
            newState
        );


        scrollToActualLine(
            oldState,
            oldCentre
        );


        scrollToActualLine(
            newState,
            newCentre
        );


        renderVirtual(
            oldState,
            true
        );


        renderVirtual(
            newState,
            true
        );
    }
);


/* =========================================================
   SYNC CHECKBOX
   ========================================================= */

syncScrollCheckbox.addEventListener(
    "change",
    function() {

        if (
            syncScrollCheckbox.checked
        ) {

            syncStatus.textContent =
                "Sync on";


            scheduleSync(
                oldState,
                newState
            );
        }

        else {

            syncStatus.textContent =
                "Sync off";
        }
    }
);


/* =========================================================
   KEYBOARD

   ← Previous change
   → Next change

   ↑ ↓ PageUp PageDown remain ordinary scrolling.
   ========================================================= */

document.addEventListener(
    "keydown",
    function(event) {

        if (
            event.ctrlKey
            ||
            event.altKey
            ||
            event.metaKey
        ) {

            return;
        }


        if (
            event.key === "ArrowRight"
        ) {

            event.preventDefault();

            nextChange();
        }


        else if (
            event.key === "ArrowLeft"
        ) {

            event.preventDefault();

            previousChange();
        }
    }
);


/* =========================================================
   RESIZE
   ========================================================= */

window.addEventListener(
    "resize",
    function() {

        oldState.renderStart =
            -1;

        oldState.renderEnd =
            -1;

        newState.renderStart =
            -1;

        newState.renderEnd =
            -1;


        renderVirtual(
            oldState,
            true
        );


        renderVirtual(
            newState,
            true
        );
    }
);


/* =========================================================
   INITIAL DISPLAY
   ========================================================= */

configureScrollSpace(
    oldState
);


configureScrollSpace(
    newState
);


renderVirtual(
    oldState,
    true
);


renderVirtual(
    newState,
    true
);


if (
    changes.length
) {

    goToChange(0);
}

</script>


</body>

</html>
"""
)


# ============================================================
# BUILD FILE REPORT
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

    old_lines = (
        old_text.splitlines()
    )

    new_lines = (
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

    old_line_count = len(
        old_lines
    )

    new_line_count = len(
        new_lines
    )


    # ========================================================
    # NEW-DOCUMENT NAVIGATION POSITION
    #
    # NEW is the definitive navigation coordinate system.
    #
    # Changed:
    #     use actual New Line.
    #
    # Added:
    #     use actual New Line.
    #
    # Removed:
    #     no New Line exists, so map the OLD position into
    #     the NEW document.
    #
    # This prevents Next Change from jumping down/up/down.
    # ========================================================

    def estimated_new_line_from_old(
        old_line,
    ):
        if old_line is None:
            return None

        if new_line_count <= 0:
            return None

        if old_line_count <= 1:
            return 1.0

        relative_position = (
            old_line - 1
        ) / (
            old_line_count - 1
        )

        estimated_new_line = (
            relative_position
            *
            max(
                0,
                new_line_count - 1
            )
            +
            1
        )

        return estimated_new_line


    def navigation_position(
        row,
    ):
        new_line = row.get(
            "New Line"
        )

        old_line = row.get(
            "Old Line"
        )

        # Actual NEW line always wins.
        if new_line is not None:
            return float(
                new_line
            )

        # Removed item:
        # map OLD location into NEW coordinate system.
        estimated_new = (
            estimated_new_line_from_old(
                old_line
            )
        )

        if estimated_new is not None:
            return estimated_new

        # Entire NEW file absent:
        # fall back to OLD order.
        if old_line is not None:
            return float(
                old_line
            )

        return float(
            "inf"
        )


    changes = []

    for row in differences:
        old_line = row.get(
            "Old Line"
        )

        new_line = row.get(
            "New Line"
        )

        nav_new = None

        if new_line is not None:
            nav_new = new_line

        elif old_line is not None:
            estimated = (
                estimated_new_line_from_old(
                    old_line
                )
            )

            if estimated is not None:
                nav_new = max(
                    1,
                    min(
                        max(
                            1,
                            new_line_count
                        ),
                        int(
                            round(
                                estimated
                            )
                        ),
                    )
                )

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

            "navNew":
                nav_new,

            "_position":
                navigation_position(
                    row
                ),
        })


    # ========================================================
    # STRICT TOP -> BOTTOM ORDER IN NEW
    # ========================================================

    changes.sort(
        key=lambda change: (
            change["_position"],

            change.get("new")
            if change.get("new") is not None
            else float("inf"),

            change.get("old")
            if change.get("old") is not None
            else float("inf"),

            change["key"],
        )
    )


    for change in changes:
        change.pop(
            "_position",
            None,
        )


    old_status_json = {
        str(key): value
        for key, value
        in old_status.items()
    }


    new_status_json = {
        str(key): value
        for key, value
        in new_status.items()
    }


    # ========================================================
    # SAFE JSON FOR JAVASCRIPT
    # ========================================================

    def js_json(value):
        return (
            json.dumps(
                value,
                ensure_ascii=True,
            )
            .replace(
                "</",
                "<\\/"
            )
        )


    return FILE_TEMPLATE.safe_substitute(

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

        changes_json=js_json(
            changes
        ),

        old_lines_json=js_json(
            old_lines
        ),

        new_lines_json=js_json(
            new_lines
        ),

        old_status_json=js_json(
            old_status_json
        ),

        new_status_json=js_json(
            new_status_json
        ),

        old_width=estimate_width_px(
            old_lines
        ),

        new_width=estimate_width_px(
            new_lines
        ),
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

    return MASTER_TEMPLATE.safe_substitute(
        rows="\n".join(
            rows
        )
    )


# ============================================================
# GENERATE REPORTS
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