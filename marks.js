// ==========================================
// SETTINGS - edit these to match your data
// ==========================================

// One Excel file per class + section.
// Add your other classes / sections here the same way.
const FILES = {
    "6": {
        "A": "6-a.xlsx",
        "B": "6-b.xlsx",
        "C": "6-c.xlsx",
        "D": "6-d.xlsx"
    },
    "7": {
        "A": "7-a.xlsx",
        "B": "7-b.xlsx",
        "C": "7-c.xlsx",
        "D": "7-d.xlsx"
    },
    "8": {
        "A": "8-a.xlsx",
        "B": "8-b.xlsx",
        "C": "8-c.xlsx",
        "D": "8-d.xlsx"
    },
    "9": {
        "A": "9-a.xlsx",
        "B": "9-b.xlsx",
        "C": "9-c.xlsx",
        "D": "9-d.xlsx"
    },
    "10": {
        "A": "10-A-mark.xlsx",
        "B": "10-b.xlsx",
        "C": "10-c.xlsx",
        "D": "10-d.xlsx"
    },
    "11": {
        "A": "11-a.xlsx",
        "B": "11-b.xlsx",
        "C": "11-c.xlsx",
        "D": "11-d.xlsx"
    },
    "12":{
        "A": "12-a.xlsx",
        "B": "12-b.xlsx",
        "C": "12-c.xlsx",
        "D": "12-d.xlsx"
    }
};
const EXAMS = [
    { name: "PT1",         sheet: "P-T-1",       max: 40, lowBelow: 13 },
    { name: "PT2",         sheet: "P-T-2",       max: 40, lowBelow: 13 },
    { name: "Half-Yearly", sheet: "Half yearly", max: 80, lowBelow: 27 }
    // Final exam: add a line like this (and a comma after the line above)
    // { name: "Final", sheet: "Final", max: 80 }
];

const ROLL_COL = 0;             // A
const NAME_COL = 1;             // B
const FIRST_SUBJECT_COL = 2;    // C onwards = subjects
const SUBJECT_HEADER_ROW = 0;   // Row 1
const FIRST_STUDENT_ROW = 1;    // Row 2


// ==========================================
// STATE
// ==========================================

let currentResults = [];   // one entry per exam: { exam, data (or null) }
let currentClass = "";
let currentSection = "";

const workbookCache = {};


// ==========================================
// HELPERS
// ==========================================

function $(id) {
    return document.getElementById(id);
}

function setText(id, value) {
    $(id).textContent = value;
}

function hideAllResults() {
    ["overall", "details"].forEach(function (id) {
        $(id).style.display = "none";
    });
}

function normalizeName(name) {
    return String(name).toLowerCase().replace(/[^a-z0-9]/g, "");
}

function cell(row, col) {
    return String(row[col] ?? "").trim();
}

// Names come from Excel, so make them safe before putting them in the page
function escapeHtml(text) {
    return String(text)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
}

// Returns a number, or null for blank / "AB" / any text
function toNumber(value) {
    const text = String(value ?? "").trim();
    if (text === "") return null;

    const number = Number(text);
    return Number.isFinite(number) ? number : null;
}

// true if this exam has a "lowBelow" limit and the marks are under it
function isLow(exam, marks) {
    return exam.lowBelow !== undefined && marks !== null && marks < exam.lowBelow;
}

function formatNumber(value) {
    return String(Number(value.toFixed(2)));
}

function resetSelect(select) {
    select.length = 1;   // keep the first "Select ..." option
    select.value = "";
    select.disabled = true;
}


// ==========================================
// EXCEL LOADING (each file is loaded once)
// ==========================================

function loadWorkbook(file) {
    if (!workbookCache[file]) {
        const promise = fetch(file)
            .then(function (response) {
                if (!response.ok) {
                    throw new Error("HTTP " + response.status + " for " + file);
                }
                return response.arrayBuffer();
            })
            .then(function (buffer) {
                return XLSX.read(buffer, { type: "array" });
            });

        // If loading fails, allow a retry next time
        promise.catch(function () {
            delete workbookCache[file];
        });

        workbookCache[file] = promise;
    }

    return workbookCache[file];
}

function getRows(workbook, sheetName) {
    return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
        header: 1,
        defval: "",
        raw: false
    });
}

function findStudent(rows, roll) {
    for (let i = FIRST_STUDENT_ROW; i < rows.length; i++) {
        if (cell(rows[i], ROLL_COL) === roll) {
            return rows[i];
        }
    }

    return null;
}


// ==========================================
// MARKS CALCULATION
// ==========================================

// Marks of one student in one exam sheet (null if sheet / student is missing)
function getExamData(workbook, exam, roll) {
    const wanted = normalizeName(exam.sheet);

    const sheetName = workbook.SheetNames.find(function (name) {
        return normalizeName(name) === wanted;
    });

    if (!sheetName) return null;

    const rows = getRows(workbook, sheetName);
    const header = rows[SUBJECT_HEADER_ROW] || [];
    const student = findStudent(rows, roll);

    if (!student) return null;

    const subjects = [];
    let obtained = 0;
    let counted = 0;
    let lowCount = 0;

    for (let c = FIRST_SUBJECT_COL; c < header.length; c++) {
        const subjectName = cell(header, c);

        if (subjectName === "") continue;

        const raw = cell(student, c);
        const marks = toNumber(raw);

        subjects.push({ name: subjectName, raw: raw, marks: marks });

        if (marks !== null) {
            obtained += marks;
            counted++;
        }

        if (isLow(exam, marks)) lowCount++;
    }

    const total = subjects.length * exam.max;

    return {
        student: student,
        subjects: subjects,
        obtained: obtained,
        total: total,
        hasMarks: counted > 0,
        lowCount: lowCount,
        percentage: total > 0 ? (obtained / total) * 100 : 0
    };
}


// ==========================================
// DROPDOWNS
// ==========================================

Object.keys(FILES).forEach(function (cls) {
    $("classSelect").add(new Option("Class " + cls, cls));
});

hideAllResults();

// Class -> sections
$("classSelect").addEventListener("change", function () {
    resetSelect($("sectionSelect"));
    hideAllResults();

    const sections = Object.keys(FILES[this.value] || {});

    sections.forEach(function (sec) {
        $("sectionSelect").add(new Option("Section " + sec, sec));
    });

    $("sectionSelect").disabled = sections.length === 0;
});

["sectionSelect", "rollInput"].forEach(function (id) {
    $(id).addEventListener("change", hideAllResults);
});

// Exam dropdown (inside Subject-wise Marks)
$("examSelect").addEventListener("change", renderDetails);


// ==========================================
// SEARCH STUDENT
// ==========================================

async function searchStudent() {
    const cls = $("classSelect").value;
    const section = $("sectionSelect").value;
    const roll = $("rollInput").value.trim();

    if (!cls) return alert("Class select karo!");
    if (!section) return alert("Section select karo!");
    if (roll === "") return alert("Roll Number enter karo!");

    // Clear old results from any previous search
    hideAllResults();
    currentResults = [];

    const file = FILES[cls][section];

    try {
        const workbook = await loadWorkbook(file);

        currentResults = EXAMS.map(function (exam) {
            return { exam: exam, data: getExamData(workbook, exam, roll) };
        });
    } catch (error) {
        console.error("Excel load error:", error);
        alert("Excel file load nahi ho rahi (" + file + ")! Local server se page open karo aur file ka naam check karo.");
        return;
    }

    const found = currentResults.some(function (r) {
        return r.data !== null;
    });

    if (!found) {
        alert("Student not found!");
        return;
    }

    currentClass = cls;
    currentSection = section;

    fillExamSelect();
    showOverall();
}


// ==========================================
// OVERALL (summary of every exam)
// ==========================================

function showOverall() {
    $("overall").style.display = "block";

    const firstFound = currentResults.find(function (r) {
        return r.data !== null;
    });

    setText("studentRoll", cell(firstFound.data.student, ROLL_COL));
    setText("studentName", cell(firstFound.data.student, NAME_COL));
    setText("studentClass", currentClass + "-" + currentSection);

    let html = `
        <table class="marks-table">
            <thead>
                <tr>
                    <th>Exam</th>
                    <th>Marks</th>
                    <th>Percentage</th>
                </tr>
            </thead>
            <tbody>
    `;

    currentResults.forEach(function (r) {
        const data = r.data;

        if (data && data.hasMarks) {
            html += `
                <tr>
                    <td>${escapeHtml(r.exam.name)}</td>
                    <td>
                        ${formatNumber(data.obtained)} / ${data.total}
                        ${data.lowCount > 0
                            ? `<span class="low-note">${data.lowCount} subject(s) below ${r.exam.lowBelow}</span>`
                            : ""}
                    </td>
                    <td><span class="percent-pill">${data.percentage.toFixed(2)}%</span></td>
                </tr>
            `;
        } else {
            html += `
                <tr>
                    <td>${escapeHtml(r.exam.name)}</td>
                    <td colspan="2" class="muted">Not available</td>
                </tr>
            `;
        }
    });

    html += "</tbody></table>";

    $("summaryData").innerHTML = html;
}


// ==========================================
// SUBJECT-WISE MARKS (one exam at a time)
// ==========================================

function fillExamSelect() {
    const select = $("examSelect");

    select.length = 0;

    let defaultIndex = 0;

    currentResults.forEach(function (r, index) {
        select.add(new Option(r.exam.name, String(index)));

        // Default = latest exam that has marks
        if (r.data && r.data.hasMarks) defaultIndex = index;
    });

    select.value = String(defaultIndex);
}

function showDetails() {
    if (currentResults.length === 0) return;

    $("details").style.display = "block";
    renderDetails();
}

function renderDetails() {
    const result = currentResults[Number($("examSelect").value)];

    if (!result) return;

    const exam = result.exam;
    const data = result.data;

    if (!data || !data.hasMarks) {
        $("detailsData").innerHTML = `<p class="muted">${escapeHtml(exam.name)} marks are not available yet.</p>`;
        return;
    }

    let html = `
        <table class="marks-table">
            <thead>
                <tr>
                    <th>Subject</th>
                    <th>Marks</th>
                    <th>Out of</th>
                </tr>
            </thead>
            <tbody>
    `;

    data.subjects.forEach(function (subject) {
        const shown = subject.raw === "" ? "-" : escapeHtml(subject.raw);

        html += `
            <tr>
                <td>${escapeHtml(subject.name)}</td>
                <td class="${isLow(exam, subject.marks) ? "low-marks" : ""}">${shown}</td>
                <td>${exam.max}</td>
            </tr>
        `;
    });

    html += `
            <tr class="total-row">
                <td>Total</td>
                <td>${formatNumber(data.obtained)}</td>
                <td>${data.total}</td>
            </tr>
            <tr class="total-row">
                <td>Percentage</td>
                <td colspan="2">${data.percentage.toFixed(2)}%</td>
            </tr>
        </tbody>
    </table>
    `;

    $("detailsData").innerHTML = html;
}


// ==========================================
// PRESS ENTER TO SEARCH
// ==========================================

$("rollInput").addEventListener("keydown", function (e) {
    if (e.key === "Enter") {
        searchStudent();
    }
});