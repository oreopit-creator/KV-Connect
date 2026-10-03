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
        "A": "10-a.xlsx",
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

// Inside each file: one sheet per month (Apr, May, ... Mar)
// Sheet layout (A = 0, B = 1, C = 2 ...):
//   A = Roll No | B = Name | C onwards = daily attendance
//   Row 3 onward = students
// Sundays have no column in the sheet.
const ROLL_COL = 0;
const NAME_COL = 1;
const FIRST_DATE_COL = 2;
const FIRST_STUDENT_ROW = 2;   // Row 3 in Excel

// Weeks: by default every 6 date columns = 1 week (Mon-Sat).
// If a month needs exact week labels, add it here.
// The key is the sheet name in lowercase letters/numbers only (e.g. "Sep" -> "sep").
const WORKING_DAYS_PER_WEEK = 6;

const WEEK_OVERRIDES = {
    "sep": [
        { name: "Week 1", dates: "1 Sep - 7 Sep",   start: 2,  end: 7  },
        { name: "Week 2", dates: "8 Sep - 14 Sep",  start: 8,  end: 14 },
        { name: "Week 3", dates: "15 Sep - 21 Sep", start: 15, end: 21 },
        { name: "Week 4", dates: "22 Sep - 28 Sep", start: 22, end: 28 },
        { name: "Week 5", dates: "29 Sep - 30 Sep", start: 28, end: 30 }
    ]
};


// ==========================================
// STATE
// ==========================================

let selectedClass = "";
let selectedSection = "";

// One entry per month sheet where this student was found:
// { name, student, lastCol }
let studentMonths = [];

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
    ["overall", "monthly", "weekly"].forEach(function (id) {
        $(id).style.display = "none";
    });
}

function normalizeName(name) {
    return String(name).toLowerCase().replace(/[^a-z0-9]/g, "");
}

function cell(row, col) {
    return String(row[col] ?? "").trim();
}

function countStatus(student, start, end) {
    const counts = { present: 0, absent: 0, leave: 0 };

    for (let i = start; i <= end; i++) {
        const status = cell(student, i).toLowerCase();

        if (Object.hasOwn(counts, status)) {
            counts[status]++;
        }
    }

    return counts;
}

// Leave counts as a working day AND counts towards attendance:
// percentage = (present + leave) / total working days
function calcPercentage(present, absent, leave) {
    const workingDays = present + absent + leave;
    return workingDays > 0 ? ((present + leave) / workingDays) * 100 : 0;
}

function resetSelect(select) {
    select.length = 1;   // keep the first "Select ..." option
    select.value = "";
    select.disabled = true;
}

function getSelectedMonth() {
    const name = $("monthSelect").value;

    return studentMonths.find(function (m) {
        return m.name === name;
    });
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

// Last date column = widest student row in the sheet
function getLastCol(rows) {
    let last = FIRST_DATE_COL;

    for (let i = FIRST_STUDENT_ROW; i < rows.length; i++) {
        last = Math.max(last, rows[i].length - 1);
    }

    return last;
}

function findStudent(rows, roll) {
    for (let i = FIRST_STUDENT_ROW; i < rows.length; i++) {
        if (cell(rows[i], ROLL_COL) === roll) {
            return rows[i];
        }
    }

    return null;
}

function getWeeks(sheetName, lastCol) {
    const override = WEEK_OVERRIDES[normalizeName(sheetName)];
    if (override) return override;

    const weeks = [];
    let number = 1;

    for (let s = FIRST_DATE_COL; s <= lastCol; s += WORKING_DAYS_PER_WEEK) {
        weeks.push({
            name: "Week " + number,
            dates: "",
            start: s,
            end: Math.min(s + WORKING_DAYS_PER_WEEK - 1, lastCol)
        });
        number++;
    }

    return weeks;
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

// Month dropdown lives inside Monthly History
$("monthSelect").addEventListener("change", function () {
    renderMonthly();

    // If weekly history is already open, refresh it for the new month
    if ($("weekly").style.display === "block") {
        showWeekly();
    }
});


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
    studentMonths = [];

    const file = FILES[cls][section];

    try {
        const workbook = await loadWorkbook(file);

        // Look for the student in every month sheet
        workbook.SheetNames.forEach(function (name) {
            const rows = getRows(workbook, name);
            const student = findStudent(rows, roll);

            if (student) {
                studentMonths.push({
                    name: name,
                    student: student,
                    lastCol: getLastCol(rows)
                });
            }
        });
    } catch (error) {
        console.error("Excel load error:", error);
        alert("Excel file load nahi ho rahi (" + file + ")! Local server se page open karo aur file ka naam check karo.");
        return;
    }

    if (studentMonths.length === 0) {
        alert("Student not found!");
        return;
    }

    selectedClass = cls;
    selectedSection = section;

    fillMonthSelect();
    showOverall();
}


// ==========================================
// MONTH DROPDOWN (months where the student has data)
// ==========================================

function fillMonthSelect() {
    const select = $("monthSelect");

    select.length = 0;

    studentMonths.forEach(function (m) {
        select.add(new Option(m.name, m.name));
    });

    // Default = latest month
    select.value = studentMonths[studentMonths.length - 1].name;
}


// ==========================================
// OVERALL (full year = all month sheets together)
// ==========================================

function showOverall() {
    $("overall").style.display = "block";

    const first = studentMonths[0].student;

    setText("studentRoll", cell(first, ROLL_COL));
    setText("studentName", cell(first, NAME_COL));
    setText("studentClass", selectedClass + "-" + selectedSection);
    setText("studentPeriod", "Full Year (" + studentMonths.length + " months)");

    let present = 0;
    let absent = 0;
    let leave = 0;

    studentMonths.forEach(function (m) {
        const c = countStatus(m.student, FIRST_DATE_COL, m.lastCol);
        present += c.present;
        absent += c.absent;
        leave += c.leave;
    });

    // Working days = Present + Absent + Leave (Sundays are not in the sheet)
    setText("totalDays", present + absent + leave);
    setText("presentDays", present);
    setText("absentDays", absent);
    setText("leaveDays", leave);
    setText("attendancePercent", calcPercentage(present, absent, leave).toFixed(2) + "%");
}


// ==========================================
// MONTHLY HISTORY (one month at a time)
// ==========================================

function showMonthly() {
    if (studentMonths.length === 0) return;

    $("monthly").style.display = "block";
    renderMonthly();
}

function renderMonthly() {
    const month = getSelectedMonth();

    if (!month) {
        $("monthlyData").innerHTML = "";
        return;
    }

    const { present, absent, leave } =
        countStatus(month.student, FIRST_DATE_COL, month.lastCol);

    $("monthlyData").innerHTML = `
        <div class="month">
            <h3>${month.name}</h3>
            <p>Total Working Days: <strong>${present + absent + leave}</strong></p>
            <p>Present: <strong>${present}</strong></p>
            <p>Absent: <strong>${absent}</strong></p>
            <p>Leave: <strong>${leave}</strong></p>
            <p>Attendance: <strong>${calcPercentage(present, absent, leave).toFixed(2)}%</strong></p>
        </div>
    `;
}


// ==========================================
// WEEKLY HISTORY (for the month selected in Monthly History)
// ==========================================

function showWeekly() {
    const month = getSelectedMonth();

    if (!month) return;

    $("weekly").style.display = "block";

    let html = `<h3>${month.name}</h3>`;

    getWeeks(month.name, month.lastCol).forEach(function (week) {
        const { present, absent, leave } =
            countStatus(month.student, week.start, week.end);

        html += `
            <div class="week">
                <h3>${week.name}</h3>
                ${week.dates ? `<p>${week.dates}</p>` : ""}
                <p>Present: <strong>${present}</strong></p>
                <p>Absent: <strong>${absent}</strong></p>
                <p>Leave: <strong>${leave}</strong></p>
                <p>Attendance: <strong>${calcPercentage(present, absent, leave).toFixed(2)}%</strong></p>
            </div>
        `;
    });

    $("weeklyData").innerHTML = html;
}


// ==========================================
// PRESS ENTER TO SEARCH
// ==========================================

$("rollInput").addEventListener("keydown", function (e) {
    if (e.key === "Enter") {
        searchStudent();
    }
});