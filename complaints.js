// ==========================================
// SETTINGS - edit these
// ==========================================

// Paste the Web app URL you get after deploying Code.gs (Google Apps Script)
const API_URL = "https://script.google.com/macros/s/AKfycbwjz0IlvFtOv49lMIHhq3Ig3FemSw3qGJbCaNscKB4ZyguOrJv6jaJ7nmrsozjOVwg/exec";

// Classes and their sections
const CLASSES = {
    "6": ["A", "B", "C","D"],
    "7": ["A", "B", "C","D"],
    "8": ["A", "B", "C","D"],
    "9":  ["A", "B","C","D"],
    "10": ["A", "B", "C","D"],
    "11": ["A", "B", "C","D"],
    "12": ["A", "B", "C","D"],
};


// ==========================================
// STATE
// ==========================================

let currentRemarks = [];


// ==========================================
// HELPERS
// ==========================================

function $(id) {
    return document.getElementById(id);
}

// Remarks come from a sheet, so make them safe before putting them in the page
function escapeHtml(text) {
    return String(text ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
}

function apiReady() {
    if (!API_URL || API_URL.startsWith("PASTE_")) {
        alert("API_URL set nahi hai! complaints.js ke upar apna Web app URL paste karo.");
        return false;
    }
    return true;
}

function resetSelect(select) {
    select.length = 1;   // keep the first "Select ..." option
    select.value = "";
    select.disabled = true;
}

// Class dropdown -> fills its section dropdown
function setupClassSection(classId, sectionId, onChange) {
    const classSelect = $(classId);
    const sectionSelect = $(sectionId);

    Object.keys(CLASSES).forEach(function (cls) {
        classSelect.add(new Option("Class " + cls, cls));
    });

    classSelect.addEventListener("change", function () {
        resetSelect(sectionSelect);

        const sections = CLASSES[this.value] || [];

        sections.forEach(function (sec) {
            sectionSelect.add(new Option("Section " + sec, sec));
        });

        sectionSelect.disabled = sections.length === 0;

        if (onChange) onChange();
    });

    if (onChange) {
        sectionSelect.addEventListener("change", onChange);
    }
}

function hideRemarks() {
    $("remarksResult").style.display = "none";
}

setupClassSection("classSelect", "sectionSelect", hideRemarks);
setupClassSection("tClass", "tSection");
setupClassSection("dClass", "dSection", clearDeleteList);
$("rollInput").addEventListener("change", hideRemarks);
$("typeFilter").addEventListener("change", renderRemarks);

hideRemarks();


// ==========================================
// VIEW REMARKS
// ==========================================

async function viewRemarks() {
    const cls = $("classSelect").value;
    const sec = $("sectionSelect").value;
    const roll = $("rollInput").value.trim();

    if (!cls) return alert("Class select karo!");
    if (!sec) return alert("Section select karo!");
    if (roll === "") return alert("Roll Number enter karo!");
    if (!apiReady()) return;

    hideRemarks();

    const button = $("viewBtn");
    button.disabled = true;
    button.textContent = "Loading...";

    try {
        const url = API_URL +
            "?cls=" + encodeURIComponent(cls) +
            "&sec=" + encodeURIComponent(sec) +
            "&roll=" + encodeURIComponent(roll);

        const response = await fetch(url);
        const data = await response.json();

        if (!data.ok) {
            throw new Error(data.error || "Unknown error");
        }

        currentRemarks = data.remarks;
    } catch (error) {
        console.error("Load error:", error);
        alert("Remarks load nahi hue! Internet aur API_URL check karo.");
        return;
    } finally {
        button.disabled = false;
        button.textContent = "View Remarks";
    }

    // Name: use the one teachers typed, if any
    const withName = currentRemarks.find(function (r) {
        return r.name;
    });

    $("studentLabel").textContent = withName ? withName.name + " (Roll " + roll + ")" : "Roll " + roll;
    $("studentClass").textContent = cls + "-" + sec;

    $("typeFilter").value = "";
    renderSummary();
    renderRemarks();

    $("remarksResult").style.display = "block";
}


// ==========================================
// SHOW REMARKS
// ==========================================

function countType(type) {
    return currentRemarks.filter(function (r) {
        return r.type === type;
    }).length;
}

function renderSummary() {
    $("summaryData").innerHTML = `
        <span class="pill pill-complaint">Complaints: ${countType("Complaint")}</span>
        <span class="pill pill-warning">Warnings: ${countType("Warning")}</span>
        <span class="pill pill-appreciation">Appreciation: ${countType("Appreciation")}</span>
    `;
}

function renderRemarks() {
    const filter = $("typeFilter").value;

    // Newest first
    const list = currentRemarks
        .filter(function (r) {
            return filter === "" || r.type === filter;
        })
        .reverse();

    if (list.length === 0) {
        $("remarksList").innerHTML = `<p class="muted empty">Koi remark nahi mila.</p>`;
        return;
    }

    let html = "";

    list.forEach(function (r) {
        const typeClass = String(r.type).toLowerCase().replace(/[^a-z]/g, "");
        const meta = [r.teacher, r.subject].filter(Boolean).map(escapeHtml).join(" &middot; ");

        html += `
            <div class="remark remark-${typeClass}">
                <div class="remark-top">
                    <span class="badge badge-${typeClass}">${escapeHtml(r.type)}</span>
                    <span class="remark-date">${escapeHtml(r.date)}</span>
                </div>
                <div class="remark-text">${escapeHtml(r.remark)}</div>
                <div class="remark-meta">${meta}</div>
            </div>
        `;
    });

    $("remarksList").innerHTML = html;
}


// ==========================================
// TEACHER FORM
// ==========================================

function togglePanel(id, otherId) {
    const panel = $(id);
    const opening = panel.style.display !== "block";

    $(otherId).style.display = "none";
    panel.style.display = opening ? "block" : "none";

    if (opening) {
        panel.scrollIntoView({ behavior: "smooth", block: "start" });
    }
}

function toggleTeacherForm() {
    togglePanel("teacherForm", "deleteForm");
}

function toggleDeleteForm() {
    togglePanel("deleteForm", "teacherForm");
}

function showMessage(text, ok) {
    const box = $("formMessage");
    box.textContent = text;
    box.className = "form-message " + (ok ? "ok" : "error");
}

async function submitRemark() {
    const payload = {
        teacher: $("tTeacher").value.trim(),
        action:  "add",
        pin:     $("tPin").value,
        cls:     $("tClass").value,
        sec:     $("tSection").value,
        roll:    $("tRoll").value.trim(),
        name:    $("tName").value.trim(),
        subject: $("tSubject").value.trim(),
        type:    $("tType").value,
        remark:  $("tRemark").value.trim()
    };

    if (!payload.teacher) return showMessage("Teacher name Required.", false);
    if (!payload.pin)     return showMessage("Teacher PIN Required.", false);
    if (!payload.cls)     return showMessage("Class select Required.", false);
    if (!payload.sec)     return showMessage("Section select Required.", false);
    if (!payload.roll)    return showMessage("Roll number Required.", false);
    if (!payload.remark)  return showMessage("Remark Required.", false);
    if (!apiReady()) return;

    const button = $("saveBtn");
    button.disabled = true;
    button.textContent = "Saving...";
    showMessage("", true);

    try {
        // No custom headers, so the browser sends a simple request (no CORS preflight)
        const response = await fetch(API_URL, {
            method: "POST",
            body: JSON.stringify(payload)
        });

        const data = await response.json();

        if (!data.ok) {
            showMessage(data.error || "Save nahi hua.", false);
            return;
        }

        showMessage("Remark save ho gaya!", true);
        $("tRemark").value = "";
        $("tSubject").value = "";
    } catch (error) {
        console.error("Save error:", error);
        showMessage("Save nahi hua. Internet aur API_URL check karo.", false);
    } finally {
        button.disabled = false;
        button.textContent = "Save Remark";
    }
}


// ==========================================
// DELETE A REMARK (teachers)
// ==========================================

let deleteRemarks = [];
let deleteContext = null;   // class / section / roll of the loaded list

function showDeleteMessage(text, ok) {
    const box = $("deleteMessage");
    box.textContent = text;
    box.className = "form-message " + (ok ? "ok" : "error");
}

function clearDeleteList() {
    deleteRemarks = [];
    deleteContext = null;
    $("deleteList").innerHTML = "";
    showDeleteMessage("", true);
}

async function loadForDelete() {
    const cls = $("dClass").value;
    const sec = $("dSection").value;
    const roll = $("dRoll").value.trim();

    if (!cls) return showDeleteMessage("Class select Required.", false);
    if (!sec) return showDeleteMessage("Section select Required.", false);
    if (!roll) return showDeleteMessage("Roll number Required.", false);
    if (!apiReady()) return;

    const button = $("loadDeleteBtn");
    button.disabled = true;
    button.textContent = "Loading...";

    try {
        const url = API_URL +
            "?cls=" + encodeURIComponent(cls) +
            "&sec=" + encodeURIComponent(sec) +
            "&roll=" + encodeURIComponent(roll);

        const response = await fetch(url);
        const data = await response.json();

        if (!data.ok) {
            throw new Error(data.error || "Unknown error");
        }

        deleteRemarks = data.remarks;
        deleteContext = { cls: cls, sec: sec, roll: roll };
    } catch (error) {
        console.error("Load error:", error);
        showDeleteMessage("Remarks load nahi hue. Internet aur API_URL check karo.", false);
        return;
    } finally {
        button.disabled = false;
        button.textContent = "Load Remarks";
    }

    renderDeleteList();
}

function renderDeleteList() {
    if (deleteRemarks.length === 0) {
        $("deleteList").innerHTML = `<p class="muted empty">NO Remark Found.</p>`;
        return;
    }

    let html = "";

    // Newest first (keep the real index so the right remark is deleted)
    for (let i = deleteRemarks.length - 1; i >= 0; i--) {
        const r = deleteRemarks[i];
        const typeClass = String(r.type).toLowerCase().replace(/[^a-z]/g, "");
        const meta = [r.teacher, r.subject].filter(Boolean).map(escapeHtml).join(" &middot; ");

        html += `
            <div class="remark remark-${typeClass}">
                <div class="remark-top">
                    <span class="badge badge-${typeClass}">${escapeHtml(r.type)}</span>
                    <span class="remark-date">${escapeHtml(r.date)}</span>
                </div>
                <div class="remark-text">${escapeHtml(r.remark)}</div>
                <div class="remark-meta">${meta}</div>
                <button type="button" class="delete-btn" data-index="${i}">Delete</button>
            </div>
        `;
    }

    $("deleteList").innerHTML = html;
}

$("deleteList").addEventListener("click", function (e) {
    const button = e.target.closest(".delete-btn");

    if (button) {
        deleteRemark(Number(button.dataset.index), button);
    }
});

async function deleteRemark(index, button) {
    const remark = deleteRemarks[index];
    const pin = $("dPin").value;

    if (!remark || !deleteContext) return;
    if (!pin) return showDeleteMessage("Teacher PIN Required.", false);
    if (!confirm("Are you Sure you want to delete Remark?")) return;

    button.disabled = true;
    button.textContent = "Deleting...";
    showDeleteMessage("", true);

    try {
        const response = await fetch(API_URL, {
            method: "POST",
            body: JSON.stringify({
                action: "delete",
                pin:    pin,
                cls:    deleteContext.cls,
                sec:    deleteContext.sec,
                roll:   deleteContext.roll,
                row:    remark.id,
                remark: remark.remark
            })
        });

        const data = await response.json();

        if (!data.ok) {
            showDeleteMessage(data.error || "Can't Delete.", false);
            button.disabled = false;
            button.textContent = "Delete";
            return;
        }

        showDeleteMessage("Remark deleted.", true);
        await loadForDelete();
        showDeleteMessage("Remark deleted.", true);
    } catch (error) {
        console.error("Delete error:", error);
        showDeleteMessage("can't Delete.", false);
        button.disabled = false;
        button.textContent = "Delete";
    }
}


// ==========================================
// PRESS ENTER TO SEARCH
// ==========================================

$("rollInput").addEventListener("keydown", function (e) {
    if (e.key === "Enter") {
        viewRemarks();
    }
});