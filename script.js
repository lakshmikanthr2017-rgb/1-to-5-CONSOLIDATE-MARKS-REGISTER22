/**
 * 1 to 5th Consolidated Marks Register Portal Script
 * Excel Upload & Auto-Calculation Engine
 * Designed by : LAKSHMIKANTH R (Mob:9844444871)
 */

let studentCount = 0;

document.addEventListener("DOMContentLoaded", () => {
    addNewStudent(); // Load default student block
});

// Excel File Upload Handler
function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        
        // Read first sheet
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        parseExcelData(jsonData);
    };
    reader.readAsArrayBuffer(file);
}

// Parse Excel Data Rows
function parseExcelData(data) {
    // Clear existing table content
    document.getElementById("studentTableBody").innerHTML = "";
    studentCount = 0;

    // Search for data rows (skipping headers)
    for (let r = 11; r < data.length; r += 9) {
        const row = data[r];
        if (!row || row.length === 0) continue;

        const sNo = row[0] || (studentCount + 1);
        const name = row[1] || `STUDENT ${sNo}`;
        const presentDays = row[2] || 0;
        const totalDays = row[3] || 0;

        // Extract marks for 6 sub-exams (FA1, FA2, SA1, FA3, FA4, SA2)
        let examMarks = [];
        for (let i = 0; i < 9; i++) {
            const subRow = data[r + i] || [];
            if (i !== 3 && i !== 7 && i !== 8) { // Skip total rows
                examMarks.push({
                    kan: parseFloat(subRow[7]) || 0,
                    eng: parseFloat(subRow[9]) || 0,
                    mat: parseFloat(subRow[11]) || 0,
                    evs: parseFloat(subRow[13]) || 0
                });
            }
        }

        addNewStudentWithData(name, presentDays, totalDays, examMarks);
    }

    alert("Excel register uploaded and parsed successfully!");
}

function addNewStudent() {
    addNewStudentWithData(`STUDENT ${studentCount + 1}`, 210, 220, null);
}

function addNewStudentWithData(name, presentDays, totalDays, prefilledMarks) {
    studentCount++;
    const tbody = document.getElementById("studentTableBody");
    const blockId = `student_${studentCount}`;

    const examList = [
        { name: "FA-1", sem: "SEM-1", weight: 0.15, max: 60 },
        { name: "FA-2", sem: "SEM-1", weight: 0.15, max: 60 },
        { name: "SA-1", sem: "SEM-1", weight: 0.20, max: 80 },
        { name: "SEM-1 TOTAL", sem: "SEM-1", weight: 0.50, max: 200, isTotal: true },
        { name: "FA-3", sem: "SEM-2", weight: 0.15, max: 60 },
        { name: "FA-4", sem: "SEM-2", weight: 0.15, max: 60 },
        { name: "SA-2", sem: "SEM-2", weight: 0.20, max: 80 },
        { name: "SEM-2 TOTAL", sem: "SEM-2", weight: 0.50, max: 200, isTotal: true },
        { name: "FINAL CONSOLIDATED", sem: "SEM 1+2", weight: 1.00, max: 400, isFinal: true }
    ];

    let inputIdx = 0;

    examList.forEach((exam, idx) => {
        const tr = document.createElement("tr");
        tr.classList.add(blockId);
        if (exam.isTotal) tr.classList.add("sem-total");
        if (exam.isFinal) tr.classList.add("final-total");

        let rowHtml = "";

        if (idx === 0) {
            rowHtml += `<td rowspan="9">${studentCount}</td>`;
            rowHtml += `<td rowspan="9"><input type="text" id="${blockId}_name" value="${name}"></td>`;
            rowHtml += `<td rowspan="9"><input type="number" value="${presentDays}"></td>`;
            rowHtml += `<td rowspan="9"><input type="number" value="${totalDays}"></td>`;
        }

        rowHtml += `<td>${exam.sem}</td>`;
        rowHtml += `<td><strong>${exam.name}</strong></td>`;
        rowHtml += `<td>${exam.weight}</td>`;

        const subjects = ["kan", "eng", "mat", "evs"];
        subjects.forEach(sub => {
            if (exam.isTotal || exam.isFinal) {
                rowHtml += `<td id="${blockId}_${sub}_m_${idx}">0</td>`;
            } else {
                const markVal = prefilledMarks ? prefilledMarks[inputIdx][sub] : 0;
                rowHtml += `<td><input type="number" class="mark-input ${blockId}_input ${sub}" data-exam="${idx}" value="${markVal}" min="0" max="${exam.max / 4}"></td>`;
            }
            rowHtml += `<td id="${blockId}_${sub}_g_${idx}">-</td>`;
        });

        if (!exam.isTotal && !exam.isFinal) inputIdx++;

        rowHtml += `<td id="${blockId}_tot_m_${idx}">0</td>`;
        rowHtml += `<td id="${blockId}_tot_p_${idx}">0%</td>`;
        rowHtml += `<td id="${blockId}_tot_g_${idx}">-</td>`;

        if (idx === 0) {
            rowHtml += `<td rowspan="9" id="${blockId}_result" style="font-weight:bold;">PENDING</td>`;
        }

        tr.innerHTML = rowHtml;
        tbody.appendChild(tr);
    });

    document.querySelectorAll(`.${blockId}_input`).forEach(input => {
        input.addEventListener("input", () => calculateStudentData(blockId));
    });

    calculateStudentData(blockId);
}

function getGrade(percentage) {
    if (percentage >= 85) return "A+";
    if (percentage >= 70) return "A";
    if (percentage >= 50) return "B+";
    if (percentage >= 35) return "B";
    return "C";
}

function calculateStudentData(blockId) {
    const subjects = ["kan", "eng", "mat", "evs"];
    const maxMarks = [60, 60, 80, 200, 60, 60, 80, 200, 400];
    let marks = { kan: Array(9).fill(0), eng: Array(9).fill(0), mat: Array(9).fill(0), evs: Array(9).fill(0) };

    document.querySelectorAll(`.${blockId}_input`).forEach(input => {
        const sub = input.classList[2];
        const examIdx = parseInt(input.dataset.exam);
        const val = parseFloat(input.value) || 0;
        marks[sub][examIdx] = val;

        const grade = getGrade((val / (maxMarks[examIdx] / 4)) * 100);
        document.getElementById(`${blockId}_${sub}_g_${examIdx}`).innerText = grade;
    });

    subjects.forEach(sub => {
        marks[sub][3] = marks[sub][0] + marks[sub][1] + marks[sub][2];
        marks[sub][7] = marks[sub][4] + marks[sub][5] + marks[sub][6];
        marks[sub][8] = marks[sub][3] + marks[sub][7];

        [3, 7, 8].forEach(idx => {
            document.getElementById(`${blockId}_${sub}_m_${idx}`).innerText = marks[sub][idx];
            const grade = getGrade((marks[sub][idx] / (maxMarks[idx] / 4)) * 100);
            document.getElementById(`${blockId}_${sub}_g_${idx}`).innerText = grade;
        });
    });

    for (let i = 0; i < 9; i++) {
        const tot = marks.kan[i] + marks.eng[i] + marks.mat[i] + marks.evs[i];
        const pct = (tot / maxMarks[i]) * 100;
        const grade = getGrade(pct);

        document.getElementById(`${blockId}_tot_m_${i}`).innerText = tot;
        document.getElementById(`${blockId}_tot_p_${i}`).innerText = pct.toFixed(1) + "%";
        document.getElementById(`${blockId}_tot_g_${i}`).innerText = grade;
    }

    const finalPct = (marks.kan[8] + marks.eng[8] + marks.mat[8] + marks.evs[8]) / 400 * 100;
    const resCell = document.getElementById(`${blockId}_result`);
    if (finalPct >= 35) {
        resCell.innerText = "PASS";
        resCell.style.color = "green";
    } else {
        resCell.innerText = "FAIL";
        resCell.style.color = "red";
    }
}

function filterStudents() {
    const query = document.getElementById("studentSearch").value.toLowerCase();
    for (let i = 1; i <= studentCount; i++) {
        const blockId = `student_${i}`;
        const nameInput = document.getElementById(`${blockId}_name`);
        const nameVal = nameInput ? nameInput.value.toLowerCase() : "";
        const isMatch = nameVal.includes(query) || i.toString() === query;

        document.querySelectorAll(`.${blockId}`).forEach(row => {
            row.style.display = isMatch ? "" : "none";
        });
    }
}
