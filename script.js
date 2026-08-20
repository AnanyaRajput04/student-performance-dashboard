// ==========================================
// STUDENT PERFORMANCE ANALYTICS DASHBOARD
// Excel Data + Charts
// ==========================================

let students = [];
let performanceChart;
let resultChart;
let attendanceChart;

// ==========================================
// LOAD EXCEL FILE
// ==========================================

async function loadExcelData() {

    try {

        const response = await fetch("Student_Data.xlsx");
        const arrayBuffer = await response.arrayBuffer();

        const workbook = XLSX.read(arrayBuffer, {
            type: "array"
        });

        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        students = XLSX.utils.sheet_to_json(worksheet);

        console.log("Excel loaded successfully");
        console.log("Students loaded:", students.length);
        console.log("First student:", students[0]);

        updateDashboard();

    } catch (error) {

        console.error("Excel loading error:", error);

    }
}


// ==========================================
// UPDATE DASHBOARD
// ==========================================

function updateDashboard() {

    if (!students.length) return;

    // --------------------------------------
    // TOTAL STUDENTS
    // --------------------------------------

    const totalStudents = students.length;

    // --------------------------------------
    // AVERAGE MARKS
    // Use Excel Average_Marks column
    // --------------------------------------

    const totalAverageMarks = students.reduce(
        (sum, student) =>
            sum + Number(student.Average_Marks || 0),
        0
    );

    const averageMarks =
        totalAverageMarks / totalStudents;


    // --------------------------------------
    // AVERAGE ATTENDANCE
    // --------------------------------------

    const totalAttendance = students.reduce(
        (sum, student) =>
            sum + Number(student.Attendance || 0),
        0
    );

    const averageAttendance =
        totalAttendance / totalStudents;


    // --------------------------------------
    // HIGHEST & LOWEST MARKS
    // --------------------------------------

    const highestMarks = Math.max(
        ...students.map(student =>
            Number(student.Average_Marks || 0)
        )
    );

    const lowestMarks = Math.min(
        ...students.map(student =>
            Number(student.Average_Marks || 0)
        )
    );


    // --------------------------------------
    // PASS / FAIL
    // --------------------------------------

    const passStudents = students.filter(
        student =>
            String(student.Result).toLowerCase() === "pass"
    ).length;

    const failStudents =
        totalStudents - passStudents;

    const passPercentage =
        (passStudents / totalStudents) * 100;


    // ======================================
    // UPDATE KPI CARDS
    // ======================================

    const totalElement =
        document.getElementById("totalStudents");

    const averageElement =
        document.getElementById("averageMarks");

    const attendanceElement =
        document.getElementById("averageAttendance");

    const passElement =
        document.getElementById("passPercentage");


    if (totalElement) {
        totalElement.textContent = totalStudents;
    }

    if (averageElement) {
        averageElement.textContent =
            averageMarks.toFixed(3) + "%";
    }

    if (attendanceElement) {
        attendanceElement.textContent =
            averageAttendance.toFixed(2) + "%";
    }

    if (passElement) {
        passElement.textContent =
            passPercentage.toFixed(0) + "%";
    }


    // ======================================
    // CREATE CHARTS
    // ======================================

    createPerformanceChart();
    createResultChart();
    createAttendanceChart();
    createTopPerformers();
    displayStudentsTable();

}


// ==========================================
// PERFORMANCE CHART
// ==========================================

function createPerformanceChart() {

    const canvas =
        document.getElementById("performanceChart");

    if (!canvas) return;

    const ctx = canvas.getContext("2d");

    if (performanceChart) {
        performanceChart.destroy();
    }

    const subjects = [
        "DBMS",
        "Python",
        "DSA",
        "Web_Development",
        "OS"
    ];

    const labels = [
        "DBMS",
        "Python",
        "DSA",
        "Web Development",
        "OS"
    ];

    const subjectAverages = subjects.map(subject => {

        const total = students.reduce(
            (sum, student) =>
                sum + Number(student[subject] || 0),
            0
        );

        return total / students.length;

    });


    performanceChart = new Chart(ctx, {

        type: "bar",

        data: {

            labels: labels,

            datasets: [{

                label: "Average Marks",

                data: subjectAverages,

                borderWidth: 1

            }]

        },

        options: {

            responsive: true,

            maintainAspectRatio: false,

            plugins: {

                legend: {
                    display: false
                }

            },

            scales: {

                y: {

                    beginAtZero: false,

                    suggestedMin: 60,

                    suggestedMax: 100

                }

            }

        }

    });

}


// ==========================================
// PASS / FAIL CHART
// ==========================================

function createResultChart() {

    const canvas =
        document.getElementById("resultChart");

    if (!canvas) return;

    const ctx = canvas.getContext("2d");

    if (resultChart) {
        resultChart.destroy();
    }

    const passStudents = students.filter(
        student =>
            String(student.Result).toLowerCase() === "pass"
    ).length;

    const failStudents =
        students.length - passStudents;


    resultChart = new Chart(ctx, {

        type: "doughnut",

        data: {

            labels: [
                "Pass Students",
                "Fail Students"
            ],

            datasets: [{

                data: [
                    passStudents,
                    failStudents
                ],

                borderWidth: 1

            }]

        },

        options: {

            responsive: true,

            maintainAspectRatio: false,

            plugins: {

                legend: {
                    position: "bottom"
                }

            }

        }

    });

}


// ==========================================
// ATTENDANCE CHART
// ==========================================

function createAttendanceChart() {

    const canvas =
        document.getElementById("attendanceChart");

    if (!canvas) return;

    const ctx = canvas.getContext("2d");

    if (attendanceChart) {
        attendanceChart.destroy();
    }


    // Attendance ranges

    const below60 = students.filter(
        student => Number(student.Attendance) < 60
    ).length;

    const between60_75 = students.filter(
        student =>
            Number(student.Attendance) >= 60 &&
            Number(student.Attendance) < 75
    ).length;

    const between75_90 = students.filter(
        student =>
            Number(student.Attendance) >= 75 &&
            Number(student.Attendance) < 90
    ).length;

    const between90_100 = students.filter(
        student =>
            Number(student.Attendance) >= 90
    ).length;


    attendanceChart = new Chart(ctx, {

        type: "bar",

        data: {

            labels: [
                "Below 60%",
                "60% - 75%",
                "75% - 90%",
                "90% - 100%"
            ],

            datasets: [{

                label: "Number of Students",

                data: [
                    below60,
                    between60_75,
                    between75_90,
                    between90_100
                ],

                borderWidth: 1

            }]

        },

        options: {

            responsive: true,

            maintainAspectRatio: false,

            scales: {

                y: {

                    beginAtZero: true,

                    ticks: {
                        stepSize: 5
                    }

                }

            }

        }

    });

}

// ==========================================
// TOP PERFORMERS
// ==========================================

function createTopPerformers() {

    const container =
        document.getElementById("topPerformersList");

    if (!container) return;

    // Sort students by Average Marks
    const topStudents = [...students]
        .sort((a, b) =>
            Number(b.Average_Marks || 0) -
            Number(a.Average_Marks || 0)
        )
        .slice(0, 4);

    container.innerHTML = "";

    topStudents.forEach(student => {

        const name =
            student.Student_Name || "Unknown Student";

        const average =
            Number(student.Average_Marks || 0);

        const department =
            student.Department || "IT";

        const semester =
            student.Semester || "-";

        const firstLetter =
            name.charAt(0).toUpperCase();

        container.innerHTML += `
            <div class="student-row">

                <div class="student-avatar">
                    ${firstLetter}
                </div>

                <div class="student-details">
                    <strong>${name}</strong>
                    <span>${department} • Semester ${semester}</span>
                </div>

                <strong class="student-score">
                    ${average.toFixed(1)}%
                </strong>

            </div>
        `;
    });
}
// ==========================================
// STUDENTS TABLE
// ==========================================

function displayStudentsTable(data = students) {

    const tableBody =
        document.getElementById("studentsTableBody");

    if (!tableBody) return;

    tableBody.innerHTML = "";

    data.forEach(student => {

        const studentId =
            student.Student_ID || "";

        const studentName =
            student.Student_Name || "Unknown";

        const gender =
            student.Gender || "-";

        const department =
            student.Department || "-";

        const semester =
            student.Semester || "-";

        const averageMarks =
            Number(student.Average_Marks || 0);

        const attendance =
            Number(student.Attendance || 0);

        const result =
            student.Result || "-";

        const resultClass =
            String(result).toLowerCase() === "pass"
                ? "result-pass"
                : "result-fail";

        const row = document.createElement("tr");

        row.innerHTML = `
            <td>${studentId}</td>

            <td>${studentName}</td>

            <td>${gender}</td>

            <td>${department}</td>

            <td>${semester}</td>

            <td>${averageMarks.toFixed(1)}%</td>

            <td>${attendance.toFixed(1)}%</td>

            <td>
                <span class="${resultClass}">
                    ${result}
                </span>
            </td>
        `;

        tableBody.appendChild(row);
    });
}
// ==========================================
// STUDENT SEARCH & FILTER
// ==========================================

function filterStudents() {

    const searchInput =
        document.getElementById("studentSearch");

    const resultFilter =
        document.getElementById("resultFilter");

    const departmentFilter =
        document.getElementById("departmentFilter");

    const searchValue =
        searchInput.value.toLowerCase().trim();

    const selectedResult =
        resultFilter.value;

    const selectedDepartment =
        departmentFilter.value;


    const filteredStudents =
        students.filter(student => {

            const studentId =
                String(student.Student_ID || "")
                    .toLowerCase();

            const studentName =
                String(student.Student_Name || "")
                    .toLowerCase();

            const result =
                String(student.Result || "");

            const department =
                String(student.Department || "");


            const matchesSearch =
                studentId.includes(searchValue) ||
                studentName.includes(searchValue);

            const matchesResult =
                selectedResult === "All" ||
                result === selectedResult;

            const matchesDepartment =
                selectedDepartment === "All" ||
                department === selectedDepartment;


            return (
                matchesSearch &&
                matchesResult &&
                matchesDepartment
            );

        });


    displayStudentsTable(filteredStudents);
}


// ==========================================
// SEARCH EVENT
// ==========================================

const studentSearch =
    document.getElementById("studentSearch");

if (studentSearch) {

    studentSearch.addEventListener(
        "input",
        filterStudents
    );

}


// ==========================================
// RESULT FILTER EVENT
// ==========================================

const resultFilter =
    document.getElementById("resultFilter");

if (resultFilter) {

    resultFilter.addEventListener(
        "change",
        filterStudents
    );

}


// ==========================================
// DEPARTMENT FILTER EVENT
// ==========================================

const departmentFilter =
    document.getElementById("departmentFilter");

if (departmentFilter) {

    departmentFilter.addEventListener(
        "change",
        filterStudents
    );

}
// ==========================================
// START
// ==========================================

loadExcelData();

console.log(
    "Student Performance Dashboard Loaded Successfully!"
);