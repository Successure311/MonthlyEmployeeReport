// ============================================
// APP.JS - Employee Attendance Dashboard
// ============================================

const API_BASE = "";
let currentData = null;
let leaveData = null; // Stored aggregated leave data
let charts = {};

// ============================================
// INITIALIZATION
// ============================================
document.addEventListener("DOMContentLoaded", () => {
  loadReports();
  loadLeaves(); // Pre-load leaves for export and page switch
  setupEventListeners();
});

function setupEventListeners() {
  document
    .getElementById("reportSelect")
    .addEventListener("change", onReportChange);
  document
    .getElementById("employeeSelect")
    .addEventListener("change", onEmployeeFilter);
  document.getElementById("refreshBtn").addEventListener("click", onRefresh);
  document.getElementById("addReportBtn").addEventListener("click", () => {
    document.getElementById("fileInput").click();
  });
  document.getElementById("fileInput").addEventListener("change", onFileUpload);
  document.getElementById("menuToggle").addEventListener("click", () => {
    document.getElementById("sidebar").classList.toggle("open");
  });
  document.getElementById("exportExcelBtn").addEventListener("click", handleExportExcel);
  document.getElementById("themeToggle").addEventListener("click", toggleTheme);
  
  // Navigation
  document.getElementById("navDashboard").addEventListener("click", () => switchPage("dashboard"));
  document.getElementById("navLeave").addEventListener("click", () => switchPage("leave"));
  
  // Initialize Theme
  const savedTheme = localStorage.getItem("theme");
  if (savedTheme === "light") {
    document.body.classList.add("light-theme");
    updateThemeUI(true);
  }
}

function toggleTheme() {
  const isLight = document.body.classList.toggle("light-theme");
  localStorage.setItem("theme", isLight ? "light" : "dark");
  updateThemeUI(isLight);
  
  // Refresh charts to update grid/text colors if needed
  if (currentData) {
    onEmployeeFilter(); // Re-renders whichever view is active
  }
}

function updateThemeUI(isLight) {
  const themeIcon = document.getElementById("themeIcon");
  themeIcon.textContent = isLight ? "light_mode" : "dark_mode";
  
  // Update Chart defaults based on theme
  const textColor = isLight ? "#4a5568" : "#8b8fa3";
  const mutedColor = isLight ? "#718096" : "#5a5e73";
  const gridColor = isLight ? "rgba(0,0,0,0.05)" : "rgba(255,255,255,0.04)";
  
  chartDefaults.plugins.legend.labels.color = textColor;
  chartDefaults.scales.x.title.color = mutedColor;
  chartDefaults.scales.x.ticks.color = mutedColor;
  chartDefaults.scales.x.grid.color = gridColor;
  chartDefaults.scales.y.title.color = mutedColor;
  chartDefaults.scales.y.ticks.color = mutedColor;
  chartDefaults.scales.y.grid.color = gridColor;
}

// ============================================
// API CALLS
// ============================================
async function loadReports() {
  try {
    const res = await fetch(`${API_BASE}/api/reports?v=${Date.now()}`);
    const data = await res.json();
    const select = document.getElementById("reportSelect");
    select.innerHTML = '<option value="">-- Choose Month --</option>';
    data.reports.forEach((r) => {
      const opt = document.createElement("option");
      opt.value = r.filename;
      opt.textContent = r.month;
      select.appendChild(opt);
    });
  } catch (err) {
    console.error("Error loading reports:", err);
  }
}

async function loadLeaves() {
  showLoading(true);
  try {
    const res = await fetch(`${API_BASE}/api/leaves?v=${Date.now()}`);
    const data = await res.json();
    if (data.error) {
      alert(data.error);
      showLoading(false);
      return;
    }
    leaveData = data;
    renderLeavePage(data);
  } catch (err) {
    console.error("Error loading leaves:", err);
    alert("Failed to load leave data.");
  }
  showLoading(false);
}

function switchPage(pageId) {
  const dashboardContent = document.getElementById("dashboardContent");
  const leaveContent = document.getElementById("leaveContent");
  const welcomeState = document.getElementById("welcomeState");
  const reportSelectSection = document.getElementById("reportSelectSection");
  const employeeFilterSection = document.getElementById("employeeFilterSection");
  const pageTitle = document.getElementById("pageTitle");
  const dateRangeLabel = document.getElementById("dateRangeLabel");

  // Update Nav buttons
  document.querySelectorAll(".nav-btn").forEach((btn) => btn.classList.remove("active"));

  if (pageId === "dashboard") {
    document.getElementById("navDashboard").classList.add("active");
    leaveContent.style.display = "none";
    reportSelectSection.style.display = "";
    if (currentData) {
      dashboardContent.style.display = "";
      welcomeState.style.display = "none";
      employeeFilterSection.style.display = "";
      pageTitle.textContent = `${currentData.month} ${currentData.year} - Attendance Analysis`;
      dateRangeLabel.textContent = currentData.dateRange || "";
    } else {
      dashboardContent.style.display = "none";
      welcomeState.style.display = "";
      employeeFilterSection.style.display = "none";
      pageTitle.textContent = "Employee Attendance Dashboard";
      dateRangeLabel.textContent = "";
    }
  } else if (pageId === "leave") {
    document.getElementById("navLeave").classList.add("active");
    dashboardContent.style.display = "none";
    welcomeState.style.display = "none";
    leaveContent.style.display = "";
    reportSelectSection.style.display = "none";
    employeeFilterSection.style.display = "none";
    pageTitle.textContent = "Combined Leave Tracking";
    dateRangeLabel.textContent = "Consolidated Professional Leave Report";
    loadLeaves();
  }
}

async function loadAnalysis(filename) {
  showLoading(true);
  try {
    const res = await fetch(
      `${API_BASE}/api/analysis/${encodeURIComponent(filename)}?v=${Date.now()}`,
    );
    const data = await res.json();
    if (data.error) {
      alert(data.error);
      showLoading(false);
      return;
    }
    currentData = data;
    console.log("API Response employees[0]:", data.employees[0]);
    console.log("saturdayCount:", data.employees[0]?.saturdayCount, "sundayCount:", data.employees[0]?.sundayCount);
    renderDashboard(data);
    document.getElementById("exportExcelBtn").style.display = "flex";
  } catch (err) {
    console.error("Error loading analysis:", err);
    alert("Failed to load analysis data.");
  }
  showLoading(false);
}

async function handleExportExcel() {
  if (!currentData) return;
  
  const empSelect = document.getElementById("employeeSelect");
  const selectedName = empSelect.value;
  
  let exportData = {
    month: currentData.month,
    year: currentData.year,
    numDays: currentData.numDays,
    employees: [],
    leaveData: leaveData // Include leave data for export
  };

  if (selectedName === "all") {
    exportData.employees = currentData.employees;
  } else {
    const emp = currentData.employees.find(e => e.name === selectedName);
    if (!emp) return;
    exportData.employees = [emp];
  }

  showLoading(true);
  try {
    const response = await fetch(`${API_BASE}/api/export_excel`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(exportData)
    });

    if (!response.ok) throw new Error("Failed to export Excel");

    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Attendance_Report_${currentData.month}_${currentData.year}.xlsx`;
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);
    a.remove();
  } catch (err) {
    console.error("Export Error:", err);
    alert("Failed to export Excel. Please try again.");
  }
  showLoading(false);
}

// ============================================
// EVENT HANDLERS
// ============================================
function onReportChange() {
  const filename = document.getElementById("reportSelect").value;
  if (!filename) {
    document.getElementById("welcomeState").style.display = "";
    document.getElementById("dashboardContent").style.display = "none";
    return;
  }
  loadAnalysis(filename);
}

function onEmployeeFilter() {
  if (!currentData) return;
  const empName = document.getElementById("employeeSelect").value;
  if (empName === "all") {
    renderDashboard(currentData);
  } else {
    const emp = currentData.employees.find((e) => e.name === empName);
    if (emp) {
      renderSingleEmployee(emp, currentData);
    }
  }
}

function onRefresh() {
  loadReports();
  const filename = document.getElementById("reportSelect").value;
  if (filename) loadAnalysis(filename);
}

async function onFileUpload(e) {
  const file = e.target.files[0];
  if (!file) return;

  const formData = new FormData();
  formData.append("file", file);

  showLoading(true);
  try {
    const res = await fetch(`${API_BASE}/api/upload`, {
      method: "POST",
      body: formData,
    });
    const data = await res.json();
    if (data.error) {
      alert(data.error);
    } else {
      alert("Report added successfully!");
      await loadReports();
      // Select the new report
      document.getElementById("reportSelect").value = data.filename;
      onReportChange();
    }
  } catch (err) {
    console.error("Error uploading file:", err);
    alert("Failed to upload report.");
  }
  showLoading(false);
  // Reset input
  e.target.value = "";
}

// ============================================
// RENDERING
// ============================================
function showLoading(show) {
  document.getElementById("loadingOverlay").style.display = show
    ? "flex"
    : "none";
}

function renderDashboard(data) {
  document.getElementById("welcomeState").style.display = "none";
  document.getElementById("dashboardContent").style.display = "";
  document.getElementById("dateRangeLabel").textContent = data.dateRange || "";
  document.getElementById("pageTitle").textContent =
    `${data.month} ${data.year} - Attendance Analysis`;

  // Populate employee filter
  const empSelect = document.getElementById("employeeSelect");
  empSelect.innerHTML = '<option value="all">All Employees</option>';
  data.employees.forEach((emp) => {
    const opt = document.createElement("option");
    opt.value = emp.name;
    opt.textContent = emp.name;
    empSelect.appendChild(opt);
  });
  document.getElementById("employeeFilterSection").style.display = "";

  // Calculate totals
  let totalPresent = 0,
    totalAbsent = 0,
    totalWO = 0,
    totalLate = 0,
    totalEarlyCount = 0,
    totalSat = 0,
    totalSun = 0,
    totalNS = 0,
    totalHalfDay = 0;
  let totalDurMin = 0,
    totalLateMins = 0,
    totalEarlyMins = 0,
    totalLateOutMins = 0,
    totalLateOutCount = 0,
    maxLateOutMinTotal = 0;

  data.employees.forEach((emp) => {
    totalPresent += emp.presentCount;
    totalAbsent += emp.absentCount;
    totalWO += emp.weeklyOffCount;
    totalLate += emp.lateCount;
    totalEarlyCount += emp.earlyDepartureCount || 0;
    totalSat += emp.saturdayCount || 0;
    totalSun += emp.sundayCount || 0;
    totalNS += emp.nsShiftCount || 0;
    totalHalfDay += emp.halfDayCount || 0;
    totalDurMin += emp.totalDurationMinutes;
    totalLateMins += emp.totalLateMinutes;
    totalEarlyMins += emp.totalEarlyMinutes;
    totalLateOutMins += emp.totalLateOutMinutes || 0;
    totalLateOutCount += emp.lateOutCount || 0;
    const empMaxLO = parseTimeToMinutes(emp.maxLateOut);
    if (empMaxLO > maxLateOutMinTotal) maxLateOutMinTotal = empMaxLO;
  });

  const avgDurPerEmp =
    data.employees.length > 0 ? totalDurMin / data.employees.length : 0;

  // KPI Cards
  renderKPIs([
    {
      label: "Total Employees",
      value: data.employees.length,
      cls: "blue",
      sub: `For ${data.month} ${data.year}`,
    },
    {
      label: "Avg Present Days",
      value: data.employees.length
        ? Math.round(totalPresent / data.employees.length)
        : 0,
      cls: "green",
      sub: `of ${data.numDays} days`,
    },
    {
      label: "Avg Absent | Half",
      value: data.employees.length
        ? `${Math.round(totalAbsent / data.employees.length)} | ${Math.round(totalHalfDay / data.employees.length)}`
        : "0 | 0",
      cls: "red",
      sub: `Per employee`,
    },
    {
      label: "Late Arrival",
      value: `${totalLate} ( Avg. ${minutesToHHMM(totalLate ? totalLateMins / totalLate : 0)} )`,
      cls: "orange",
      sub: `Across all employees`,
    },
    {
      label: "Early Departure",
      value: `${totalEarlyCount} ( Avg. ${minutesToHHMM(totalEarlyCount ? totalEarlyMins / totalEarlyCount : 0)} )`,
      cls: "purple",
      sub: `Across all employees`,
    },
    
    {
      label: "Avg Work Hours",
      value: minutesToHHMM(
        avgDurPerEmp / (totalPresent / data.employees.length || 1),
      ),
      cls: "cyan",
      sub: `Per present day`,
    },
    {
      label: "SAT | SUN",
      value: `${data.employees.length ? Math.round(totalSat / data.employees.length) : 0} | ${data.employees.length ? Math.round(totalSun / data.employees.length) : 0}`,
      cls: "cyan",
      sub: `Avg Saturday/Sunday`,
    },
    {
      label: "Avg. Late Out",
      value: minutesToHHMM(totalLateOutCount ? totalLateOutMins / totalLateOutCount : 0),
      cls: "pink",
      sub: `Per extra stay day`,
    },
    {
      label: "Max Late Out",
      value: minutesToHHMM(maxLateOutMinTotal),
      cls: "red",
      sub: `Across all employees`,
    },
  ]);

  // Charts
  renderAttendanceChart(data.employees);
  renderDurationChart(data.employees);
  renderLateChart(data.employees);
  renderEarlyChart(data.employees);

  // Summary Table
  renderSummaryTable(data.employees);

  // Hide daily detail
  document.getElementById("dailyDetailCard").style.display = "none";
}

function renderSingleEmployee(emp, data) {
  // KPI Cards for single employee
  renderKPIs([
    {
      label: "Present Days",
      value: emp.presentCount,
      cls: "green",
      sub: `of ${data.numDays} days`,
    },
    {
      label: "Absent | Half",
      value: `${emp.absentCount} | ${emp.halfDayCount || 0}`,
      cls: "red",
      sub: `Days missed`,
    },
    {
      label: "Holidays",
      value: emp.weeklyOffCount,
      cls: "purple",
      sub: `Rest days`,
    },
    {
      label: "Late Arrival",
      value: `${emp.lateCount} ( Avg. ${emp.avgLateBy} )`,
      cls: "orange",
      sub: `of ${emp.presentCount} present`,
    },
    {
      label: "Early Departure",
      value: `${emp.earlyDepartureCount} ( Avg. ${emp.avgEarlyBy} )`,
      cls: "purple",
      sub: `of ${emp.presentCount} present`,
    },
    {
      label: "Total Work Hours",
      value: emp.totalDuration,
      cls: "blue",
      sub: `Total accumulated`,
    },
    {
      label: "Avg Daily Hours",
      value: emp.avgDuration,
      cls: "cyan",
      sub: `Per present day`,
    },
    {
      label: "SAT | SUN",
      value: `${emp.saturdayCount || 0} | ${emp.sundayCount || 0}`,
      cls: "cyan",
      sub: `In this month`,
    },
    {
      label: "Avg. Late Out",
      value: emp.avgLateOut,
      cls: "pink",
      sub: `Per extra stay day`,
    },
    {
      label: "Max Late Out",
      value: emp.maxLateOut,
      cls: "red",
      sub: `Maximum extra time`,
    },
  ]);

  // Single employee charts - daily duration
  renderSingleDurationChart(emp);
  renderSingleLateChart(emp);
  renderSingleAttendanceDonut(emp, data);
  renderSingleEarlyChart(emp);

  // Summary Table with just one employee
  renderSummaryTable([emp]);

  // Show daily detail
  renderDailyDetail(emp);
}

// ============================================
// LEAVE PAGE RENDERING
// ============================================
function renderLeavePage(data) {
  const header = document.getElementById("leaveSummaryHeader");
  const body = document.getElementById("leaveSummaryBody");
  
  // Clear header except first two (Employee, Total)
  while (header.cells.length > 2) {
    header.deleteCell(2);
  }
  
  // Add Month headers
  data.months.forEach(month => {
    const th = document.createElement("th");
    th.textContent = month;
    header.appendChild(th);
  });
  
  // Populate body
  body.innerHTML = data.leaves.map(emp => {
    let monthlyHtml = data.months.map(m => `<td>${emp.monthly[m] || 0}</td>`).join("");
    return `
      <tr onclick="showLeaveDetails('${emp.name}', ${JSON.stringify(emp.dates).replace(/"/g, '&quot;')})">
        <td>${emp.name}</td>
        <td style="font-weight: bold; color: var(--accent-red);">${emp.total}</td>
        ${monthlyHtml}
      </tr>
    `;
  }).join("");
  
  // Hide detail card by default
  document.getElementById("leaveDetailCard").style.display = "none";
}

window.showLeaveDetails = function(name, dates) {
  const card = document.getElementById("leaveDetailCard");
  const title = document.getElementById("leaveDetailTitle");
  const body = document.getElementById("leaveDetailBody");
  
  title.textContent = `Leave Dates: ${name} (Total ${dates.length})`;
  
  body.innerHTML = dates.map((d, i) => {
    try {
        // Robust parsing: d.month can be "Jan-2026", "January-2026", "Jan", etc.
        const monthMap = {
          'jan': 0, 'january': 0,
          'feb': 1, 'february': 1,
          'mar': 2, 'march': 2,
          'apr': 3, 'april': 3,
          'may': 4,
          'jun': 5, 'june': 5,
          'jul': 6, 'july': 6,
          'aug': 7, 'august': 7,
          'sep': 8, 'september': 8,
          'oct': 9, 'october': 9,
          'nov': 10, 'november': 10,
          'dec': 11, 'december': 11
        };

        const monthFull = d.month || "";
        const parts = monthFull.split('-');
        const monStr = parts[0].toLowerCase();
        let yearNum = parts.length > 1 ? parseInt(parts[1]) : new Date().getFullYear();
        
        if (isNaN(yearNum)) yearNum = new Date().getFullYear();
        
        const mIdx = monthMap[monStr] !== undefined ? monthMap[monStr] : 0;
        
        // Extract number from "Jan 14, Wednesday"
        const dayMatch = (d.date || "").match(/\d+/);
        const dNum = dayMatch ? parseInt(dayMatch[0]) : 1;

        const dateObj = new Date(yearNum, mIdx, dNum);
        
        if (isNaN(dateObj.getTime())) throw new Error("Invalid Date Object");

        const formattedDate = dateObj.toLocaleDateString('en-GB', { 
            day: '2-digit', 
            month: 'short', 
            year: 'numeric' 
        });
        const dayName = dateObj.toLocaleDateString('en-US', { weekday: 'long' });

        return `
          <tr>
            <td>${i + 1}</td>
            <td style="font-weight: 600; font-size: 1rem;"><span class="status-badge status-A">${formattedDate}</span></td>
            <td style="font-weight: 500;">${dayName}</td>
          </tr>
        `;
    } catch (err) {
        console.error("Date parsing error:", err, d);
        return `
          <tr>
            <td>${i + 1}</td>
            <td style="font-weight: 600; font-size: 1rem;"><span class="status-badge status-A">${d.date} ${d.month}</span></td>
            <td style="font-weight: 500;">-</td>
          </tr>
        `;
    }
  }).join("");
  
  card.style.display = "block";
  card.scrollIntoView({ behavior: 'smooth' });
}

// ============================================
// KPI CARDS
// ============================================
function renderKPIs(items) {
  const grid = document.getElementById("kpiGrid");
  grid.innerHTML = items
    .map(
      (item) => `
        <div class="kpi-card ${item.cls}">
            <div class="kpi-label">${item.label}</div>
            <div class="kpi-value">${item.value}</div>
            <div class="kpi-sub">${item.sub}</div>
        </div>
    `,
    )
    .join("");
}

// ============================================
// CHART HELPERS
// ============================================
function destroyChart(key) {
  if (charts[key]) {
    charts[key].destroy();
    charts[key] = null;
  }
}

const chartColors = {
  blue: "#4f7cff",
  cyan: "#00d4ff",
  green: "#2dd4a8",
  orange: "#ff9f43",
  red: "#ff5c5c",
  purple: "#a855f7",
  pink: "#f472b6",
};

const chartDefaults = {
  responsive: true,
  maintainAspectRatio: false,
  plugins: {
    legend: {
      labels: {
        color: "#8b8fa3",
        font: { family: "Inter", size: 11 },
      },
    },
  },
  scales: {
    x: {
      title: {
        display: true,
        text: "Date",
        color: "#5a5e73",
        font: { family: "Inter", size: 10 },
      },
      ticks: {
        color: "#5a5e73",
        font: { family: "Inter", size: 10 },
        callback: function (val, index) {
          const label = this.getLabelForValue(val);
          if (label && label.includes(",")) {
            return label.split(",")[0];
          }
          return label;
        },
      },
      grid: { color: "rgba(255,255,255,0.04)" },
    },
    y: {
      title: {
        display: true,
        text: "Time",
        color: "#5a5e73",
        font: { family: "Inter", size: 10 },
      },
      ticks: { color: "#5a5e73", font: { family: "Inter", size: 10 } },
      grid: { color: "rgba(255,255,255,0.04)" },
    },
  },
};

// ============================================
// ALL-EMPLOYEES CHARTS
// ============================================
function renderAttendanceChart(employees) {
  destroyChart("attendance");
  const ctx = document.getElementById("attendanceChart").getContext("2d");
  charts.attendance = new Chart(ctx, {
    type: "bar",
    data: {
      labels: employees.map((e) => e.name),
      datasets: [
        {
          label: "Present",
          data: employees.map((e) => e.presentCount),
          backgroundColor: chartColors.green,
          borderRadius: 6,
          barPercentage: 0.6,
        },
        {
          label: "Absent",
          data: employees.map((e) => e.absentCount),
          backgroundColor: chartColors.red,
          borderRadius: 6,
          barPercentage: 0.6,
        },
        {
          label: "Holidays",
          data: employees.map((e) => e.weeklyOffCount),
          backgroundColor: chartColors.purple,
          borderRadius: 6,
          barPercentage: 0.6,
        },
        {
          label: "Saturday",
          data: employees.map((e) => e.saturdayCount || 0),
          backgroundColor: chartColors.blue,
          borderRadius: 6,
          barPercentage: 0.6,
        },
        {
          label: "Sunday",
          data: employees.map((e) => e.sundayCount || 0),
          backgroundColor: chartColors.pink,
          borderRadius: 6,
          barPercentage: 0.6,
        },
      ],
    },
    options: {
      ...chartDefaults,
      plugins: {
        ...chartDefaults.plugins,
        title: { display: false },
      },
      scales: {
        ...chartDefaults.scales,
        x: {
          ...chartDefaults.scales.x,
          title: { display: true, text: "Date", color: "#5a5e73", font: { family: "Inter", size: 10 } }
        },
        y: {
          ...chartDefaults.scales.y,
          title: { display: true, text: "Time", color: "#5a5e73", font: { family: "Inter", size: 10 } }
        }
      }
    },
  });

  // Clear custom legend for multi-employee view
  const legendContainer = document.getElementById("attendanceLegend");
  if (legendContainer) legendContainer.innerHTML = "";
}

function renderDurationChart(employees) {
  destroyChart("duration");
  const ctx = document.getElementById("durationChart").getContext("2d");
  charts.duration = new Chart(ctx, {
    type: "bar",
    data: {
      labels: employees.map((e) => e.name),
      datasets: [
        {
          label: "Avg Hours",
          data: employees.map((e) => e.avgDurationMinutes / 60),
          backgroundColor: createGradient(
            ctx,
            chartColors.cyan,
            chartColors.blue,
          ),
          borderRadius: 6,
          barPercentage: 0.5,
        },
      ],
    },
    options: {
      ...chartDefaults,
      plugins: {
        ...chartDefaults.plugins,
        legend: { display: false },
      },
      scales: {
        ...chartDefaults.scales,
        x: {
          ...chartDefaults.scales.x,
          title: { display: true, text: "Date", color: "#5a5e73", font: { family: "Inter", size: 10 } }
        },
        y: {
          ...chartDefaults.scales.y,
          title: { display: true, text: "Time", color: "#5a5e73", font: { family: "Inter", size: 10 } }
        }
      }
    },
  });
}

function renderLateChart(employees) {
  destroyChart("late");
  const ctx = document.getElementById("lateChart").getContext("2d");
  charts.late = new Chart(ctx, {
    type: "bar",
    data: {
      labels: employees.map((e) => e.name),
      datasets: [
        {
          label: "Avg Late (mins)",
          data: employees.map((e) => e.avgLateMinutes),
          backgroundColor: chartColors.orange,
          borderRadius: 6,
          barPercentage: 0.5,
        },
      ],
    },
    options: {
      ...chartDefaults,
      plugins: { ...chartDefaults.plugins, legend: { display: false } },
      scales: {
        ...chartDefaults.scales,
        x: {
          ...chartDefaults.scales.x,
          title: { display: true, text: "Date", color: "#5a5e73", font: { family: "Inter", size: 10 } }
        },
        y: {
          ...chartDefaults.scales.y,
          title: { display: true, text: "Time", color: "#5a5e73", font: { family: "Inter", size: 10 } }
        }
      }
    },
  });
}

function renderEarlyChart(employees) {
  destroyChart("early");
  const ctx = document.getElementById("earlyChart").getContext("2d");
  charts.early = new Chart(ctx, {
    type: "bar",
    data: {
      labels: employees.map((e) => e.name),
      datasets: [
        {
          label: "Avg Early (mins)",
          data: employees.map((e) => e.avgEarlyMinutes),
          backgroundColor: chartColors.pink,
          borderRadius: 6,
          barPercentage: 0.5,
        },
      ],
    },
    options: {
      ...chartDefaults,
      plugins: { ...chartDefaults.plugins, legend: { display: false } },
      scales: {
        ...chartDefaults.scales,
        x: {
          ...chartDefaults.scales.x,
          title: { display: true, text: "Date", color: "#5a5e73", font: { family: "Inter", size: 10 } }
        },
        y: {
          ...chartDefaults.scales.y,
          title: { display: true, text: "Time", color: "#5a5e73", font: { family: "Inter", size: 10 } }
        }
      }
    },
  });
}

// ============================================
// SINGLE EMPLOYEE CHARTS
// ============================================
function renderSingleDurationChart(emp) {
  destroyChart("duration");
  const ctx = document.getElementById("durationChart").getContext("2d");
  const presentDays = emp.dailyData.filter((d) => d.status === "P" && (!d.shift || d.shift.toUpperCase() !== "HO"));
  charts.duration = new Chart(ctx, {
    type: "line",
    data: {
      labels: presentDays.map((d) => d.day),
      datasets: [
        {
          label: "Work Hours",
          data: presentDays.map((d) => parseTimeToMinutes(d.duration) / 60),
          borderColor: chartColors.cyan,
          backgroundColor: "rgba(0,212,255,0.08)",
          fill: true,
          tension: 0.4,
          pointRadius: 4,
          pointBackgroundColor: chartColors.cyan,
        },
      ],
    },
    options: {
      ...chartDefaults,
      plugins: { ...chartDefaults.plugins, legend: { display: false } },
      scales: {
        ...chartDefaults.scales,
        x: {
          ...chartDefaults.scales.x,
          title: { display: true, text: "Date", color: "#5a5e73", font: { family: "Inter", size: 10 } }
        },
        y: {
          ...chartDefaults.scales.y,
          title: { display: true, text: "Time", color: "#5a5e73", font: { family: "Inter", size: 10 } }
        }
      }
    },
  });
  document.getElementById("durationChartCard").querySelector("h3").textContent =
    `${emp.name} - Daily Work Hours`;
}

function renderSingleLateChart(emp) {
  destroyChart("late");
  const ctx = document.getElementById("lateChart").getContext("2d");
  const presentDays = emp.dailyData.filter((d) => d.status === "P" && (!d.shift || d.shift.toUpperCase() !== "HO"));
  charts.late = new Chart(ctx, {
    type: "bar",
    data: {
      labels: presentDays.map((d) => d.day),
      datasets: [
        {
          label: "Late By (mins)",
          data: presentDays.map((d) => parseTimeToMinutes(d.lateBy)),
          backgroundColor: presentDays.map((d) =>
            parseTimeToMinutes(d.lateBy) > 60
              ? chartColors.red
              : chartColors.orange,
          ),
          borderRadius: 4,
          barPercentage: 0.7,
        },
      ],
    },
    options: {
      ...chartDefaults,
      plugins: { ...chartDefaults.plugins, legend: { display: false } },
      scales: {
        ...chartDefaults.scales,
        x: {
          ...chartDefaults.scales.x,
          title: { display: true, text: "Date", color: "#5a5e73", font: { family: "Inter", size: 10 } }
        },
        y: {
          ...chartDefaults.scales.y,
          title: { display: true, text: "Time", color: "#5a5e73", font: { family: "Inter", size: 10 } }
        }
      }
    },
  });
  document.getElementById("lateChartCard").querySelector("h3").textContent =
    `${emp.name} - Daily Late Arrivals`;
}

function renderSingleAttendanceDonut(emp, data) {
  destroyChart("attendance");
  const ctx = document.getElementById("attendanceChart").getContext("2d");
  charts.attendance = new Chart(ctx, {
    type: "doughnut",
    data: {
      labels: ["Present", "Absent", "Holidays", "Saturday", "Sunday"],
      datasets: [
        {
          data: [
            emp.presentCount, 
            emp.absentCount, 
            emp.weeklyOffCount,
            emp.saturdayCount || 0,
            emp.sundayCount || 0
          ],
          backgroundColor: [
            chartColors.green,
            chartColors.red,
            chartColors.purple,
            chartColors.blue,
            chartColors.pink,
          ],
          borderColor: "#1c1e2e",
          borderWidth: 3,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: "65%",
      plugins: {
        legend: {
          display: false // Hide built-in legend as we use custom one
        },
      },
    },
  });

  // Populate Custom Legend
  const legendContainer = document.getElementById("attendanceLegend");
  if (legendContainer) {
    const data = [
      { label: "Present", value: emp.presentCount, color: chartColors.green },
      { label: "Absent", value: emp.absentCount, color: chartColors.red },
      { label: "Holidays", value: emp.weeklyOffCount, color: chartColors.purple },
      { label: "Saturday", value: emp.saturdayCount || 0, color: chartColors.blue },
      { label: "Sunday", value: emp.sundayCount || 0, color: chartColors.pink },
    ];
    
    legendContainer.innerHTML = data.map(item => `
        <div class="legend-item">
            <div class="legend-dot" style="background: ${item.color}"></div>
            <div class="legend-label">${item.label}</div>
            <div class="legend-value">${item.value}</div>
        </div>
    `).join("");
  }
  document
    .getElementById("attendanceChartCard")
    .querySelector("h3").textContent = `${emp.name} - Attendance Breakdown`;
}

function renderSingleEarlyChart(emp) {
  destroyChart("early");
  const ctx = document.getElementById("earlyChart").getContext("2d");
  const presentDays = emp.dailyData.filter((d) => d.status === "P" && (!d.shift || d.shift.toUpperCase() !== "HO"));
  charts.early = new Chart(ctx, {
    type: "bar",
    data: {
      labels: presentDays.map((d) => d.day),
      datasets: [
        {
          label: "Early By (mins)",
          data: presentDays.map((d) => parseTimeToMinutes(d.earlyBy)),
          backgroundColor: chartColors.pink,
          borderRadius: 4,
          barPercentage: 0.7,
        },
      ],
    },
    options: {
      ...chartDefaults,
      plugins: { ...chartDefaults.plugins, legend: { display: false } },
      scales: {
        ...chartDefaults.scales,
        x: {
          ...chartDefaults.scales.x,
          title: { display: true, text: "Date", color: "#5a5e73", font: { family: "Inter", size: 10 } }
        },
        y: {
          ...chartDefaults.scales.y,
          title: { display: true, text: "Time", color: "#5a5e73", font: { family: "Inter", size: 10 } }
        }
      }
    },
  });
  document.getElementById("earlyChartCard").querySelector("h3").textContent =
    `${emp.name} - Daily Early Departure`;
}

// ============================================
// TABLES
// ============================================
function renderSummaryTable(employees) {
  const tbody = document.getElementById("employeeSummaryBody");
  tbody.innerHTML = employees
    .map(
      (emp) => `
        <tr onclick="selectEmployee('${emp.name}')">
            <td>${emp.name}</td>
            <td><span class="status-badge status-P">${emp.presentCount}</span></td>
            <td><span class="status-badge status-A">${emp.absentCount} | ${emp.halfDayCount || 0}</span></td>
            <td><span class="status-badge status-WO">${emp.weeklyOffCount}</span></td>
            <td style="color:${emp.lateCount > 15 ? chartColors.red : chartColors.orange}">${emp.lateCount}</td>
            <td style="color:${chartColors.green}">${emp.onTimeCount}</td>
            <td>${emp.totalDuration}</td>
            <td style="font-weight:600;">${emp.avgDuration}</td>
            <td style="color:${parseTimeToMinutes(emp.avgLateBy) > 0 ? chartColors.red : chartColors.orange}">${emp.avgLateBy}</td>
            <td style="color:${parseTimeToMinutes(emp.avgEarlyBy) > 0 ? chartColors.red : chartColors.pink}">${emp.avgEarlyBy}</td>
            <td style="color:${parseTimeToMinutes(emp.avgLateOut) > 0 ? chartColors.green : ""}">${emp.avgLateOut}</td>
            <td style="color:${parseTimeToMinutes(emp.maxLateOut) > 0 ? chartColors.green : ""}">${emp.maxLateOut}</td>
        </tr>
    `,
    )
    .join("");
}

function renderDailyDetail(emp) {
  const card = document.getElementById("dailyDetailCard");
  card.style.display = "";
  document.getElementById("dailyDetailTitle").textContent =
    `${emp.name} - Daily Attendance Detail`;

  const tbody = document.getElementById("dailyDetailBody");
  tbody.innerHTML = emp.dailyData
    .map((d) => {
      const statusCls = `status-${d.status}`;
      const hasPunch = (d.inTime && d.inTime !== '-') || (d.outTime && d.outTime !== '-');
      let rowCls = (d.shift && d.shift.toUpperCase() === 'HO' && hasPunch) ? 'row-HO' : '';
      if (d.isHalfDay) rowCls += ' row-half-day';
      return `
            <tr class="${rowCls}">
                <td>${d.day}</td>
                <td><span class="status-badge ${statusCls}">${d.status || "-"}</span></td>
                <td style="color:${d.isForgotPunchIn ? chartColors.red : ''}">${d.inTime || "-"}</td>
                <td style="color:${d.isForgotPunchOut ? chartColors.red : ''}">${d.outTime || "-"}</td>
                <td>${d.duration || "-"}</td>
                <td style="color:${parseTimeToMinutes(d.lateBy) > 0 ? chartColors.red : ""}">${d.lateBy || "-"}</td>
                <td style="color:${parseTimeToMinutes(d.earlyBy) > 0 ? chartColors.red : ""}">${d.earlyBy || "-"}</td>
                <td style="color:${parseTimeToMinutes(d.lateOut) > 0 ? chartColors.green : ""}">${d.lateOut || "-"}</td>
                <td>${d.shift || "-"}</td>
            </tr>
        `;
    })
    .join("");
}

// ============================================
// UTILITIES
// ============================================
function selectEmployee(name) {
  document.getElementById("employeeSelect").value = name;
  onEmployeeFilter();
}

function parseTimeToMinutes(timeStr) {
  if (!timeStr || timeStr === "00:00" || timeStr === "-") return 0;
  const parts = timeStr.split(":");
  if (parts.length < 2) return 0;
  return parseInt(parts[0]) * 60 + parseInt(parts[1]);
}

function minutesToHHMM(minutes) {
  if (!minutes || minutes === 0) return "00:00";
  const h = Math.floor(minutes / 60);
  const m = Math.round(minutes % 60);
  if (h === 0) {
    return `${m} min`;
  }
  return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")} h`;
}

function createGradient(ctx, color1, color2) {
  const gradient = ctx.createLinearGradient(0, 0, 0, 280);
  gradient.addColorStop(0, color1);
  gradient.addColorStop(1, color2);
  return gradient;
}
