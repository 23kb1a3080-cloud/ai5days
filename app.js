/**
 * AI-Powered Microsoft Excel Workshop - Comprehensive Application Engine
 * N.B.K.R. Institute of Science & Technology (NBKRIST)
 */

// ==========================================================================
// 1. DEFAULT DATA STATE (Exact OCR extraction from 12-page document)
// ==========================================================================
const DEFAULT_DATA = {
  profile: {
    studentName: "P.SRINIVAS",
    rollNo: "23KB1A3080",
    branch: "AI & DS – A Section / B.Tech 4th Year",
    college: "N.B.K.R. Institute of Science and Technology (NBKRIST)",
    labName: "AI-Powered Microsoft Excel Laboratory",
    acadYear: "2025-2026",
    staffInCharge: "Dr. K. Ramanjaneyulu, M.Tech, Ph.D",
    examiner: "Prof. M. Srinivasulu",
    dates: {
      day1: "18 / 08 / 2026",
      day2: "19 / 08 / 2026",
      day3: "20 / 08 / 2026",
      day4: "21 / 08 / 2026",
      day5: "22 / 08 / 2026"
    }
  },
  days: [
    {
      id: "day1",
      dayNumber: 1,
      title: "Excel basics, tables, formulas, functions, and Microsoft Copilot integrations.",
      focus: "Excel basics, tables, formulas, functions, and Microsoft Copilot integrations.",
      session1: "Excel basics and creating tables",
      session2: "Formulas, functions, and Microsoft Copilot",
      aiTools: "Microsoft Copilot in Excel, ChatGPT for formula generation and syntax debugging",
      excelFeatures: "Structured Tables, Cell Referencing ($A$1 Absolute vs Relative), Data Validation dropdown lists, Basic statistical functions (SUM, AVERAGE, COUNT, COUNTA, MAX, MIN), Conditional Formatting (Data Bars & Color Scales), and AutoFilter.",
      handsOn: "Created a comprehensive Student Academic & Attendance Tracker table with 50 sample records, configured column data types, applied data validation dropdowns for departments, and set up dynamic conditional formatting rules.",
      practicalTask: "Calculated total marks, average percentages, and grade categories using nested formulas; verified accuracy against manual test computations; and used Microsoft Copilot prompts to generate conditional IF statements for student pass/fail status.",
      whatILearned: "Learned how Excel structured tables automatically expand formula references and formatting to new rows. Understood the critical distinction between absolute and relative cell referencing during autofill, and explored how Copilot can suggest optimal formula syntax from plain-English prompts.",
      difficulties: "Faced formula calculation errors when copying formulas across columns due to incorrect relative cell referencing, and encountered circular reference warnings when writing range totals.",
      howSolved: "Used Excel's Formula Auditing tools (Trace Precedents) to locate misreferenced cells, locked fixed benchmark values using the dollar sign ($) absolute references, and consulted Copilot to structure clean table-column references.",
      confidence: "Intermediate", // Beginner, Intermediate, Advanced
      rating: 5,
      suggestions: "The step-by-step approach to explaining tables and formula syntax was very effective. A short quiz at the end of each session would help reinforce the concepts learned.",
      reflections: [
        {
          q: "What new Excel/AI concepts did I learn today?",
          a: "Today I learned how to create structured Excel tables with automatic calculated columns, implement Data Validation dropdown lists, apply conditional formatting with color scales, master absolute vs. relative cell referencing ($A$1 vs A1), and utilize Microsoft Copilot for natural-language formula generation."
        },
        {
          q: "Which AI tool did I use and for what purpose?",
          a: "I used Microsoft Copilot in Excel to generate formulas for weighted average calculations, explain formula syntax errors in plain English, and generate conditional logic prompts like 'Write a formula to assign Grade A if marks >= 85, Grade B if marks >= 70, otherwise Grade C'."
        },
        {
          q: "What practical task did I complete?",
          a: "I built a 50-student academic performance tracker containing student names, roll numbers, subject marks, calculated totals, weighted grades, and attendance percentages, enhanced with multi-level sorting and conditional formatting."
        },
        {
          q: "What difficulty did I face and how did I overcome it?",
          a: "I encountered circular reference warnings and unexpected values when copying formulas across columns due to improper relative referencing. I solved this by examining cell references, anchoring fixed weights with absolute references ($), and testing calculation steps incrementally using Formula Auditing."
        },
        {
          q: "How will I apply today's learning in a real-world situation?",
          a: "I can apply these Excel skills to maintain class attendance records, calculate semester marks, prepare budget sheets for college events, and organize any structured data in a clear and professional format."
        }
      ]
    },
    {
      id: "day2",
      dayNumber: 2,
      title: "Data cleaning workflows, Power Query mechanics, and advanced lookup or logical functions.",
      focus: "Data cleaning workflows, Power Query mechanics, and advanced lookup or logical functions.",
      session1: "Data cleaning and Power Query",
      session2: "Lookup and logical functions",
      aiTools: "Microsoft Copilot, Power Query M-Code Assistant",
      excelFeatures: "Power Query ETL pipeline (Extract, Transform, Load), Text functions (TRIM, PROPER, TEXTSPLIT, CONCAT), Advanced Lookup functions (XLOOKUP, VLOOKUP, INDEX-MATCH), and Logical functions (IFS, SWITCH, nested IF-AND).",
      handsOn: "Imported an unformatted, messy multi-department employee dataset with trailing whitespace, inconsistent phone/date formats, and missing fields; cleaned and normalized the dataset using automated Power Query transformation steps.",
      practicalTask: "Constructed an automated Power Query data-cleaning recipe to eliminate blank rows, trim whitespace, and standardize date formats. Connected separate salary and department tables using XLOOKUP to dynamically retrieve employee compensation records without #N/A errors.",
      whatILearned: "Understood the non-destructive nature of Power Query Applied Steps and how transformations can be re-executed on refreshed data. Learned why XLOOKUP is more resilient and versatile than VLOOKUP for two-way lookups, and recognized the necessity of data hygiene prior to analysis.",
      difficulties: "Encountered date parsing errors in Power Query due to mixed regional date formats (DD/MM/YYYY vs MM/DD/YYYY), and experienced #N/A errors when VLOOKUP attempted to look leftward.",
      howSolved: "Configured the Power Query column type transformation 'Using Locale: English (India)' to parse ambiguous date strings accurately, and transitioned from legacy VLOOKUP to XLOOKUP which naturally supports bidirectional lookups and built-in error handling.",
      confidence: "Intermediate",
      rating: 5,
      suggestions: "Providing a comparison chart showing when to use VLOOKUP versus XLOOKUP versus INDEX-MATCH would have been helpful. A real dataset for practice made the session much more engaging than working with dummy numbers.",
      reflections: [
        {
          q: "What new Excel/AI concepts did I learn today?",
          a: "Today I learned automated data cleaning using Power Query (ETL pipeline), text cleaning functions (TRIM, PROPER, TEXTSPLIT), multi-condition logical operations using IFS and SWITCH, and modern data lookup techniques using XLOOKUP and INDEX-MATCH."
        },
        {
          q: "Which AI tool did I use and for what purpose?",
          a: "I used Microsoft Copilot to generate Power Query M-code transformation expressions for custom date formatting and to write complex nested IFS expressions for categorizing employee salary bands and performance tiers."
        },
        {
          q: "What practical task did I complete?",
          a: "I cleaned and transformed an unstandardized raw employee dataset with 200+ records using Power Query, normalized text cases, eliminated duplicates, and built an automated lookup sheet linking employee IDs to department details and pay scales using XLOOKUP."
        },
        {
          q: "What difficulty did I face and how did I overcome it?",
          a: "I faced #N/A lookup errors when matching text-formatted IDs with numeric IDs, as well as date parse errors in Power Query. I resolved this by enforcing consistent data types (Text) across both tables in Power Query and setting the if_not_found parameter in XLOOKUP."
        },
        {
          q: "How will I apply today's learning in a real-world situation?",
          a: "I can use Power Query to clean and standardize survey data or attendance records before presenting them. XLOOKUP can help me cross-reference student roll numbers with marks from different sheets, and IF functions can automate pass or fail categorization."
        }
      ]
    },
    {
      id: "day3",
      dayNumber: 3,
      title: "Pivot Tables, Pivot Charts, AI-based descriptive data analysis, and automated insights.",
      focus: "Pivot Tables, Pivot Charts, AI-based descriptive data analysis, and automated insights.",
      session1: "Pivot Tables and summarizing data",
      session2: "Pivot Charts and AI-based descriptive analysis",
      aiTools: "Microsoft Copilot in Excel (Analyze Data & Natural Language Query Engine)",
      excelFeatures: "Pivot Table field arrangement (Rows, Columns, Values, Filters), Value Field Settings (Sum, Average, Count, % of Grand Total, Running Total), Custom Calculated Fields, Linked Pivot Charts, Interactive Slicers, and Timelines.",
      handsOn: "Summarized a 500-row multi-region retail transaction dataset into multi-dimensional Pivot Tables to evaluate regional revenue distribution, top-performing product categories, and monthly profit growth trends.",
      practicalTask: "Created interconnected Pivot Tables with custom Calculated Fields for Profit Margin %, built linked interactive Pivot Charts (Clustered Column and Line charts), configured synchronized multi-table Slicers, and used Copilot to generate automated executive summaries.",
      whatILearned: "Discovered how Pivot Tables aggregate complex datasets instantaneously without cumbersome manual formulas. Learned how to customize value representations as percentages of parent totals, and how Copilot can automatically discover trends, outliers, and correlations in summarized records.",
      difficulties: "Confused when dragging a numeric field into Values and Excel defaulted to 'Count' instead of 'Sum', and encountered disconnected slicers that only filtered one Pivot Table instead of all related tables on the sheet.",
      howSolved: "Cleaned blank and non-numeric characters from the source data column so Excel recognized it as numeric (allowing 'Sum'), and connected the Slicers across all Pivot Tables by configuring 'Report Connections' in the Slicer toolbar.",
      confidence: "Intermediate",
      rating: 5,
      suggestions: "Showing a before-and-after comparison of raw data versus a Pivot Table summary made the purpose very clear. It would be useful to also cover how to refresh a Pivot Table when the source data changes.",
      reflections: [
        {
          q: "What new Excel/AI concepts did I learn today?",
          a: "Today I learned how to build multi-dimensional Pivot Tables, configure Value Field Settings (% of Grand Total, Running Totals), create custom Calculated Fields, design linked Pivot Charts, and utilize Copilot's 'Analyze Data' feature for automated statistical insights."
        },
        {
          q: "Which AI tool did I use and for what purpose?",
          a: "I used Microsoft Copilot to run descriptive data analysis queries, ask natural-language questions such as 'Which product category drove the highest profit margin in Q3?', and generate instant natural language summary cards for executive review."
        },
        {
          q: "What practical task did I complete?",
          a: "I created a complete retail sales summary workbook with 3 dynamic Pivot Tables aggregating regional sales performance, product category profitability, and quarterly growth, supplemented by interactive Pivot Charts and multi-table synchronized slicers."
        },
        {
          q: "What difficulty did I face and how did I overcome it?",
          a: "I struggled with disconnected slicers that only filtered one table, and unexpected 'Count' aggregations caused by blank cells. I resolved this by enabling 'Report Connections' for all tables on the sheet and sanitizing the source range to ensure pure numeric data types."
        },
        {
          q: "How will I apply today's learning in a real-world situation?",
          a: "I can use Pivot Tables to summarize semester-wise marks across multiple subjects, create department-wise attendance reports, or analyze survey response distributions. Pivot Charts will help me present these summaries visually during project presentations."
        }
      ]
    },
    {
      id: "day4",
      dayNumber: 4,
      title: "Interactive data visualizations, advanced KPIs, dashboard architecture, Power Query, and Power Pivot structures.",
      focus: "Interactive data visualizations, advanced KPIs, dashboard architecture, Power Query, and Power Pivot structures.",
      session1: "Interactive data visualizations and KPIs",
      session2: "Dashboard structure, Power Query, and Power Pivot",
      aiTools: "Microsoft Copilot, DAX Assistant for Power Pivot Data Modeling",
      excelFeatures: "Power Pivot Data Modeling (Star Schema & Relationships), DAX Measures (CALCULATE, RELATED, SUMX, DIVIDE), Custom KPI Metric Summary Cards, Dynamic Slicers & Timelines, Visual Hierarchy, and Dashboard Layout Design.",
      handsOn: "Architected and constructed an interactive executive Business Performance & Academic KPI Dashboard integrating multiple relational tables (Students, Courses, Attendance, Fees) into a single cohesive visual interface.",
      practicalTask: "Built a Power Pivot data model linking 3 relational tables, authored DAX measures for Year-over-Year Growth Rate and Attendance % with safe error handling, created dynamic KPI cards with trend indicators, and aligned interactive charts into a clean dashboard view.",
      whatILearned: "Understood the architectural advantages of relational Data Models over flat spreadsheets. Learned how DAX calculates dynamic measures within varying filter contexts, and gained insight into visual hierarchy, uncluttered dashboard principles, and interactive storytelling with data.",
      difficulties: "Struggled with division-by-zero errors in DAX percentage growth calculations, and faced relationship creation failures in Power Pivot due to duplicate foreign keys in the dimension table.",
      howSolved: "Replaced standard division operators with the DIVIDE(num, den, 0) DAX function to safely manage divide-by-zero scenarios, and used Power Query's 'Remove Duplicates' on the key column before creating one-to-many relationships in Power Pivot.",
      confidence: "Advanced",
      rating: 5,
      suggestions: "The live demonstration of building a dashboard from scratch was the most helpful part. Including a template with placeholder KPI cards that students can fill would make the activity even more hands-on.",
      reflections: [
        {
          q: "What new Excel/AI concepts did I learn today?",
          a: "Today I learned Power Pivot relational data modeling (creating 1-to-Many relationships), authoring dynamic DAX measures (SUMX, CALCULATE, DIVIDE), creating custom KPI summary cards with trend sparklines, and designing executive dashboard architecture."
        },
        {
          q: "Which AI tool did I use and for what purpose?",
          a: "I used Microsoft Copilot to formulate and debug DAX expressions for calculating percentage metrics with filter context overrides, and to recommend the most effective visualization types (e.g., Donut vs. Clustered Bar vs. Line trend) for specific KPI metrics."
        },
        {
          q: "What practical task did I complete?",
          a: "I developed an interactive Executive KPI Dashboard that consolidates student performance, attendance metrics, and departmental benchmarks into high-impact visual cards, dynamic line charts, and interactive slicers powered by a Power Pivot data model."
        },
        {
          q: "What difficulty did I face and how did I overcome it?",
          a: "I faced relationship errors in Power Pivot due to duplicate primary keys in the lookup table, and chart formatting distortions when slicers were toggled. I resolved this by deduplicating the lookup dimension in Power Query and locking chart properties to 'Don't move or size with cells'."
        },
        {
          q: "How will I apply today's learning in a real-world situation?",
          a: "I can build a student performance dashboard to track marks, attendance, and assignment submission rates across all subjects for an entire class. In a professional setting, the same approach can create executive summaries that update automatically when the underlying data changes."
        }
      ]
    },
    {
      id: "day5",
      dayNumber: 5,
      title: "Workflow automation, Office Scripts, Capstone project implementation, and synthesis of overall learning.",
      focus: "Workflow automation, Office Scripts, Capstone project implementation, and synthesis of overall learning.",
      session1: "Workflow automation and Office Scripts",
      session2: "Capstone project implementation and overall review",
      aiTools: "Microsoft Copilot, Office Scripts Code Editor (TypeScript Engine)",
      excelFeatures: "Office Scripts automation engine (TypeScript APIs for Excel), Macro recorder vs. TypeScript Office Scripts, Batch data formatting automation, Capstone Project synthesis (ETL -> Modeling -> Analysis -> Visuals -> Automation), and Export workflows.",
      handsOn: "Wrote and deployed an end-to-end Office Script in TypeScript that automates raw dataset importation, applies standardized academic report styling, computes summary KPI metrics, generates a dynamic chart, and exports the clean report with a single click.",
      practicalTask: "Executed the complete 5-Day Capstone Project by integrating Power Query data ingestion, Power Pivot relational modeling, an interactive KPI dashboard, and an automated Office Script trigger button for continuous automated reporting.",
      whatILearned: "Learned how TypeScript-based Office Scripts run reliably and securely across Excel Web and Desktop environments without VBA macro compatibility issues. Synthesized all workshop disciplines into an automated, end-to-end enterprise data analytics pipeline.",
      difficulties: "Encountered TypeScript type-checking errors when accessing cell range values (mixing string and number types), and struggled with dynamically finding the last row index in datasets with variable lengths.",
      howSolved: "Implemented TypeScript type casting using Number() and String() conversions, and utilized range.getUsedRange().getRowCount() to dynamically calculate sheet boundaries with guidance from Microsoft Copilot prompts.",
      confidence: "Advanced",
      rating: 5,
      suggestions: "The capstone project was the most rewarding activity of the entire workshop because it required applying everything learned over five days in one coherent task. Sharing example capstone outputs from previous batches would motivate students to aim higher.",
      reflections: [
        {
          q: "What new Excel/AI concepts did I learn today?",
          a: "Today I learned Office Scripts programming using TypeScript to automate repetitive Excel tasks, the differences between cloud-native Office Scripts and legacy VBA, and how to synthesize the entire 5-day toolchain (Power Query, Power Pivot, DAX, Dashboards, and Automation) into a unified Capstone project."
        },
        {
          q: "Which AI tool did I use and for what purpose?",
          a: "I used Microsoft Copilot to write, debug, and optimize TypeScript code for Office Scripts, generating custom automation functions for table formatting, dynamic column calculations, and automated report generation."
        },
        {
          q: "What practical task did I complete?",
          a: "I successfully built and delivered the comprehensive Workshop Capstone Project, creating an automated analytics pipeline that ingests raw multi-sheet records, cleans and models the data, updates an executive KPI dashboard, and automates report generation via an Office Script button."
        },
        {
          q: "What difficulty did I face and how did I overcome it?",
          a: "I encountered syntax errors with Excel TypeScript range methods when attempting to dynamically find the last used row in a variable dataset. I resolved this using range.getUsedRange() and prompting Copilot to refine the boundary detection logic."
        },
        {
          q: "How will I apply today's learning in a real-world situation?",
          a: "I can automate monthly academic report generation so that formatting, chart updates, and sheet organization happen automatically each time new data is added. In an internship or job, Office Scripts would eliminate hours of manual Excel work by running automated routines on scheduled data files."
        }
      ]
    }
  ]
};

// Application State
let appData = null;
let currentActiveDayTab = "overview";
let currentTheme = "dark";
let chartInstanceDay3 = null;
let chartInstanceDay4 = null;

const STORAGE_KEY = "nbkrist_excel_workshop_data_v3";

// ==========================================================================
// 2. INITIALIZATION & STORAGE
// ==========================================================================

function initApp() {
  // Load state from localStorage or use default
  const savedData = localStorage.getItem(STORAGE_KEY);
  if (savedData) {
    try {
      appData = JSON.parse(savedData);
    } catch (e) {
      console.warn("Invalid localStorage data, resetting to default", e);
      appData = JSON.parse(JSON.stringify(DEFAULT_DATA));
    }
  } else {
    // Check if legacy storage had customized profile details to preserve
    appData = JSON.parse(JSON.stringify(DEFAULT_DATA));
    const legacyKeys = ["nbkrist_excel_workshop_data_v2", "nbkrist_excel_workshop_data"];
    for (const k of legacyKeys) {
      const legacyRaw = localStorage.getItem(k);
      if (legacyRaw) {
        try {
          const legacyObj = JSON.parse(legacyRaw);
          if (legacyObj && legacyObj.profile) {
            appData.profile = { ...DEFAULT_DATA.profile, ...legacyObj.profile };
          }
          break;
        } catch (err) {}
      }
    }
    localStorage.setItem(STORAGE_KEY, JSON.stringify(appData));
  }

  // Load theme preference
  const savedTheme = localStorage.getItem("nbkrist_theme") || "dark";
  setTheme(savedTheme);

  // Sync quick bar with profile
  updateQuickBar();

  // Render initial studio tab and printable pages
  renderStudioTab(currentActiveDayTab);
  renderDocumentPages();
}

function saveDataToStorage() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(appData));
  updateQuickBar();
  renderDocumentPages();
  showToast("Changes saved successfully", "success");
}

function updateQuickBar() {
  const p = appData.profile;
  document.getElementById("bar-student-name").textContent = p.studentName || "—";
  document.getElementById("bar-roll-no").textContent = p.rollNo || "—";
  document.getElementById("bar-branch").textContent = p.branch || "—";
  document.getElementById("bar-college").textContent = p.college || "—";
}

// ==========================================================================
// 3. NAVIGATION & VIEW SWITCHING
// ==========================================================================

function switchMainView(view) {
  const studioView = document.getElementById("view-studio");
  const docView = document.getElementById("view-document");
  const tabStudio = document.getElementById("tab-studio-view");
  const tabDoc = document.getElementById("tab-document-view");

  if (view === "studio") {
    studioView.classList.add("active");
    docView.classList.remove("active");
    tabStudio.classList.add("active");
    tabDoc.classList.remove("active");
  } else {
    studioView.classList.remove("active");
    docView.classList.add("active");
    tabStudio.classList.remove("active");
    tabDoc.classList.add("active");
    // Re-render pages to ensure latest edits are shown
    renderDocumentPages();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }
}

function selectDayTab(tabId) {
  currentActiveDayTab = tabId;
  document.querySelectorAll(".day-tab-btn").forEach(btn => {
    btn.classList.remove("active");
  });
  event?.currentTarget?.classList.add("active");
  renderStudioTab(tabId);
}

function handlePrintDocument() {
  // Ensure document mode is active before printing
  switchMainView('document');
  setTimeout(() => {
    window.print();
  }, 300);
}

function toggleTheme() {
  const nextTheme = currentTheme === "dark" ? "light" : "dark";
  setTheme(nextTheme);
}

function setTheme(theme) {
  currentTheme = theme;
  document.documentElement.setAttribute("data-theme", theme);
  localStorage.setItem("nbkrist_theme", theme);
  const icon = document.querySelector("#theme-toggle-btn i");
  if (icon) {
    icon.className = theme === "dark" ? "fa-solid fa-sun" : "fa-solid fa-moon";
  }
}

// ==========================================================================
// 4. STUDIO VIEW RENDERER (Left Report Form + Right Live Simulators)
// ==========================================================================

function renderStudioTab(tabId) {
  const container = document.getElementById("studio-tab-content");
  if (!container) return;

  if (tabId === "overview") {
    renderOverviewTab(container);
    return;
  }

  const dayIndex = parseInt(tabId.replace("day", "")) - 1;
  const day = appData.days[dayIndex];
  if (!day) return;

  const dateVal = appData.profile.dates[tabId] || "____ / ____ / 2026";

  container.innerHTML = `
    <div class="day-grid">
      <!-- LEFT COLUMN: EDITABLE DAILY REPORT -->
      <div class="day-col">
        
        <!-- Header Banner Card -->
        <div class="card-box">
          <div class="card-header-flex">
            <h3><i class="fa-solid fa-calendar-check text-blue"></i> Day ${day.dayNumber} Official Report</h3>
            <span class="badge badge-primary">Session Date: ${dateVal}</span>
          </div>

          <div class="info-field-group">
            <div class="info-item">
              <span class="info-item-label">Focus &amp; Title</span>
              <div class="info-item-value editable-area" contenteditable="true" onblur="updateDayField(${dayIndex}, 'title', this.innerText)">${escapeHtml(day.title)}</div>
            </div>

            <div class="info-item">
              <span class="info-item-label">Session 1 Topic</span>
              <div class="info-item-value editable-area" contenteditable="true" onblur="updateDayField(${dayIndex}, 'session1', this.innerText)">${escapeHtml(day.session1)}</div>
            </div>

            <div class="info-item">
              <span class="info-item-label">Session 2 Topic</span>
              <div class="info-item-value editable-area" contenteditable="true" onblur="updateDayField(${dayIndex}, 'session2', this.innerText)">${escapeHtml(day.session2)}</div>
            </div>

            <div class="info-item">
              <span class="info-item-label">AI Tools Used</span>
              <div class="info-item-value editable-area" contenteditable="true" onblur="updateDayField(${dayIndex}, 'aiTools', this.innerText)">${escapeHtml(day.aiTools)}</div>
            </div>

            <div class="info-item">
              <span class="info-item-label">Excel Features Learned</span>
              <div class="info-item-value editable-area" contenteditable="true" onblur="updateDayField(${dayIndex}, 'excelFeatures', this.innerText)">${escapeHtml(day.excelFeatures)}</div>
            </div>

            <div class="info-item">
              <span class="info-item-label">Hands-on Activity Completed</span>
              <div class="info-item-value editable-area" contenteditable="true" onblur="updateDayField(${dayIndex}, 'handsOn', this.innerText)">${escapeHtml(day.handsOn)}</div>
            </div>

            <div class="info-item">
              <span class="info-item-label">Practical Task / Exercise</span>
              <div class="info-item-value editable-area" contenteditable="true" onblur="updateDayField(${dayIndex}, 'practicalTask', this.innerText)">${escapeHtml(day.practicalTask)}</div>
            </div>

            <div class="info-item">
              <span class="info-item-label">What I Learned Today</span>
              <div class="info-item-value editable-area" contenteditable="true" onblur="updateDayField(${dayIndex}, 'whatILearned', this.innerText)">${escapeHtml(day.whatILearned)}</div>
            </div>
          </div>
        </div>

        <!-- Difficulties & Feedback Card -->
        <div class="card-box">
          <div class="card-header-flex">
            <h3><i class="fa-solid fa-wrench text-amber"></i> Difficulties &amp; Problem Solving</h3>
          </div>

          <div class="info-field-group mb-4">
            <div class="info-item">
              <span class="info-item-label">Difficulties Faced</span>
              <div class="info-item-value editable-area" contenteditable="true" onblur="updateDayField(${dayIndex}, 'difficulties', this.innerText)">${escapeHtml(day.difficulties)}</div>
            </div>

            <div class="info-item">
              <span class="info-item-label">How I Solved the Difficulties</span>
              <div class="info-item-value editable-area" contenteditable="true" onblur="updateDayField(${dayIndex}, 'howSolved', this.innerText)">${escapeHtml(day.howSolved)}</div>
            </div>
          </div>

          <h4 class="mb-2"><i class="fa-solid fa-sliders text-emerald"></i> Evaluation &amp; Confidence</h4>
          
          <div class="confidence-selector">
            <button class="confidence-pill ${day.confidence === 'Beginner' ? 'active' : ''}" onclick="setConfidence(${dayIndex}, 'Beginner')">Beginner</button>
            <button class="confidence-pill ${day.confidence === 'Intermediate' ? 'active' : ''}" onclick="setConfidence(${dayIndex}, 'Intermediate')">Intermediate</button>
            <button class="confidence-pill ${day.confidence === 'Advanced' ? 'active' : ''}" onclick="setConfidence(${dayIndex}, 'Advanced')">Advanced</button>
          </div>

          <div class="mt-4">
            <span class="info-item-label">Workshop Rating</span>
            <div class="star-rating-box" id="star-box-${dayIndex}">
              ${renderStarRating(dayIndex, day.rating)}
            </div>
          </div>

          <div class="info-item mt-4">
            <span class="info-item-label">Suggestions / Feedback</span>
            <div class="info-item-value editable-area" contenteditable="true" onblur="updateDayField(${dayIndex}, 'suggestions', this.innerText)">${escapeHtml(day.suggestions)}</div>
          </div>
        </div>

        <!-- Reflection Questions Card -->
        <div class="card-box">
          <div class="card-header-flex">
            <h3><i class="fa-solid fa-lightbulb text-cyan"></i> Short Daily Reflection Questions (5 Q&amp;As)</h3>
          </div>

          <div class="reflection-list">
            ${day.reflections.map((rf, rIdx) => `
              <div class="reflection-card">
                <div class="reflection-question">
                  <span>${rIdx + 1}.</span>
                  <div class="editable-area" contenteditable="true" onblur="updateReflection(${dayIndex}, ${rIdx}, 'q', this.innerText)">${escapeHtml(rf.q)}</div>
                </div>
                <div class="reflection-answer">
                  <strong>Ans:</strong>
                  <span class="editable-area" contenteditable="true" onblur="updateReflection(${dayIndex}, ${rIdx}, 'a', this.innerText)">${escapeHtml(rf.a)}</span>
                </div>
              </div>
            `).join('')}
          </div>
        </div>

      </div>

      <!-- RIGHT COLUMN: LIVE INTERACTIVE EXCEL / AI SIMULATOR -->
      <div class="day-col">
        ${renderInteractiveSimulator(day.dayNumber)}
      </div>
    </div>
  `;

  // Attach dynamic widget listeners
  initSimulatorWidget(day.dayNumber);
}

function renderOverviewTab(container) {
  const p = appData.profile;
  container.innerHTML = `
    <div class="overview-grid" style="margin-bottom: 2rem;">
      <!-- Left Column: Syllabus & Summary -->
      <div class="card-box">
        <div class="card-header-flex">
          <h3><i class="fa-solid fa-book-open-reader text-blue"></i> 5-Day Workshop Comprehensive Curriculum</h3>
          <button class="btn btn-outline btn-sm" onclick="switchMainView('document')">
            <i class="fa-solid fa-print"></i> View Institutional PDF (12 Pages)
          </button>
        </div>

        <div class="workshop-syllabus-list">
          ${appData.days.map(d => `
            <div class="syllabus-item" onclick="selectDayTab('day${d.dayNumber}')">
              <div class="syllabus-day-badge">
                <span>0${d.dayNumber}</span>
                DAY
              </div>
              <div class="syllabus-info">
                <h4>Day ${d.dayNumber}: ${escapeHtml(d.title)}</h4>
                <p><strong>Sessions:</strong> ${escapeHtml(d.session1)} &bull; ${escapeHtml(d.session2)}</p>
                <p class="text-muted"><i class="fa-solid fa-microchip"></i> AI Tools: ${escapeHtml(d.aiTools)} | Confidence: <span class="text-emerald">${escapeHtml(d.confidence)}</span></p>
              </div>
            </div>
          `).join('')}
        </div>
      </div>

      <!-- Right Column: Certificate Preview Card -->
      <div class="card-box">
        <div class="card-header-flex">
          <h3><i class="fa-solid fa-award text-amber"></i> Official Bonafide Certificate</h3>
          <button class="btn btn-primary btn-sm" onclick="switchMainView('document')">
            <i class="fa-solid fa-print"></i> View / Print PDF
          </button>
        </div>

        <div style="background: #ffffff; padding: 0.75rem; border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.25); text-align: center;">
          <img src="assets/certificate.png" alt="Official Certificate Preview" style="width: 100%; max-height: 440px; object-fit: contain; display: block; margin: 0 auto; border-radius: 4px;" />
        </div>

        <div class="mt-4" style="text-align: center;">
          <button class="btn btn-outline" style="width: 100%;" onclick="switchMainView('document')">
            <i class="fa-solid fa-file-pdf"></i> View &amp; Print Full 12-Page Institutional Record
          </button>
        </div>
      </div>
    </div>

    <!-- Complete 5-Day Daily Student Reports Feed -->
    <div class="card-box" style="margin-bottom: 2rem;">
      <div class="card-header-flex">
        <h3><i class="fa-solid fa-layer-group text-cyan"></i> Complete 5-Day Daily Student Learning Logs &amp; Reflections</h3>
        <span class="badge badge-success">All 5 Days Loaded</span>
      </div>

      <div style="display: flex; flex-direction: column; gap: 2rem;">
        ${appData.days.map((day, dIdx) => {
          const dateVal = p.dates[`day${day.dayNumber}`] || "____ / ____ / 2026";
          return `
            <div style="background: var(--bg-surface); border: 1px solid var(--border-color); border-radius: var(--radius-md); padding: 1.5rem;">
              <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid var(--border-color); padding-bottom: 0.75rem; margin-bottom: 1rem; flex-wrap: wrap; gap: 0.5rem;">
                <h4 style="font-size: 1.15rem; color: var(--primary); margin: 0;">
                  <i class="fa-regular fa-calendar-check"></i> Day ${day.dayNumber}: ${escapeHtml(day.title)}
                </h4>
                <div style="display: flex; gap: 0.5rem; align-items: center;">
                  <span class="badge badge-primary">Session: ${dateVal}</span>
                  <button class="btn btn-outline btn-sm" onclick="selectDayTab('day${day.dayNumber}')">
                    <i class="fa-solid fa-laptop-code"></i> Open Day ${day.dayNumber} Studio Simulator
                  </button>
                </div>
              </div>

              <div class="form-grid mb-4">
                <div class="info-item">
                  <span class="info-item-label">Session 1 Topic</span>
                  <div class="info-item-value">${escapeHtml(day.session1)}</div>
                </div>
                <div class="info-item">
                  <span class="info-item-label">Session 2 Topic</span>
                  <div class="info-item-value">${escapeHtml(day.session2)}</div>
                </div>
                <div class="info-item">
                  <span class="info-item-label">AI Tools Used</span>
                  <div class="info-item-value">${escapeHtml(day.aiTools)}</div>
                </div>
                <div class="info-item">
                  <span class="info-item-label">Excel Features Learned</span>
                  <div class="info-item-value">${escapeHtml(day.excelFeatures)}</div>
                </div>
              </div>

              <div class="info-item mb-4">
                <span class="info-item-label">Hands-on Activity Completed</span>
                <div class="info-item-value">${escapeHtml(day.handsOn)}</div>
              </div>

              <div class="info-item mb-4">
                <span class="info-item-label">Practical Task / Exercise</span>
                <div class="info-item-value">${escapeHtml(day.practicalTask)}</div>
              </div>

              <div class="info-item mb-4">
                <span class="info-item-label">What I Learned</span>
                <div class="info-item-value">${escapeHtml(day.whatILearned)}</div>
              </div>

              <div class="form-grid mb-4">
                <div class="info-item">
                  <span class="info-item-label">Difficulties Faced</span>
                  <div class="info-item-value">${escapeHtml(day.difficulties)}</div>
                </div>
                <div class="info-item">
                  <span class="info-item-label">How I Solved the Difficulties</span>
                  <div class="info-item-value">${escapeHtml(day.howSolved)}</div>
                </div>
              </div>

              <!-- Reflection Q&A -->
              <h5 style="margin: 1.2rem 0 0.6rem 0; font-size: 0.95rem; color: #93c5fd;">
                <i class="fa-solid fa-circle-question"></i> Short Reflection Q&amp;As (Day ${day.dayNumber})
              </h5>
              <div class="reflection-list">
                ${day.reflections.map((rf, rIdx) => `
                  <div class="reflection-card">
                    <div class="reflection-question">${rIdx + 1}. ${escapeHtml(rf.q)}</div>
                    <div class="reflection-answer"><strong>Ans:</strong> ${escapeHtml(rf.a)}</div>
                  </div>
                `).join('')}
              </div>
            </div>
          `;
        }).join('')}
      </div>
    </div>
  `;
}

// ==========================================================================
// 5. INTERACTIVE SIMULATORS (EXCEL & AI TOOLS)
// ==========================================================================

function renderInteractiveSimulator(dayNumber) {
  switch (dayNumber) {
    case 1:
      return `
        <div class="simulator-panel">
          <div class="sim-header">
            <span class="sim-title"><i class="fa-solid fa-table-cells text-cyan"></i> Day 1 Interactive Formula Sandbox</span>
            <span class="badge badge-info">Excel + Copilot</span>
          </div>
          <p class="text-muted" style="font-size: 0.85rem;">
            Practice formulas (<code class="font-mono">SUM</code>, <code class="font-mono">AVERAGE</code>, <code class="font-mono">IF</code>) on live student records. Edit numbers below and watch calculations update instantly!
          </p>

          <div class="excel-table-wrapper">
            <table class="excel-grid" id="excel-sim-grid">
              <thead>
                <tr>
                  <th>Student ID</th>
                  <th>Name</th>
                  <th>Math (M1)</th>
                  <th>DS (M2)</th>
                  <th>AI Lab (M3)</th>
                  <th>Total (SUM)</th>
                  <th>Avg (AVG)</th>
                  <th>Status (IF)</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td>23KB1A3096</td>
                  <td>Shaik Haneefa</td>
                  <td><input type="number" class="score-input" value="92" oninput="recalcDay1Sim()" /></td>
                  <td><input type="number" class="score-input" value="88" oninput="recalcDay1Sim()" /></td>
                  <td><input type="number" class="score-input" value="95" oninput="recalcDay1Sim()" /></td>
                  <td class="calc-cell" id="d1-tot-1">275</td>
                  <td class="calc-cell" id="d1-avg-1">91.67</td>
                  <td class="calc-cell text-emerald" id="d1-stat-1">Distinction</td>
                </tr>
                <tr>
                  <td>23KB1A3097</td>
                  <td>A. Sai Krishna</td>
                  <td><input type="number" class="score-input" value="78" oninput="recalcDay1Sim()" /></td>
                  <td><input type="number" class="score-input" value="82" oninput="recalcDay1Sim()" /></td>
                  <td><input type="number" class="score-input" value="80" oninput="recalcDay1Sim()" /></td>
                  <td class="calc-cell" id="d1-tot-2">240</td>
                  <td class="calc-cell" id="d1-avg-2">80.00</td>
                  <td class="calc-cell text-blue" id="d1-stat-2">First Class</td>
                </tr>
                <tr>
                  <td>23KB1A3098</td>
                  <td>M. Sneha</td>
                  <td><input type="number" class="score-input" value="65" oninput="recalcDay1Sim()" /></td>
                  <td><input type="number" class="score-input" value="70" oninput="recalcDay1Sim()" /></td>
                  <td><input type="number" class="score-input" value="72" oninput="recalcDay1Sim()" /></td>
                  <td class="calc-cell" id="d1-tot-3">207</td>
                  <td class="calc-cell" id="d1-avg-3">69.00</td>
                  <td class="calc-cell text-amber" id="d1-stat-3">Pass</td>
                </tr>
              </tbody>
            </table>
          </div>

          <!-- Copilot Formula Generator Assistant -->
          <div class="copilot-box">
            <div class="copilot-header">
              <i class="fa-solid fa-sparkles"></i> Microsoft Copilot Formula Assistant
            </div>
            <div class="copilot-input-group">
              <input type="text" id="copilot-input-prompt" placeholder="Ask Copilot (e.g. 'Calculate letter grade based on Average')" />
              <button class="btn btn-primary btn-sm" onclick="runCopilotPrompt()">
                <i class="fa-solid fa-paper-plane"></i> Ask AI
              </button>
            </div>
            <div id="copilot-output-area" class="copilot-response">
              <strong>Copilot Formula:</strong> <code>=IFS(G2>=90, "O", G2>=80, "A+", G2>=70, "A", G2>=60, "B+", TRUE, "Pass")</code><br/>
              <em>Explanation: Evaluates student average score and dynamically maps to university grading scale.</em>
            </div>
          </div>
        </div>
      `;

    case 2:
      return `
        <div class="simulator-panel">
          <div class="sim-header">
            <span class="sim-title"><i class="fa-solid fa-filter-circle-xmark text-cyan"></i> Day 2 Power Query Data Cleaner</span>
            <span class="badge badge-info">ETL Pipeline</span>
          </div>
          <p class="text-muted" style="font-size: 0.85rem;">
            Experience automated ETL workflows in Power Query. Click through cleaning transformations to transform messy raw data into sanitized records!
          </p>

          <div class="pipeline-steps">
            <button class="pipeline-step-badge active" onclick="applyPowerQueryStep('raw')">1. Raw Messy Data</button>
            <button class="pipeline-step-badge" onclick="applyPowerQueryStep('trim')">2. Trim &amp; Clean Text</button>
            <button class="pipeline-step-badge" onclick="applyPowerQueryStep('nulls')">3. Handle Missing Values</button>
            <button class="pipeline-step-badge" onclick="applyPowerQueryStep('lookup')">4. Apply XLOOKUP &amp; Format</button>
          </div>

          <div class="excel-table-wrapper">
            <table class="excel-grid" id="powerquery-grid">
              <!-- Rendered via JS -->
            </table>
          </div>

          <div class="copilot-box" style="margin-top: 1rem;">
            <div class="copilot-header">
              <i class="fa-solid fa-code"></i> Power Query M-Code Pipeline
            </div>
            <pre class="font-mono" id="m-code-display" style="font-size: 0.78rem; color: #a5d6ff; margin: 0; white-space: pre-wrap;">
let
    Source = Csv.Document(File.Contents("StudentRecords.csv"),[Delimiter=","]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    #"Trimmed Text" = Table.TransformColumns(#"Promoted Headers",{{"StudentName", Text.Trim, type text}}),
    #"Replaced Value" = Table.ReplaceValue(#"Trimmed Text", null, 0, Replacer.ReplaceValue, {"Score"}),
    #"Merged Queries" = Table.NestedJoin(#"Replaced Value", {"DeptCode"}, DepartmentMaster, {"DeptCode"}, "DeptTable")
in
    #"Merged Queries"</pre>
          </div>
        </div>
      `;

    case 3:
      return `
        <div class="simulator-panel">
          <div class="sim-header">
            <span class="sim-title"><i class="fa-solid fa-chart-pie text-cyan"></i> Day 3 Pivot Table &amp; AI Descriptive Analytics</span>
            <span class="badge badge-info">Dynamic Summaries</span>
          </div>
          <p class="text-muted" style="font-size: 0.85rem;">
            Slice and summarize high-dimensional data without writing manual formulas. Switch aggregations to instantly update the Pivot Chart!
          </p>

          <div style="display: flex; gap: 0.5rem; margin-bottom: 1rem; flex-wrap: wrap;">
            <button class="btn btn-outline btn-sm active" id="pivot-btn-dept" onclick="updatePivotChart('Department')">Group by Department</button>
            <button class="btn btn-outline btn-sm" id="pivot-btn-year" onclick="updatePivotChart('Year')">Group by Academic Year</button>
            <button class="btn btn-outline btn-sm" id="pivot-btn-perf" onclick="updatePivotChart('Performance')">Group by Grade Band</button>
          </div>

          <div style="background: var(--bg-app); border: 1px solid var(--border-color); border-radius: var(--radius-md); padding: 1rem; height: 260px; position: relative;">
            <canvas id="pivotChartCanvas"></canvas>
          </div>

          <div class="copilot-box" style="margin-top: 1rem;">
            <div class="copilot-header">
              <i class="fa-solid fa-brain"></i> Copilot Automated Insights
            </div>
            <p id="pivot-ai-insight" style="font-size: 0.82rem; color: #93c5fd; margin: 0;">
              AI Insight: <strong>AI &amp; DS Department</strong> accounts for the highest average score (89.4%) with a 14% higher distinction rate compared to baseline.
            </p>
          </div>
        </div>
      `;

    case 4:
      return `
        <div class="simulator-panel">
          <div class="sim-header">
            <span class="sim-title"><i class="fa-solid fa-chart-line text-cyan"></i> Day 4 Executive KPI &amp; Power Pivot Dashboard</span>
            <span class="badge badge-info">Interactive KPIs</span>
          </div>

          <div class="kpi-cards-grid">
            <div class="kpi-metric-box">
              <span class="info-item-label">Total Student Cohort</span>
              <span class="kpi-val" id="kpi-cohort">180</span>
              <span class="kpi-sub"><i class="fa-solid fa-arrow-trend-up"></i> +12% YoY</span>
            </div>
            <div class="kpi-metric-box">
              <span class="info-item-label">Workshop Pass Rate</span>
              <span class="kpi-val" id="kpi-pass">98.5%</span>
              <span class="kpi-sub"><i class="fa-solid fa-check"></i> Benchmark Met</span>
            </div>
            <div class="kpi-metric-box">
              <span class="info-item-label">AI Copilot Adoption</span>
              <span class="kpi-val" id="kpi-copilot">94.2%</span>
              <span class="kpi-sub"><i class="fa-solid fa-bolt"></i> High Efficiency</span>
            </div>
            <div class="kpi-metric-box">
              <span class="info-item-label">Student Satisfaction</span>
              <span class="kpi-val" id="kpi-sat">5.0 / 5</span>
              <span class="kpi-sub text-amber"><i class="fa-solid fa-star"></i> Top Tier</span>
            </div>
          </div>

          <div style="background: var(--bg-app); border: 1px solid var(--border-color); border-radius: var(--radius-md); padding: 1rem; height: 230px; position: relative;">
            <canvas id="kpiTrendChart"></canvas>
          </div>
        </div>
      `;

    case 5:
      return `
        <div class="simulator-panel">
          <div class="sim-header">
            <span class="sim-title"><i class="fa-solid fa-terminal text-cyan"></i> Day 5 Office Scripts Automation Playground</span>
            <span class="badge badge-info">TypeScript Engine</span>
          </div>
          <p class="text-muted" style="font-size: 0.85rem;">
            Run automated TypeScript macros in Microsoft Excel to auto-format reports, calculate stats, and generate presentation charts with 1 click!
          </p>

          <div class="code-editor-box">
            <span class="code-comment">// Microsoft Excel Office Script Automation</span><br/>
            <span class="code-keyword">function</span> <span class="code-func">main</span>(workbook: <span class="code-string">ExcelScript.Workbook</span>) {<br/>
            &nbsp;&nbsp;<span class="code-keyword">const</span> sheet = workbook.getActiveWorksheet();<br/>
            &nbsp;&nbsp;<span class="code-keyword">const</span> table = sheet.getTables()[0];<br/>
            &nbsp;&nbsp;<span class="code-comment">// Auto-format header colors &amp; freeze rows</span><br/>
            &nbsp;&nbsp;table.getHeaderRowRange().getFormat().getFill().setColor(<span class="code-string">"#1e3a8a"</span>);<br/>
            &nbsp;&nbsp;table.getHeaderRowRange().getFormat().getFont().setColor(<span class="code-string">"#FFFFFF"</span>);<br/>
            &nbsp;&nbsp;<span class="code-comment">// Compute Grand Average across Day 1 to 5</span><br/>
            &nbsp;&nbsp;sheet.getRange(<span class="code-string">"H10"</span>).setFormula(<span class="code-string">"=AVERAGE(F2:F9)"</span>);<br/>
            &nbsp;&nbsp;console.log(<span class="code-string">"Automation completed: 5 Worksheets synthesized."</span>);<br/>
            }
          </div>

          <div style="display: flex; gap: 0.75rem; align-items: center;">
            <button class="btn btn-primary" id="btn-run-script" onclick="runOfficeScriptSimulation()">
              <i class="fa-solid fa-play"></i> Execute Office Script
            </button>
            <span id="script-status-text" class="text-muted" style="font-size: 0.82rem;">Ready to run script.</span>
          </div>

          <div id="script-console-output" class="copilot-box" style="display: none; margin-top: 1rem;">
            <div class="copilot-header"><i class="fa-solid fa-circle-check text-emerald"></i> Script Execution Log</div>
            <div class="font-mono" style="font-size: 0.8rem; color: #34d399;" id="script-log-text"></div>
          </div>
        </div>
      `;

    default:
      return "";
  }
}

function initSimulatorWidget(dayNumber) {
  if (dayNumber === 2) {
    applyPowerQueryStep('raw');
  } else if (dayNumber === 3) {
    setTimeout(() => {
      updatePivotChart('Department');
    }, 50);
  } else if (dayNumber === 4) {
    setTimeout(() => {
      renderKpiChart();
    }, 50);
  }
}

// Day 1 Calculator
function recalcDay1Sim() {
  const rows = document.querySelectorAll("#excel-sim-grid tbody tr");
  rows.forEach((row, i) => {
    const inputs = row.querySelectorAll(".score-input");
    const m1 = parseFloat(inputs[0].value) || 0;
    const m2 = parseFloat(inputs[1].value) || 0;
    const m3 = parseFloat(inputs[2].value) || 0;
    const total = m1 + m2 + m3;
    const avg = total / 3;

    let status = "Pass";
    let statusClass = "text-amber";
    if (avg >= 85) {
      status = "Distinction";
      statusClass = "text-emerald";
    } else if (avg >= 75) {
      status = "First Class";
      statusClass = "text-blue";
    }

    const totCell = document.getElementById(`d1-tot-${i + 1}`);
    const avgCell = document.getElementById(`d1-avg-${i + 1}`);
    const statCell = document.getElementById(`d1-stat-${i + 1}`);

    if (totCell) totCell.textContent = total;
    if (avgCell) avgCell.textContent = avg.toFixed(2);
    if (statCell) {
      statCell.textContent = status;
      statCell.className = `calc-cell ${statusClass}`;
    }
  });
}

function runCopilotPrompt() {
  const promptInput = document.getElementById("copilot-input-prompt");
  const output = document.getElementById("copilot-output-area");
  const query = promptInput ? promptInput.value.trim() : "";

  if (!query) {
    output.innerHTML = `<strong>Copilot Formula:</strong> <code>=IFS(G2>=90, "O", G2>=80, "A+", G2>=70, "A", G2>=60, "B+", TRUE, "Pass")</code><br/><em>Explanation: Evaluates student average score and dynamically maps to university grading scale.</em>`;
    return;
  }

  output.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Copilot analyzing request: "${escapeHtml(query)}"...`;
  setTimeout(() => {
    output.innerHTML = `
      <strong>Copilot Generated Formula:</strong> <code>=XLOOKUP(A2, MasterData!A:A, MasterData!F:F, "Not Found", 0)</code><br/>
      <em>Generated for: "${escapeHtml(query)}". Matches student ID against master schema and returns computed analytics.</em>
    `;
  }, 600);
}

// Day 2 Power Query Pipeline
function applyPowerQueryStep(step) {
  const grid = document.getElementById("powerquery-grid");
  if (!grid) return;

  document.querySelectorAll(".pipeline-step-badge").forEach(b => b.classList.remove("active"));
  const btn = event?.currentTarget;
  if (btn) btn.classList.add("active");

  if (step === 'raw') {
    grid.innerHTML = `
      <thead>
        <tr>
          <th>Roll No</th>
          <th>Student Name (Dirty)</th>
          <th>Lab Score</th>
          <th>Dept Code</th>
        </tr>
      </thead>
      <tbody>
        <tr><td>23KB1A3096</td><td>&nbsp;&nbsp;shaik haneefa&nbsp;&nbsp;</td><td>95</td><td>AI_DS</td></tr>
        <tr><td>23KB1A3097</td><td>SAI KRISHNA</td><td class="text-amber">null</td><td>AI_DS</td></tr>
        <tr><td>23KB1A3098</td><td>&nbsp;&nbsp;sneha m.</td><td>72</td><td>IT</td></tr>
      </tbody>
    `;
  } else if (step === 'trim') {
    grid.innerHTML = `
      <thead>
        <tr>
          <th>Roll No</th>
          <th>Student Name (Trimmed &amp; Proper)</th>
          <th>Lab Score</th>
          <th>Dept Code</th>
        </tr>
      </thead>
      <tbody>
        <tr><td>23KB1A3096</td><td><strong>Shaik Haneefa</strong></td><td>95</td><td>AI_DS</td></tr>
        <tr><td>23KB1A3097</td><td><strong>Sai Krishna</strong></td><td class="text-amber">null</td><td>AI_DS</td></tr>
        <tr><td>23KB1A3098</td><td><strong>Sneha M.</strong></td><td>72</td><td>IT</td></tr>
      </tbody>
    `;
  } else if (step === 'nulls') {
    grid.innerHTML = `
      <thead>
        <tr>
          <th>Roll No</th>
          <th>Student Name</th>
          <th>Lab Score (Imputed)</th>
          <th>Dept Code</th>
        </tr>
      </thead>
      <tbody>
        <tr><td>23KB1A3096</td><td>Shaik Haneefa</td><td>95</td><td>AI_DS</td></tr>
        <tr><td>23KB1A3097</td><td>Sai Krishna</td><td class="text-emerald"><strong>80 (Dept Avg)</strong></td><td>AI_DS</td></tr>
        <tr><td>23KB1A3098</td><td>Sneha M.</td><td>72</td><td>IT</td></tr>
      </tbody>
    `;
  } else if (step === 'lookup') {
    grid.innerHTML = `
      <thead>
        <tr>
          <th>Roll No</th>
          <th>Student Name</th>
          <th>Score</th>
          <th>Department (XLOOKUP)</th>
          <th>Status</th>
        </tr>
      </thead>
      <tbody>
        <tr><td>23KB1A3096</td><td>Shaik Haneefa</td><td>95</td><td>Artificial Intelligence &amp; Data Science</td><td class="text-emerald">Validated</td></tr>
        <tr><td>23KB1A3097</td><td>Sai Krishna</td><td>80</td><td>Artificial Intelligence &amp; Data Science</td><td class="text-emerald">Validated</td></tr>
        <tr><td>23KB1A3098</td><td>Sneha M.</td><td>72</td><td>Information Technology</td><td class="text-emerald">Validated</td></tr>
      </tbody>
    `;
  }
}

// Day 3 Pivot Table & Chart
function updatePivotChart(groupBy) {
  const canvas = document.getElementById("pivotChartCanvas");
  if (!canvas) return;
  const ctx = canvas.getContext("2d");

  if (chartInstanceDay3) {
    chartInstanceDay3.destroy();
  }

  let labels = [];
  let dataVals = [];
  let bgColors = [];
  const insight = document.getElementById("pivot-ai-insight");

  if (groupBy === 'Department') {
    labels = ["AI & DS", "Information Tech", "Computer Science", "ECE"];
    dataVals = [89.4, 82.1, 84.6, 78.5];
    bgColors = ["#3b82f6", "#8b5cf6", "#10b981", "#f59e0b"];
    if (insight) insight.innerHTML = `AI Insight: <strong>AI &amp; DS Department</strong> accounts for the highest average score (89.4%) with a 14% higher distinction rate.`;
  } else if (groupBy === 'Year') {
    labels = ["1st Year", "2nd Year", "3rd Year", "4th Year (Cohort)"];
    dataVals = [76.2, 80.5, 84.1, 91.8];
    bgColors = ["#64748b", "#06b6d4", "#8b5cf6", "#3b82f6"];
    if (insight) insight.innerHTML = `AI Insight: <strong>4th Year B.Tech Cohort</strong> achieved the highest analytical mastery score (91.8%).`;
  } else {
    labels = ["Distinction (>85%)", "First Class (75-84%)", "Pass (60-74%)", "Remedial (<60%)"];
    dataVals = [58, 32, 8, 2];
    bgColors = ["#10b981", "#3b82f6", "#f59e0b", "#f43f5e"];
    if (insight) insight.innerHTML = `AI Insight: <strong>90% of students</strong> achieved First Class or Distinction in the workshop evaluation.`;
  }

  chartInstanceDay3 = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [{
        label: 'Average Performance Metric',
        data: dataVals,
        backgroundColor: bgColors,
        borderRadius: 6
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false }
      },
      scales: {
        y: { beginAtZero: true, grid: { color: 'rgba(255,255,255,0.08)' } },
        x: { grid: { display: false } }
      }
    }
  });
}

// Day 4 KPI Trend Chart
function renderKpiChart() {
  const canvas = document.getElementById("kpiTrendChart");
  if (!canvas) return;
  const ctx = canvas.getContext("2d");

  if (chartInstanceDay4) {
    chartInstanceDay4.destroy();
  }

  chartInstanceDay4 = new Chart(ctx, {
    type: 'line',
    data: {
      labels: ["Day 1: Basics", "Day 2: Cleaning", "Day 3: Pivots", "Day 4: Dashboards", "Day 5: Scripts"],
      datasets: [
        {
          label: 'Analytics Competency Index',
          data: [65, 78, 88, 95, 99],
          borderColor: '#3b82f6',
          backgroundColor: 'rgba(59, 130, 246, 0.15)',
          fill: true,
          tension: 0.4,
          pointRadius: 5,
          pointBackgroundColor: '#3b82f6'
        },
        {
          label: 'Task Completion Speed (min)',
          data: [45, 38, 28, 20, 12],
          borderColor: '#10b981',
          borderDash: [5, 5],
          fill: false,
          tension: 0.4,
          pointRadius: 4,
          pointBackgroundColor: '#10b981'
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'top', labels: { boxWidth: 12 } }
      },
      scales: {
        y: { beginAtZero: true, grid: { color: 'rgba(255,255,255,0.08)' } },
        x: { grid: { display: false } }
      }
    }
  });
}

// Day 5 Office Script Runner
function runOfficeScriptSimulation() {
  const btn = document.getElementById("btn-run-script");
  const status = document.getElementById("script-status-text");
  const consoleOutput = document.getElementById("script-console-output");
  const log = document.getElementById("script-log-text");

  if (btn) btn.disabled = true;
  if (status) status.innerHTML = `<i class="fa-solid fa-spinner fa-spin text-blue"></i> Running Office Script automation...`;

  setTimeout(() => {
    if (btn) btn.disabled = false;
    if (status) status.innerHTML = `<span class="text-emerald"><i class="fa-solid fa-check"></i> Executed in 184ms</span>`;
    if (consoleOutput) consoleOutput.style.display = "block";
    if (log) {
      log.innerHTML = `
        [00:00:01] Connected to Excel Online workbook Session 'NBKRIST_AI_Excel.xlsx'<br/>
        [00:00:02] Applied formatting: Table headers styled with #1e3a8a, frozen top panes.<br/>
        [00:00:03] Calculated Grand Average = 89.65 across 5 Day modules.<br/>
        [00:00:04] Output: Summary Dashboard generated with 0 errors. Status: SUCCESS!
      `;
    }
    showToast("Office Script automation completed successfully!", "success");
  }, 900);
}

// Star Rating & Confidence
function renderStarRating(dayIndex, currentRating) {
  let starsHtml = "";
  for (let i = 1; i <= 5; i++) {
    const isFilled = i <= currentRating;
    starsHtml += `<i class="fa-solid fa-star star-icon ${isFilled ? 'text-amber' : 'text-muted'}" onclick="setRating(${dayIndex}, ${i})"></i>`;
  }
  return starsHtml;
}

function setRating(dayIndex, rating) {
  appData.days[dayIndex].rating = rating;
  const starBox = document.getElementById(`star-box-${dayIndex}`);
  if (starBox) {
    starBox.innerHTML = renderStarRating(dayIndex, rating);
  }
  saveDataToStorage();
}

function setConfidence(dayIndex, level) {
  appData.days[dayIndex].confidence = level;
  renderStudioTab(currentActiveDayTab);
  saveDataToStorage();
}

function updateDayField(dayIndex, fieldName, value) {
  appData.days[dayIndex][fieldName] = value.trim();
  saveDataToStorage();
}

function updateReflection(dayIndex, refIndex, field, value) {
  if (appData.days[dayIndex].reflections[refIndex]) {
    appData.days[dayIndex].reflections[refIndex][field] = value.trim();
    saveDataToStorage();
  }
}

// ==========================================================================
// 6. OFFICIAL INSTITUTIONAL DOCUMENT VIEW (PRINT / PDF GENERATOR)
// ==========================================================================

function renderDocumentPages() {
  const container = document.getElementById("printable-pages-container");
  if (!container) return;

  const p = appData.profile;

  // Build complete 12 pages with outer border frames on EVERY single page
  let html = `
    <!-- ========================================== -->
    <!-- PAGE 1: OFFICIAL CERTIFICATE IMAGE (EXACT) -->
    <!-- ========================================== -->
    <div class="printable-a4-page cert-image-page">
      <div class="cert-image-container">
        <img src="assets/certificate.png" alt="N.B.K.R. Institute Certificate" class="cert-full-image" />
      </div>
    </div>

    <!-- ========================================== -->
    <!-- PAGE 2: INDEX TABLE                         -->
    <!-- ========================================== -->
    <div class="printable-a4-page">
      <div class="page-border-frame">
        <div class="index-page-content">
          <div class="index-main-heading">INDEX</div>
          
          <table class="index-table">
            <thead>
              <tr>
                <th class="col-date">Date</th>
                <th class="col-report">Day-wise Report</th>
                <th class="col-page">Page No.</th>
                <th class="col-sig">Signature</th>
              </tr>
            </thead>
            <tbody>
              ${appData.days.map((d, i) => {
                const dateVal = p.dates[`day${d.dayNumber}`] || "____ / ____ / 2026";
                const startPage = 3 + (i * 2);
                const endPage = startPage + 1;
                return `
                  <tr>
                    <td class="col-date">${dateVal}</td>
                    <td class="col-report">
                      <strong>Day ${d.dayNumber}</strong> – ${escapeHtml(d.title)}
                    </td>
                    <td class="col-page">${startPage} – ${endPage}</td>
                    <td class="col-sig" style="color: #475569; font-style: italic; font-size: 9pt;">Verified</td>
                  </tr>
                `;
              }).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `;

  // Append Daily Student Reports (Days 1 to 5, exactly 2 A4 pages each, every page wrapped in page-border-frame)
  appData.days.forEach(d => {
    const dateVal = p.dates[`day${d.dayNumber}`] || "____ / ____ / 2026";
    const ref1 = d.reflections[0] || { q: "", a: "" };
    const refRest = d.reflections.slice(1);

    html += `
      <!-- ========================================== -->
      <!-- DAY ${d.dayNumber} - SHEET 1               -->
      <!-- ========================================== -->
      <div class="printable-a4-page">
        <div class="page-border-frame">
          <div class="report-doc-header">
            <div class="report-inst-name">N.B.K.R. INSTITUTE OF SCIENCE AND TECHNOLOGY::</div>
            <div class="report-inst-city">VIDYANAGAR</div>
            <div class="report-inst-dept">(AUTONOMOUS) &bull; Department of IT and AI&amp;DS;</div>
          </div>

          <div class="report-workshop-banner">
            <div class="report-workshop-title">AI-POWERED MICROSOFT EXCEL WORKSHOP</div>
            <div class="report-workshop-sub">From Basics to Advanced Analytics — Daily Student Report</div>
          </div>

          <div class="report-day-headline">
            DAY ${d.dayNumber} ${escapeHtml(d.title)}
          </div>

          <!-- 1. Student & Session Profile Table -->
          <div class="doc-section-title">1. Student &amp; Session Profile</div>
          <table class="report-profile-table">
            <tr>
              <td class="field-title">Student Name</td>
              <td class="field-content">${escapeHtml(p.studentName)}</td>
            </tr>
            <tr>
              <td class="field-title">Roll No. / Student ID</td>
              <td class="field-content">${escapeHtml(p.rollNo)}</td>
            </tr>
            <tr>
              <td class="field-title">Branch / Year</td>
              <td class="field-content">${escapeHtml(p.branch)}</td>
            </tr>
            <tr>
              <td class="field-title">College / Institution</td>
              <td class="field-content">${escapeHtml(p.college)}</td>
            </tr>
            <tr>
              <td class="field-title">Workshop Day</td>
              <td class="field-content">Day ${d.dayNumber}</td>
            </tr>
            <tr>
              <td class="field-title">Date of Session</td>
              <td class="field-content">${dateVal}</td>
            </tr>
          </table>

          <!-- 2. Daily Technical Learning Logs -->
          <div class="doc-section-title">2. Daily Technical Learning Logs</div>
          <div class="doc-text-block"><strong>Focus:</strong> ${escapeHtml(d.focus)}</div>
          <div class="doc-text-block"><strong>Session 1 Topic:</strong> ${escapeHtml(d.session1)}</div>
          <div class="doc-text-block"><strong>Session 2 Topic:</strong> ${escapeHtml(d.session2)}</div>
          <div class="doc-text-block"><strong>AI Tools Used:</strong> ${escapeHtml(d.aiTools)}</div>
          <div class="doc-text-block"><strong>Excel Features Learned:</strong> ${escapeHtml(d.excelFeatures)}</div>
          <div class="doc-text-block"><strong>Hands-on Activity Completed:</strong> ${escapeHtml(d.handsOn)}</div>
          <div class="doc-text-block"><strong>Practical Task / Exercise:</strong> ${escapeHtml(d.practicalTask)}</div>
          <div class="doc-text-block"><strong>What I Learned Today:</strong> ${escapeHtml(d.whatILearned)}</div>

          <!-- 3. Difficulties & Problem Solving -->
          <div class="doc-section-title">3. Difficulties &amp; Problem Solving</div>
          <div class="doc-text-block"><strong>Difficulties Faced:</strong> ${escapeHtml(d.difficulties)}</div>
          <div class="doc-text-block"><strong>How I Solved the Difficulties:</strong> ${escapeHtml(d.howSolved)}</div>

          <!-- 4. Evaluation & Feedback -->
          <div class="doc-section-title">4. Evaluation &amp; Feedback</div>
          <div class="doc-text-block">
            <strong>My Confidence Level:</strong> 
            [${d.confidence === 'Beginner' ? '&#10003;' : '&nbsp;&nbsp;'}] Beginner 
            [${d.confidence === 'Intermediate' ? '&#10003;' : '&nbsp;&nbsp;'}] Intermediate 
            [${d.confidence === 'Advanced' ? '&#10003;' : '&nbsp;&nbsp;'}] Advanced
          </div>
          <div class="doc-text-block">
            <strong>Workshop Rating:</strong> 
            [${d.rating === 1 ? '&#10003;' : '&nbsp;&nbsp;'}] 1 
            [${d.rating === 2 ? '&#10003;' : '&nbsp;&nbsp;'}] 2 
            [${d.rating === 3 ? '&#10003;' : '&nbsp;&nbsp;'}] 3 
            [${d.rating === 4 ? '&#10003;' : '&nbsp;&nbsp;'}] 4 
            [${d.rating === 5 ? '&#10003;' : '&nbsp;&nbsp;'}] 5 (Highest)
          </div>
          <div class="doc-text-block"><strong>Suggestions / Feedback:</strong> ${escapeHtml(d.suggestions)}</div>

          <!-- 5. Short Daily Reflection Questions (Q1 on Page 1) -->
          <div class="doc-section-title">5. Short Daily Reflection Questions</div>
          <div class="doc-reflection-q">1. ${escapeHtml(ref1.q)}</div>
          <div class="doc-reflection-a"><strong>Ans:</strong> ${escapeHtml(ref1.a)}</div>
        </div>
      </div>

      <!-- ========================================== -->
      <!-- DAY ${d.dayNumber} - SHEET 2               -->
      <!-- ========================================== -->
      <div class="printable-a4-page">
        <div class="page-border-frame">
          <div style="padding-top: 5mm;">
            ${refRest.map((rf, rIdx) => `
              <div class="doc-reflection-q" style="margin-top: 14px;">${rIdx + 2}. ${escapeHtml(rf.q)}</div>
              <div class="doc-reflection-a" style="margin-top: 6px; line-height: 1.6;"><strong>Ans:</strong> ${escapeHtml(rf.a)}</div>
            `).join('')}
          </div>
        </div>
      </div>
    `;
  });

  container.innerHTML = html;
}

// ==========================================================================
// 7. PROFILE & JSON BACKUP MODAL MANAGEMENT
// ==========================================================================

function openStudentModal() {
  const p = appData.profile;
  document.getElementById("input-student-name").value = p.studentName || "";
  document.getElementById("input-roll-no").value = p.rollNo || "";
  document.getElementById("input-branch").value = p.branch || "";
  document.getElementById("input-college").value = p.college || "";
  document.getElementById("input-lab-name").value = p.labName || "";
  document.getElementById("input-acad-year").value = p.acadYear || "";
  document.getElementById("input-staff-incharge").value = p.staffInCharge || "";
  document.getElementById("input-examiner").value = p.examiner || "";

  document.getElementById("input-date-day1").value = p.dates.day1 || "";
  document.getElementById("input-date-day2").value = p.dates.day2 || "";
  document.getElementById("input-date-day3").value = p.dates.day3 || "";
  document.getElementById("input-date-day4").value = p.dates.day4 || "";
  document.getElementById("input-date-day5").value = p.dates.day5 || "";

  openModal('modal-profile');
}

function saveStudentProfile() {
  const p = appData.profile;
  p.studentName = document.getElementById("input-student-name").value.trim();
  p.rollNo = document.getElementById("input-roll-no").value.trim();
  p.branch = document.getElementById("input-branch").value.trim();
  p.college = document.getElementById("input-college").value.trim();
  p.labName = document.getElementById("input-lab-name").value.trim();
  p.acadYear = document.getElementById("input-acad-year").value.trim();
  p.staffInCharge = document.getElementById("input-staff-incharge").value.trim();
  p.examiner = document.getElementById("input-examiner").value.trim();

  p.dates.day1 = document.getElementById("input-date-day1").value.trim();
  p.dates.day2 = document.getElementById("input-date-day2").value.trim();
  p.dates.day3 = document.getElementById("input-date-day3").value.trim();
  p.dates.day4 = document.getElementById("input-date-day4").value.trim();
  p.dates.day5 = document.getElementById("input-date-day5").value.trim();

  closeModal('modal-profile');
  saveDataToStorage();
  renderStudioTab(currentActiveDayTab);
}

function openDataModal() {
  document.getElementById("json-data-area").value = JSON.stringify(appData, null, 2);
  openModal('modal-data');
}

function copyJsonToClipboard() {
  const area = document.getElementById("json-data-area");
  area.select();
  navigator.clipboard.writeText(area.value).then(() => {
    showToast("JSON copied to clipboard!", "info");
  });
}

function downloadJsonFile() {
  const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(appData, null, 2));
  const downloadAnchor = document.createElement('a');
  downloadAnchor.setAttribute("href", dataStr);
  downloadAnchor.setAttribute("download", `NBKRIST_Excel_Report_${appData.profile.rollNo || '23KB1A3096'}.json`);
  document.body.appendChild(downloadAnchor);
  downloadAnchor.click();
  downloadAnchor.remove();
  showToast("JSON file downloaded", "success");
}

function uploadJsonFile(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const parsed = JSON.parse(e.target.result);
      if (parsed.profile && parsed.days) {
        document.getElementById("json-data-area").value = JSON.stringify(parsed, null, 2);
        showToast("File loaded into editor. Click 'Apply JSON' to save.", "info");
      } else {
        alert("Invalid format: JSON must contain 'profile' and 'days' arrays.");
      }
    } catch (err) {
      alert("Error reading JSON file: " + err.message);
    }
  };
  reader.readAsText(file);
}

function applyJsonData() {
  try {
    const raw = document.getElementById("json-data-area").value;
    const parsed = JSON.parse(raw);
    if (!parsed.profile || !parsed.days) {
      throw new Error("Missing profile or days object");
    }
    appData = parsed;
    saveDataToStorage();
    closeModal('modal-data');
    renderStudioTab(currentActiveDayTab);
    showToast("Custom JSON applied successfully!", "success");
  } catch (e) {
    alert("Invalid JSON format. Please verify the structure and syntax.");
  }
}

function confirmResetDefaults() {
  if (confirm("Are you sure you want to reset all data back to original default document values? Any custom edits will be reverted.")) {
    appData = JSON.parse(JSON.stringify(DEFAULT_DATA));
    saveDataToStorage();
    closeModal('modal-data');
    renderStudioTab(currentActiveDayTab);
    showToast("Reset to default scanned document state", "info");
  }
}

function openModal(id) {
  const modal = document.getElementById(id);
  if (modal) modal.classList.add("open");
}

function closeModal(id) {
  const modal = document.getElementById(id);
  if (modal) modal.classList.remove("open");
}

// Utility Helpers
function escapeHtml(str) {
  if (!str) return "";
  return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function showToast(message, type = "info") {
  const container = document.getElementById("toast-container");
  if (!container) return;

  const toast = document.createElement("div");
  toast.className = `toast-msg toast-${type}`;
  toast.innerHTML = `<i class="fa-solid fa-circle-info"></i> <span>${message}</span>`;
  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = "0";
    toast.style.transform = "translateX(50px)";
    setTimeout(() => toast.remove(), 300);
  }, 3500);
}

// Bootstrap on DOM ready
document.addEventListener("DOMContentLoaded", () => {
  initApp();
});
