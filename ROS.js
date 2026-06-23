// Escape HTML special characters
function escapeHTML(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

// Pillar-wise OMBP data
const ombpData = {
    HCM: [
        { name: "Attract Talent to Onboard Workers", apiName: "uv7679dw" },
        { name: "Employee Onboarding to Retirement", apiName: "mvdj7la4" },
        { name: "Skills Development to Career Growth", apiName: "l6m686bx" },
        { name: "Workforce Goals to Performance", apiName: "nm4eeqe6" },
        { name: "Skills Insight to Workforce Agility", apiName: "5a49scxi" },
        { name: "Employee Experience to Business Outcome", apiName: "mrbp9n9q" },
        { name: "Absence Policy to Compliance", apiName: "nbxz7w7t" },
        { name: "Workforce Schedule to Payment", apiName: "rqv0bci6" },
        { name: "Incident Report to Workplace Safety", apiName: "shkqty9q" },
        { name: "Workforce Issue to Resolution", apiName: "p9v67y3j" },
        { name: "Employee Insight to Work-Life Alignment", apiName: "kaojtfcl" },
        { name: "Compensation Plans to Reward", apiName: "aq20habd" },
        { name: "Benefits Enrollment to Coverage", apiName: "bx5i86mc" },
        { name: "Payroll Input to Payment", apiName: "h0ub5ct4" }
    ],

    ERP: [
        { name: "Asset Acquisition to Retirement", apiName: "7h1l4h7l" },
        { name: "Bank Transaction to Cash Position", apiName: "79ps2rlx" },
        { name: "Capital Project to Asset", apiName: "9o8lcgt2" },
        { name: "Customer Contract to Revenue", apiName: "5pa7aqc8" },
        { name: "Customer Invoice to Receipt", apiName: "fo4u7kvy" },
        { name: "Customer Statements to Collections", apiName: "6f4xp4vb" },
        { name: "Expense Report to Reimbursement", apiName: "xr10sxvp" },
        { name: "Grant Award Funding to Closeout", apiName: "rt5wcjhq" },
        { name: "Internal Controls Automation to Certification", apiName: "ndqgnhqb" },
        { name: "Period Close to Financial Reports", apiName: "dc4vxvkg" },
        { name: "Project Contract Billing to Revenue Recognition", apiName: "d63le6l7" },
        { name: "Project Costs to Accounting", apiName: "agyyt66b" },
        { name: "Project Plan to Delivery", apiName: "uxli273u" },
        { name: "Resource Analysis to Utilization", apiName: "wem13nm9" },
        { name: "Secure Role Design to User Access Control", apiName: "u9npziqv" },
        { name: "Separation of Duties (SoD) to User Access Certification", apiName: "qux140kg" },
        { name: "Supplier Invoice to Payment", apiName: "vg643am4" },
        { name: "Task Detail to Plan Adjustment", apiName: "605kkrry" }
    ],

    SCM: [
        { name: "Contract Creation to Procurement Compliance", apiName: "ru9v7pgn" },
        { name: "Fulfilment Orchestration to Invoice", apiName: "et3ymxfp" },
        { name: "Insight to Sourcing", apiName: "nweynozw" },
        { name: "Material Request to Delivery", apiName: "ggb271ed" },
        { name: "Multi-channel Order to Promise", apiName: "c043wa7l" },
        { name: "Order to Drop Shipment", apiName: "2zl5w2pc" },
        { name: "Quote to Revenue", apiName: "apn8zlam" },
        { name: "Requisition to Payment", apiName: "g4r04kfu" },
        { name: "Supplier Return to Settlement", apiName: "i064wyml" },
        { name: "Supplier Registration to Spend Authorization", apiName: "4gwg58k2" },
        { name: "Plan To Replenish", apiName: "ahpbkquw" },
        { name: "Demand to Management", apiName: "bxn5og62" },
        { name: "Demand Forecast to Supply Plan", apiName: "jhp8yqj0" },

        // No OMBP Description available for this item.
        // This will skip the OMBP and Persona Description link.
        { name: "Production Order to Costing", apiName: "" },

        { name: "Secure Role Design to User Access Control", apiName: "h5jc9q2k" },
        { name: "Separation of Duties (SoD) to User Access Certification", apiName: "9jcnbci7" }
    ],

    CX: [],
    EPM: [],
    WMS: [],
    OTM: []
};

// Load OMBP dropdown values based on selected Pillar
document.addEventListener("DOMContentLoaded", function () {
    const pillarSelect = document.getElementById("pillarName");
    const ombpSelect = document.getElementById("ombpName");
    const apiNameInput = document.getElementById("ombpDescriptionApiName");
    const codeForm = document.getElementById("codeForm");

    pillarSelect.addEventListener("change", function () {
        const selectedPillar = pillarSelect.value;

        ombpSelect.innerHTML = `<option value="">Select OMBP Name</option>`;
        apiNameInput.value = "";

        if (!selectedPillar || !ombpData[selectedPillar]) {
            return;
        }

        if (ombpData[selectedPillar].length === 0) {
            const option = document.createElement("option");
            option.value = "";
            option.textContent = "No OMBP values available";
            ombpSelect.appendChild(option);
            return;
        }

        ombpData[selectedPillar].forEach(function (item) {
            const option = document.createElement("option");
            option.value = item.name;
            option.textContent = item.name;
            option.setAttribute("data-api-name", item.apiName);
            ombpSelect.appendChild(option);
        });
    });

    ombpSelect.addEventListener("change", function () {
        const selectedOption = ombpSelect.options[ombpSelect.selectedIndex];
        apiNameInput.value = selectedOption.getAttribute("data-api-name") || "";
    });

    codeForm.addEventListener("reset", function () {
        setTimeout(function () {
            ombpSelect.innerHTML = `<option value="">Select OMBP Name</option>`;
            apiNameInput.value = "";
            resetOutput();
        }, 0);
    });
});

// Generate HTML from uploaded Excel
function generateHTMLFromExcel() {
    const appId = document.getElementById("appId").value.trim();
    const pillarName = document.getElementById("pillarName").value.trim();
    const ombpName = document.getElementById("ombpName").value.trim();
    const ombpDescriptionApiName = document.getElementById("ombpDescriptionApiName").value.trim();
    const fileInput = document.getElementById("excelFile");

    if (!appId) {
        alert("❗ Please enter App Id.");
        return;
    }

    if (!pillarName) {
        alert("❗ Please select Pillar Name.");
        return;
    }

    if (!ombpName) {
        alert("❗ Please select OMBP Name.");
        return;
    }

    // Skip OMBP Description only when that selected OMBP has no description API name.
    const skipOmbpDescription = !ombpDescriptionApiName;

    if (!fileInput.files.length) {
        alert("❗ Please upload an Excel file.");
        return;
    }

    const file = fileInput.files[0];
    const reader = new FileReader();

    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

        if (jsonData.length < 2) {
            alert("❗ Excel file must have at least one data row starting from row 2.");
            return;
        }

        const groupedData = {};

        // Excel columns:
        // A = Role
        // B = Guide name
        // C = ApiName
        for (let i = 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            const role = row[0];
            const guideName = row[1];
            const apiName = row[2];

            if (!role || !guideName || !apiName) {
                continue;
            }

            const cleanRole = String(role).trim();
            const cleanGuideName = String(guideName).trim();
            const cleanApiName = String(apiName).trim();

            if (!groupedData[cleanRole]) {
                groupedData[cleanRole] = [];
            }

            groupedData[cleanRole].push({
                guideName: cleanGuideName,
                apiName: cleanApiName
            });
        }

        if (Object.keys(groupedData).length === 0) {
            alert("❗ No valid rows found. Please check Role, Guide name, and ApiName columns.");
            return;
        }

        let outputHTML = "";

outputHTML += `<h4 style="margin-bottom:10px !important;">
  <a
    ${!skipOmbpDescription
      ? `data-iridize-nextscenario="{&quot;nextScenario&quot;:&quot;${escapeHTML(ombpDescriptionApiName)}&quot;,&quot;dontClose&quot;:true,&quot;markClosed&quot;:false}"
         data-iridize-role="nextScenarioBt"
         href="javascript:void(0)"`
      : ''
    }
    style="
      display:flex;
      align-items:center;
      justify-content:center;
      padding:12px 14px;
      color:#000;
      text-decoration:none;
      border-radius:4px;
      font-size:16px;
      position:relative;
      font-weight:bold;
      ${!skipOmbpDescription ? 'cursor:pointer;' : 'cursor:default;'}
    "
  >
    ${escapeHTML(pillarName)} - ${escapeHTML(ombpName)}
  </a>
</h4>\n`;

        Object.keys(groupedData).forEach(function (role) {
            outputHTML += `<!-- ${escapeHTML(role)} Dropdown -->\n\n`;

            outputHTML += `<details style="
    border-radius:4px;
    cursor: alias;
    border:1px solid #D8D8D8;
    overflow:hidden;
    background-color:white;
  "><summary class="summary-arrow" style="
      cursor:pointer;
      padding:20px 18px;
      font-size: 16px;
      font-weight: bold;
      line-height:1.4;
      list-style:none;
      display:flex;
      align-items:center;
      position:relative;
    "> <span style="font-weight: bold; black"> ${escapeHTML(role)} </span> <span class="chevron" style="margin-left:auto">❯</span> </summary>

<div class="scroll-content" style="background-color:#f5f4f4; border-radius:20px; border:10px solid white; display:flex; flex-wrap:wrap; justify-content:space-between; max-height:150px; overflow-x:hidden; overflow-y:auto; padding:10px; text-align:left">\n`;

            groupedData[role].forEach(function (item) {
                const simulationHref = `https://guidedlearning.oracle.com/player/latest/api/scenario/simulation/try_it/${encodeURIComponent(appId)}/${encodeURIComponent(item.apiName)}/lang/--/?draft=dev&windowMode=unpin`;
                const stepGuideHref = `https://guidedlearning.oracle.com/player/latest/api/scenario/export/v2/${encodeURIComponent(appId)}/${encodeURIComponent(item.apiName)}/lang/--/?draft=dev&windowMode=unpin`;
                const videoHref = `https://guidedlearning.oracle.com/player/latest/api/scenario/simulation/see_it/${encodeURIComponent(appId)}/${encodeURIComponent(item.apiName)}/lang/--/?draft=dev&windowMode=unpin`;

                const menuId = `menu-${role}-${item.apiName}`.replace(/[^a-zA-Z0-9_-]/g, '_');

outputHTML += `<div style="margin-bottom:10px; margin-top:10px; width:100%">
  <div style="align-items:center; display:flex; gap:10px;">
    <a
      style="flex:1; display:block; padding:12px; font-size:14px; border-radius:10px; text-decoration:none; color:#000; background:#ffffff; font-weight:450; cursor:default;"
    >
      ${escapeHTML(item.guideName)}
    </a>

    <details class="action-details" style="position:relative; flex-shrink:0;">
      <summary
        style="
          list-style:none;
          width:34px;
          height:34px;
          border:none;
          background:#fff;
          border-radius:8px;
          cursor:pointer;
          display:flex;
          align-items:center;
          justify-content:center;
          font-size:22px;
          color:#d10000;
          font-weight:bold;
          padding:0;
        "
        title="Actions"
      >⋮</summary>

      <div
        style="
          position:absolute;
          right:0;
          top:40px;
          min-width:170px;
          background:#fff;
          border:1px solid #ddd;
          border-radius:10px;
          box-shadow:0 8px 24px rgba(0,0,0,.12);
          z-index:9999;
          overflow:hidden;
        "
      >
        <a href="${simulationHref}" target="_blank" style="display:block; padding:10px 12px; text-decoration:none; color:#000;">Simulation</a>
        <a href="${stepGuideHref}" target="_blank" style="display:block; padding:10px 12px; text-decoration:none; color:#000;">Step Guide</a>
        <a href="${videoHref}" target="_blank" style="display:block; padding:10px 12px; text-decoration:none; color:#000;">Video</a>
      </div>
    </details>
  </div>
</div>\n`;
        });

        outputHTML += `<style type="text/css">
summary.summary-arrow {
    cursor: pointer;
    padding: 12px 18px;
    font-size: 12px;
    font-weight: 600;
    line-height: 1.4;
    list-style: none;
    position: relative;
  }

  summary.summary-arrow::-webkit-details-marker {
    display: none;
  }

  summary.summary-arrow .chevron {
    position: absolute;
    right: 14px;
    top: 50%;
    transform: translateY(-50%) rotate(90deg);
    transition: transform 0.3s ease;
    color: #6b7280;
    font-size: 12px;
    font-weight: 200;
    display: inline-block;
  }

  details[open] summary.summary-arrow .chevron {
    transform: translateY(-50%) rotate(270deg);
  }

  .scroll-content {
    scrollbar-width: thin;
  }
  @keyframes blinkRed {
  0%, 100% { color: #d10000; opacity: 1; }
  50% { color: #ff4d4d; opacity: 0.35; }
}

.action-details > summary {
  animation: blinkRed 1s infinite;
}

.action-details > summary::-webkit-details-marker {
  display: none;
}

.action-details[open] > summary {
  animation: none;
  color: #d10000;
}

.action-details a:hover {
  background: #f5f5f5;
}
</style>`;

        document.getElementById("output").value = outputHTML;
    };

    reader.readAsArrayBuffer(file);
}

// Copy generated HTML to clipboard
document.getElementById("copyButton").addEventListener("click", function () {
    const outputText = document.getElementById("output");

    if (!outputText.value.trim()) {
        alert("❗ No generated code to copy.");
        return;
    }

    outputText.select();
    outputText.setSelectionRange(0, 99999);

    navigator.clipboard.writeText(outputText.value)
        .then(() => alert("✅ HTML code copied to clipboard!"))
        .catch((err) => alert("❌ Failed to copy: " + err));
});

// Download Excel template
function downloadExcelTemplate() {
    const wb = XLSX.utils.book_new();

    const ws_data = [
        ["Role", "Guide name", "ApiName"]
    ];

    const ws = XLSX.utils.aoa_to_sheet(ws_data);
    XLSX.utils.book_append_sheet(wb, ws, "Template");
    XLSX.writeFile(wb, "OMBP_Persona_Template.xlsx");
}

// Download Runbook Redwood Theme
function downloadRunbookTheme() {
    const link = document.createElement("a");
    link.href = "Runbook Redwood Theme.html";
    link.download = "Runbook Redwood Theme.html";
    link.click();
}

// Reset output
function resetOutput() {
    document.getElementById("output").value = "";
}
