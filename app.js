let pyodide = null;
let inputFileData = null;
let inputFileName = null;
let currentTab = 'up-down'; // Default tab

// Console logging to UI
function logToConsole(message) {
    const consoleLog = document.getElementById('consoleLog');
    const consoleContainer = document.getElementById('consoleContainer');
    consoleContainer.style.display = 'block';
    consoleLog.textContent += message + '\n';
    consoleLog.scrollTop = consoleLog.scrollHeight;
    console.log(message);
}

// Tab switching
document.addEventListener('DOMContentLoaded', () => {
    const tabBtns = document.querySelectorAll('.tab-btn');
    const tabContents = document.querySelectorAll('.tab-content');
    
    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            // Remove active class from all buttons and contents
            tabBtns.forEach(b => b.classList.remove('active'));
            tabContents.forEach(c => c.classList.remove('active'));
            
            // Add active class to clicked button and corresponding content
            btn.classList.add('active');
            const tabId = btn.getAttribute('data-tab');
            document.getElementById(tabId).classList.add('active');
            
            // Update current tab
            currentTab = tabId;
            logToConsole(`Switched to ${tabId} version`);
        });
    });
});

// Initialize Pyodide when page loads
async function initPyodide() {
    showStatus('Initializing Python environment... This may take a minute on first load.', 'loading');
    logToConsole('Loading Pyodide...');
    try {
        pyodide = await loadPyodide();
        logToConsole('Pyodide loaded');
        
        // Load standard packages
        logToConsole('Loading packages: pandas, numpy, micropip...');
        await pyodide.loadPackage(['pandas', 'numpy', 'micropip']);
        logToConsole('Packages loaded');
        
        // Use micropip to install openpyxl from PyPI
        logToConsole('Installing openpyxl...');
        const micropip = pyodide.pyimport('micropip');
        await micropip.install('openpyxl');
        logToConsole('openpyxl installed');
        
        showStatus('Python environment ready! You can now upload your Excel file.', 'success');
        logToConsole('=== READY TO PROCESS FILES ===');
        setTimeout(() => hideStatus(), 3000);
    } catch (error) {
        showStatus('Failed to initialize Python environment: ' + error.message, 'error');
        logToConsole('ERROR: ' + error.message);
        console.error('Initialization error:', error);
    }
}

// File input handling
document.getElementById('fileInput').addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (file) {
        handleFileSelect(file);
    }
});

// Drag and drop handling
const fileLabel = document.getElementById('fileLabel');
fileLabel.addEventListener('dragover', (e) => {
    e.preventDefault();
    fileLabel.style.borderColor = '#667eea';
    fileLabel.style.background = '#f0f0ff';
});

fileLabel.addEventListener('dragleave', (e) => {
    e.preventDefault();
    fileLabel.style.borderColor = '#ccc';
    fileLabel.style.background = '#f5f5f5';
});

fileLabel.addEventListener('drop', (e) => {
    e.preventDefault();
    fileLabel.style.borderColor = '#ccc';
    fileLabel.style.background = '#f5f5f5';
    const file = e.dataTransfer.files[0];
    if (file) {
        handleFileSelect(file);
    }
});

function handleFileSelect(file) {
    if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
        showStatus('Please select an Excel file (.xlsx or .xls)', 'error');
        return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
        inputFileData = new Uint8Array(e.target.result);
        inputFileName = file.name;
        
        document.getElementById('fileLabel').classList.add('has-file');
        document.getElementById('fileLabel').innerHTML = `
            <div style="font-size: 40px; margin-bottom: 10px;">✓</div>
            <div><strong>${file.name}</strong></div>
            <div class="help-text">${(file.size / 1024).toFixed(2)} KB</div>
        `;
        
        document.getElementById('fileInfo').innerHTML = `
            <strong>File loaded:</strong> ${file.name} (${(file.size / 1024).toFixed(2)} KB)
        `;
        
        document.getElementById('processBtn').disabled = false;
    };
    reader.readAsArrayBuffer(file);
}

// Function to get Python script based on selected tab
function getPythonScriptForTab(tab, sampleMapping, axisMapping) {
    logToConsole(`Loading ${tab} processing script...`);
    
    switch(tab) {
        case 'up-down':
            return getUpDownScript(sampleMapping, axisMapping);
        case 'a3':
            return getA3Script(sampleMapping, axisMapping);
        case 'classic':
            return getClassicScript(sampleMapping, axisMapping);
        default:
            logToConsole(`Unknown tab: ${tab}, defaulting to up-down`);
            return getUpDownScript(sampleMapping, axisMapping);
    }
}

function getUpDownScript(sampleMapping, axisMapping) {
    return `
import sys
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import load_workbook

print("Python script started...")
sys.stdout.flush()

def autofit_columns(workbook, sheet_names=None):
    if sheet_names is None:
        sheet_names = workbook.sheetnames
    
    for sheet_name in sheet_names:
        if sheet_name not in workbook.sheetnames:
            continue
            
        sheet = workbook[sheet_name]
        column_widths = {}
        
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is not None:
                    column_letter = cell.column_letter
                    cell_value = str(cell.value)
                    
                    if isinstance(cell.value, (int, float)):
                        display_width = len(cell_value) + 1
                    else:
                        display_width = len(cell_value)
                    
                    if column_letter not in column_widths:
                        column_widths[column_letter] = display_width
                    else:
                        column_widths[column_letter] = max(column_widths[column_letter], display_width)
        
        for column_letter, max_width in column_widths.items():
            adjusted_width = min(max(max_width + 2, 8), 50)
            sheet.column_dimensions[column_letter].width = adjusted_width

def process_excel_file(input_file, output_file, sample_mapping, axis_mapping):
    input_workbook = pd.ExcelFile(input_file)
    sheet_names = input_workbook.sheet_names

    all_vdiff_data = pd.DataFrame()
    all_sensitivity_data = pd.DataFrame()

    for sheet_name in sheet_names[2:]:
        if "precon" in sheet_name.lower():
            print(f"Skipping sheet: {sheet_name}")
            continue
            
        df = input_workbook.parse(sheet_name)

        columns_to_keep = [
            "Sample No", "Angle", "Axis", "B_stress[mT]", "Vdd[V]",
            "B_read[mT]", "B_set[mT]", "Temp_Gauss_probe[C]",
            "Vdiff_max[mV]", "Vdiff_min[mV]", "Vdiff_mean[mV]",
            "Vdiff_stdev[uV]"
        ]

        df = df.dropna(how='all', subset=columns_to_keep)
        vdiff_data = df[columns_to_keep].copy()

        columns_to_expand = ["Axis", "Angle", "B_stress[mT]", "Vdd[V]"]
        for col in columns_to_expand:
            vdiff_data[col] = np.repeat(vdiff_data[col].values, 5)[:len(vdiff_data)]

        sample_no_value = vdiff_data["Sample No"].iloc[0]
        vdiff_data["Sample No"] = vdiff_data["Sample No"].fillna(sample_no_value)
        vdiff_data.loc[vdiff_data["Axis"].isna(), "Sample No"] = np.nan

        vdiff_data["Off_DU[mV]"] = np.nan
        
        for i in range(0, len(vdiff_data), 5):
            if i + 4 < len(vdiff_data):
                try:
                    off_d = vdiff_data.loc[i+1, "Vdiff_mean[mV]"]
                    off_u = vdiff_data.loc[i+3, "Vdiff_mean[mV]"]
                    off_du = (off_d + off_u) / 2
                    vdiff_data.loc[i, "Off_DU[mV]"] = off_du
                except (IndexError, KeyError, TypeError) as e:
                    print(f"Error calculating offsets for group starting at row {i}: {e}")
                    continue

        vdiff_data["Off_drift_DU[mV]"] = np.nan
        
        for (sample_no, axis), group in vdiff_data.groupby(["Sample No", "Axis"]):
            group_indices = group.index.tolist()
            stress_10_data = group[group["B_stress[mT]"] == 10]
            off_du_ref = None
            
            if not stress_10_data.empty:
                off_du_values = stress_10_data["Off_DU[mV]"].dropna()
                if not off_du_values.empty:
                    off_du_ref = off_du_values.iloc[0]
            
            for i in range(0, len(group_indices), 5):
                if i < len(group_indices):
                    idx = group_indices[i]
                    if off_du_ref is not None and not pd.isna(vdiff_data.loc[idx, "Off_DU[mV]"]):
                        vdiff_data.loc[idx, "Off_drift_DU[mV]"] = vdiff_data.loc[idx, "Off_DU[mV]"] - off_du_ref

        sensitivity_data = vdiff_data.copy()
        columns_to_remove = ["Off_DU[mV]", "Off_drift_DU[mV]"]
        for col in columns_to_remove:
            if col in sensitivity_data.columns:
                sensitivity_data.drop(columns=[col], inplace=True)

        sensitivity_data["Sens_DU[mV/mT]"] = np.nan

        for i in range(0, len(sensitivity_data), 5):
            if i + 4 < len(sensitivity_data):
                group = sensitivity_data.iloc[i:i+5]
                
                try:
                    d_b_set = group.iloc[0:3]["B_set[mT]"].values
                    d_vdiff = group.iloc[0:3]["Vdiff_mean[mV]"].values
                    u_b_set = group.iloc[2:5]["B_set[mT]"].values
                    u_vdiff = group.iloc[2:5]["Vdiff_mean[mV]"].values
                    
                    sens_d, sens_u = np.nan, np.nan
                    
                    if not (np.isnan(d_vdiff).any() or np.isnan(d_b_set).any()) and len(set(d_b_set)) > 1:
                        sens_d, _ = np.polyfit(d_b_set, d_vdiff, 1)
                    
                    if not (np.isnan(u_vdiff).any() or np.isnan(u_b_set).any()) and len(set(u_b_set)) > 1:
                        sens_u, _ = np.polyfit(u_b_set, u_vdiff, 1)
                    
                    if not (np.isnan(sens_d) or np.isnan(sens_u)):
                        sens_du = (sens_d + sens_u) / 2
                    else:
                        sens_du = np.nan
                    
                    for j in range(5):
                        if i + j < len(sensitivity_data):
                            sensitivity_data.loc[i+j, "Sens_DU[mV/mT]"] = sens_du
                
                except Exception as e:
                    print(f"Error calculating sensitivities for group starting at row {i}: {e}")
                    continue

        sensitivity_data["Sens_drift_DU[%]"] = np.nan
        
        for (sample_no, axis), group in sensitivity_data.groupby(["Sample No", "Axis"]):
            group_indices = group.index.tolist()
            stress_10_data = group[group["B_stress[mT]"] == 10]
            sens_du_ref = None
            
            if not stress_10_data.empty:
                sens_du_values = stress_10_data["Sens_DU[mV/mT]"].dropna()
                if not sens_du_values.empty:
                    sens_du_ref = sens_du_values.iloc[0]
            
            for idx in group_indices:
                if sens_du_ref is not None and not pd.isna(sensitivity_data.loc[idx, "Sens_DU[mV/mT]"]) and sens_du_ref != 0:
                    drift = (sensitivity_data.loc[idx, "Sens_DU[mV/mT]"] - sens_du_ref) / sens_du_ref * 100
                    sensitivity_data.loc[idx, "Sens_drift_DU[%]"] = drift

        sens_drift_columns = ["Sens_drift_DU[%]"]
        for col in sens_drift_columns:
            if col in vdiff_data.columns:
                vdiff_data.drop(columns=[col], inplace=True)

        filtered_indices = []
        for i in range(0, len(sensitivity_data), 5):
            if i < len(sensitivity_data) and not pd.isna(sensitivity_data.iloc[i]["Sens_DU[mV/mT]"]):
                filtered_indices.append(i)
        
        sensitivity_data = sensitivity_data.iloc[filtered_indices].reset_index(drop=True)

        column_rename_map = {
            "Axis": "Plane",
            "B_stress[mT]": "Magnetic Field Stress[mT]",
            "Vdd[V]": "Sample consumption[V]",
            "B_read[mT]": "Magnetic Field Read[mT]",
            "B_set[mT]": "Magnetic Field Set[mT]"
        }
        vdiff_data.rename(columns=column_rename_map, inplace=True)
        sensitivity_data.rename(columns=column_rename_map, inplace=True)

        all_vdiff_data = pd.concat([all_vdiff_data, vdiff_data], ignore_index=True)
        all_sensitivity_data = pd.concat([all_sensitivity_data, sensitivity_data], ignore_index=True)

    all_vdiff_data.dropna(how="all", inplace=True)
    all_sensitivity_data.dropna(how="all", inplace=True)

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        all_vdiff_data.to_excel(writer, sheet_name="Vdiff_data", index=False)
        all_sensitivity_data.to_excel(writer, sheet_name="Sensitivity_data", index=False)

    wb = load_workbook(output_file)
    for sheet_name in ["Vdiff_data", "Sensitivity_data"]:
        sheet = wb[sheet_name]
        for row in range(2, sheet.max_row + 1):
            sample_no_cell = sheet[f"A{row}"]
            axis_cell = sheet[f"C{row}"]

            if sample_no_cell.value in sample_mapping:
                sample_no_cell.value = sample_mapping[sample_no_cell.value]

            if axis_cell.value in axis_mapping:
                axis_cell.value = axis_mapping[axis_cell.value]

    # Generate Summary sheet
    summary = []
    
    # Read Vdiff_data for summary
    vdiff_sheet = wb["Vdiff_data"]
    vdiff_df = pd.DataFrame(vdiff_sheet.values)
    vdiff_df.columns = vdiff_df.iloc[0]
    vdiff_df = vdiff_df[1:]
    vdiff_df.columns = vdiff_df.columns.str.strip()
    vdiff_df = vdiff_df.dropna(how="all")
    
    # Process Vdiff_data drift values
    for col in ["Magnetic Field Set[mT]", "Magnetic Field Stress[mT]", "Off_drift_DU[mV]"]:
        if col in vdiff_df.columns:
            vdiff_df[col] = pd.to_numeric(vdiff_df[col], errors='coerce')
    
    vdiff_47 = vdiff_df[vdiff_df["Magnetic Field Set[mT]"] == 47]
    for idx, row in vdiff_47.iterrows():
        stress = row.get("Magnetic Field Stress[mT]")
        drift_val = row.get("Off_drift_DU[mV]")
        if pd.notna(stress) and stress < 120 and pd.notna(drift_val):
            if drift_val < -2.5 or drift_val > 2.5:
                summary.append({
                    "Sample No": row.get("Sample No"),
                    "Plane": row.get("Plane"),
                    "Stress": stress,
                    "Angle": row.get("Angle"),
                    "Value": drift_val,
                    "Type": "Vdiff"
                })
    
    # Read Sensitivity_data for summary
    sens_sheet = wb["Sensitivity_data"]
    sens_df = pd.DataFrame(sens_sheet.values)
    sens_df.columns = sens_df.iloc[0]
    sens_df = sens_df[1:]
    sens_df.columns = sens_df.columns.str.strip()
    sens_df = sens_df.dropna(how="all")
    
    # Process Sensitivity_data drift values
    for col in ["Magnetic Field Stress[mT]", "Sens_drift_DU[%]"]:
        if col in sens_df.columns:
            sens_df[col] = pd.to_numeric(sens_df[col], errors='coerce')
    
    for idx, row in sens_df.iterrows():
        stress = row.get("Magnetic Field Stress[mT]")
        drift_val = row.get("Sens_drift_DU[%]")
        if pd.notna(stress) and stress < 120 and pd.notna(drift_val):
            if drift_val < -3 or drift_val > 3:
                summary.append({
                    "Sample No": row.get("Sample No"),
                    "Plane": row.get("Plane"),
                    "Stress": stress,
                    "Angle": row.get("Angle"),
                    "Value": drift_val,
                    "Type": "Sensitivity"
                })
    
    # Create Summary sheet
    if "Summary" in wb.sheetnames:
        wb.remove(wb["Summary"])
    summary_sheet = wb.create_sheet("Summary")
    
    # Write headers
    summary_sheet.append([
        "Sample No", "Plane", "Magnetic Field Stress[mT]", "Angle", "Out of Range Value", ".",
        "Sample No (Sensitivity)", "Plane (Sensitivity)", "Magnetic Field Stress[mT] (Sensitivity)", 
        "Angle (Sensitivity)", "Out of Range Value (Sensitivity)"
    ])
    
    # Write summary data - separate Vdiff and Sensitivity entries
    vdiff_entries = [e for e in summary if e["Type"] == "Vdiff"]
    sens_entries = [e for e in summary if e["Type"] == "Sensitivity"]
    max_rows = max(len(vdiff_entries), len(sens_entries))
    
    for i in range(max_rows):
        row_data = ["", "", "", "", "", ".", "", "", "", "", ""]
        if i < len(vdiff_entries):
            e = vdiff_entries[i]
            row_data[0:5] = [e["Sample No"], e["Plane"], e["Stress"], e["Angle"], e["Value"]]
        if i < len(sens_entries):
            e = sens_entries[i]
            row_data[6:11] = [e["Sample No"], e["Plane"], e["Stress"], e["Angle"], e["Value"]]
        summary_sheet.append(row_data)
    
    autofit_columns(wb, ["Vdiff_data", "Sensitivity_data", "Summary"])
    wb.save(output_file)

# Set mappings - convert string keys to integers
sample_mapping_raw = ${JSON.stringify(sampleMapping)}
axis_mapping_raw = ${JSON.stringify(axisMapping)}

# Convert string keys to integers for proper matching
sample_mapping = {int(k): v for k, v in sample_mapping_raw.items()}
axis_mapping = {int(k): v for k, v in axis_mapping_raw.items()}

print(f"Sample mapping: {sample_mapping}")
print(f"Axis mapping: {axis_mapping}")
sys.stdout.flush()

# Process the file
try:
    print("Starting file processing...")
    sys.stdout.flush()
    process_excel_file('Book1.xlsx', 'output.xlsx', sample_mapping, axis_mapping)
    print("File processing complete!")
    sys.stdout.flush()
except Exception as e:
    print(f"Error during processing: {str(e)}")
    import traceback
    traceback.print_exc()
    sys.stdout.flush()
    raise
`;
}

function getA3Script(sampleMapping, axisMapping) {
    // A3 version: 3 sweeps (U1, D, U2), 7 data points per group
    // Calculates U, UD, and DU parameters
    return `
import sys
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import load_workbook

print("Python script started (A3 version)...")
sys.stdout.flush()

def autofit_columns(workbook, sheet_names=None):
    if sheet_names is None:
        sheet_names = workbook.sheetnames
    
    for sheet_name in sheet_names:
        if sheet_name not in workbook.sheetnames:
            continue
            
        sheet = workbook[sheet_name]
        column_widths = {}
        
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is not None:
                    column_letter = cell.column_letter
                    cell_value = str(cell.value)
                    
                    if isinstance(cell.value, (int, float)):
                        display_width = len(cell_value) + 1
                    else:
                        display_width = len(cell_value)
                    
                    if column_letter not in column_widths:
                        column_widths[column_letter] = display_width
                    else:
                        column_widths[column_letter] = max(column_widths[column_letter], display_width)
        
        for column_letter, max_width in column_widths.items():
            adjusted_width = min(max(max_width + 2, 8), 50)
            sheet.column_dimensions[column_letter].width = adjusted_width

def process_excel_file(input_file, output_file, sample_mapping, axis_mapping):
    input_workbook = pd.ExcelFile(input_file)
    sheet_names = input_workbook.sheet_names

    all_vdiff_data = pd.DataFrame()
    all_sensitivity_data = pd.DataFrame()

    for sheet_name in sheet_names[2:]:
        if "precon" in sheet_name.lower():
            print(f"Skipping sheet: {sheet_name}")
            continue
            
        df = input_workbook.parse(sheet_name)

        columns_to_keep = [
            "Sample No", "Angle", "Axis", "B_stress[mT]", "Vdd[V]",
            "B_read[mT]", "B_set[mT]", "Temp_Gauss_probe[C]",
            "Vdiff_max[mV]", "Vdiff_min[mV]", "Vdiff_mean[mV]",
            "Vdiff_stdev[uV]"
        ]

        df = df.dropna(how='all', subset=columns_to_keep)
        vdiff_data = df[columns_to_keep].copy()

        columns_to_expand = ["Axis", "Angle", "B_stress[mT]", "Vdd[V]"]
        for col in columns_to_expand:
            vdiff_data[col] = np.repeat(vdiff_data[col].values, 7)[:len(vdiff_data)]

        sample_no_value = vdiff_data["Sample No"].iloc[0]
        vdiff_data["Sample No"] = vdiff_data["Sample No"].fillna(sample_no_value)
        vdiff_data.loc[vdiff_data["Axis"].isna(), "Sample No"] = np.nan

        sensitivity_data = vdiff_data.copy()
        sensitivity_data["Sens_U[mV/mT]"] = np.nan
        sensitivity_data["Sens_UD[mV/mT]"] = np.nan
        sensitivity_data["Sens_DU[mV/mT]"] = np.nan

        for i in range(0, len(sensitivity_data), 7):
            if i + 6 < len(sensitivity_data):
                group = sensitivity_data.iloc[i:i+7]
                
                try:
                    u1_b_set = group.iloc[0:3]["B_set[mT]"].values
                    u1_vdiff = group.iloc[0:3]["Vdiff_mean[mV]"].values
                    d_b_set = group.iloc[2:5]["B_set[mT]"].values
                    d_vdiff = group.iloc[2:5]["Vdiff_mean[mV]"].values
                    u2_b_set = group.iloc[4:7]["B_set[mT]"].values
                    u2_vdiff = group.iloc[4:7]["Vdiff_mean[mV]"].values
                    
                    sens_u1, sens_d, sens_u2 = np.nan, np.nan, np.nan
                    
                    if not (np.isnan(u1_vdiff).any() or np.isnan(u1_b_set).any()) and len(set(u1_b_set)) > 1:
                        sens_u1, _ = np.polyfit(u1_b_set, u1_vdiff, 1)
                    
                    if not (np.isnan(d_vdiff).any() or np.isnan(d_b_set).any()) and len(set(d_b_set)) > 1:
                        sens_d, _ = np.polyfit(d_b_set, d_vdiff, 1)
                    
                    if not (np.isnan(u2_vdiff).any() or np.isnan(u2_b_set).any()) and len(set(u2_b_set)) > 1:
                        sens_u2, _ = np.polyfit(u2_b_set, u2_vdiff, 1)
                    
                    sens_u = sens_u1 if not np.isnan(sens_u1) else np.nan
                    sens_ud = (sens_u1 + sens_d) / 2 if not (np.isnan(sens_u1) or np.isnan(sens_d)) else np.nan
                    sens_du = (sens_d + sens_u2) / 2 if not (np.isnan(sens_d) or np.isnan(sens_u2)) else np.nan
                    
                    for j in range(7):
                        if i + j < len(sensitivity_data):
                            sensitivity_data.loc[i+j, "Sens_U[mV/mT]"] = sens_u
                            sensitivity_data.loc[i+j, "Sens_UD[mV/mT]"] = sens_ud
                            sensitivity_data.loc[i+j, "Sens_DU[mV/mT]"] = sens_du
                
                except Exception as e:
                    print(f"Error calculating sensitivities for group starting at row {i}: {e}")
                    continue

        filtered_indices = []
        for i in range(0, len(sensitivity_data), 7):
            if i < len(sensitivity_data) and not pd.isna(sensitivity_data.iloc[i]["Sens_U[mV/mT]"]):
                filtered_indices.append(i)
        
        sensitivity_data = sensitivity_data.iloc[filtered_indices].reset_index(drop=True)

        column_rename_map = {
            "Axis": "Plane",
            "B_stress[mT]": "Magnetic Field Stress[mT]",
            "Vdd[V]": "Sample consumption[V]",
            "B_read[mT]": "Magnetic Field Read[mT]",
            "B_set[mT]": "Magnetic Field Set[mT]"
        }
        vdiff_data.rename(columns=column_rename_map, inplace=True)
        sensitivity_data.rename(columns=column_rename_map, inplace=True)

        all_vdiff_data = pd.concat([all_vdiff_data, vdiff_data], ignore_index=True)
        all_sensitivity_data = pd.concat([all_sensitivity_data, sensitivity_data], ignore_index=True)

    all_vdiff_data.dropna(how="all", inplace=True)
    all_sensitivity_data.dropna(how="all", inplace=True)

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        all_vdiff_data.to_excel(writer, sheet_name="Vdiff_data", index=False)
        all_sensitivity_data.to_excel(writer, sheet_name="Sensitivity_data", index=False)

    wb = load_workbook(output_file)
    for sheet_name in ["Vdiff_data", "Sensitivity_data"]:
        sheet = wb[sheet_name]
        for row in range(2, sheet.max_row + 1):
            sample_no_cell = sheet[f"A{row}"]
            axis_cell = sheet[f"C{row}"]

            if sample_no_cell.value in sample_mapping:
                sample_no_cell.value = sample_mapping[sample_no_cell.value]

            if axis_cell.value in axis_mapping:
                axis_cell.value = axis_mapping[axis_cell.value]

    autofit_columns(wb, ["Vdiff_data", "Sensitivity_data"])
    wb.save(output_file)

# Set mappings - convert string keys to integers
sample_mapping_raw = ${JSON.stringify(sampleMapping)}
axis_mapping_raw = ${JSON.stringify(axisMapping)}

sample_mapping = {int(k): v for k, v in sample_mapping_raw.items()}
axis_mapping = {int(k): v for k, v in axis_mapping_raw.items()}

print(f"Sample mapping: {sample_mapping}")
print(f"Axis mapping: {axis_mapping}")
sys.stdout.flush()

try:
    print("Starting file processing...")
    sys.stdout.flush()
    process_excel_file('Book1.xlsx', 'output.xlsx', sample_mapping, axis_mapping)
    print("File processing complete!")
    sys.stdout.flush()
except Exception as e:
    print(f"Error during processing: {str(e)}")
    import traceback
    traceback.print_exc()
    sys.stdout.flush()
    raise
`;
}

function getClassicScript(sampleMapping, axisMapping) {
    // Classic version: Simple 3-point groups, calculates K[mV/mT]
    return `
import sys
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import load_workbook

print("Python script started (classic version)...")
sys.stdout.flush()

def autofit_columns(workbook, sheet_names=None):
    if sheet_names is None:
        sheet_names = workbook.sheetnames
    
    for sheet_name in sheet_names:
        if sheet_name not in workbook.sheetnames:
            continue
            
        sheet = workbook[sheet_name]
        column_widths = {}
        
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is not None:
                    column_letter = cell.column_letter
                    cell_value = str(cell.value)
                    
                    if isinstance(cell.value, (int, float)):
                        display_width = len(cell_value) + 1
                    else:
                        display_width = len(cell_value)
                    
                    if column_letter not in column_widths:
                        column_widths[column_letter] = display_width
                    else:
                        column_widths[column_letter] = max(column_widths[column_letter], display_width)
        
        for column_letter, max_width in column_widths.items():
            adjusted_width = min(max(max_width + 2, 8), 50)
            sheet.column_dimensions[column_letter].width = adjusted_width

def process_excel_file(input_file, output_file, sample_mapping, axis_mapping):
    input_workbook = pd.ExcelFile(input_file)
    sheet_names = input_workbook.sheet_names

    all_vdiff_data = pd.DataFrame()
    all_sensitivity_data = pd.DataFrame()

    for sheet_name in sheet_names[2:]:
        if "precon" in sheet_name.lower():
            print(f"Skipping sheet: {sheet_name}")
            continue
            
        df = input_workbook.parse(sheet_name)

        columns_to_keep = [
            "Sample No", "Angle", "Axis", "B_stress[mT]", "Vdd[V]",
            "B_read[mT]", "B_set[mT]", "Temp_Gauss_probe[C]",
            "Vdiff_max[mV]", "Vdiff_min[mV]", "Vdiff_mean[mV]",
            "Vdiff_stdev[uV]"
        ]

        df = df.dropna(how='all', subset=columns_to_keep)
        vdiff_data = df[columns_to_keep].copy()

        columns_to_expand = ["Axis", "Angle", "B_stress[mT]", "Vdd[V]"]
        for col in columns_to_expand:
            vdiff_data[col] = np.repeat(vdiff_data[col].values, 3)[:len(vdiff_data)]

        sample_no_value = vdiff_data["Sample No"].iloc[0]
        vdiff_data["Sample No"] = vdiff_data["Sample No"].fillna(sample_no_value)
        vdiff_data.loc[vdiff_data["Axis"].isna(), "Sample No"] = np.nan

        sensitivity_data = vdiff_data.copy()
        sensitivity_data["K[mV/mT]"] = None

        for i in range(0, len(sensitivity_data), 3):
            if i + 2 < len(sensitivity_data):
                x_values = sensitivity_data.loc[i:i+2, "B_read[mT]"].values
                y_values = sensitivity_data.loc[i:i+2, "Vdiff_mean[mV]"].values

                if np.isnan(x_values).any() or np.isnan(y_values).any():
                    continue
                if len(set(x_values)) == 1 or len(set(y_values)) == 1:
                    continue

                try:
                    slope, _ = np.polyfit(x_values, y_values, 1)
                    sensitivity_data.loc[i, "K[mV/mT]"] = slope
                except Exception as e:
                    print(f"Error calculating slope for rows {i}-{i+2}: {e}")
                    continue

        sensitivity_data = sensitivity_data[sensitivity_data["K[mV/mT]"].notna()]

        column_rename_map = {
            "Axis": "Plane",
            "B_stress[mT]": "Magnetic Field Stress[mT]",
            "Vdd[V]": "Sample consumption[V]",
            "B_read[mT]": "Magnetic Field Read[mT]",
            "B_set[mT]": "Magnetic Field Set[mT]"
        }
        vdiff_data.rename(columns=column_rename_map, inplace=True)
        sensitivity_data.rename(columns=column_rename_map, inplace=True)

        all_vdiff_data = pd.concat([all_vdiff_data, vdiff_data], ignore_index=True)
        all_sensitivity_data = pd.concat([all_sensitivity_data, sensitivity_data], ignore_index=True)

    all_vdiff_data.dropna(how="all", inplace=True)
    all_sensitivity_data.dropna(how="all", inplace=True)

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        all_vdiff_data.to_excel(writer, sheet_name="Vdiff_data", index=False)
        all_sensitivity_data.to_excel(writer, sheet_name="Sensitivity_data", index=False)

    wb = load_workbook(output_file)
    for sheet_name in ["Vdiff_data", "Sensitivity_data"]:
        sheet = wb[sheet_name]
        for row in range(2, sheet.max_row + 1):
            sample_no_cell = sheet[f"A{row}"]
            axis_cell = sheet[f"C{row}"]

            if sample_no_cell.value in sample_mapping:
                sample_no_cell.value = sample_mapping[sample_no_cell.value]

            if axis_cell.value in axis_mapping:
                axis_cell.value = axis_mapping[axis_cell.value]

    autofit_columns(wb, ["Vdiff_data", "Sensitivity_data"])
    wb.save(output_file)

# Set mappings - convert string keys to integers
sample_mapping_raw = ${JSON.stringify(sampleMapping)}
axis_mapping_raw = ${JSON.stringify(axisMapping)}

sample_mapping = {int(k): v for k, v in sample_mapping_raw.items()}
axis_mapping = {int(k): v for k, v in axis_mapping_raw.items()}

print(f"Sample mapping: {sample_mapping}")
print(f"Axis mapping: {axis_mapping}")
sys.stdout.flush()

try:
    print("Starting file processing...")
    sys.stdout.flush()
    process_excel_file('Book1.xlsx', 'output.xlsx', sample_mapping, axis_mapping)
    print("File processing complete!")
    sys.stdout.flush()
except Exception as e:
    print(f"Error during processing: {str(e)}")
    import traceback
    traceback.print_exc()
    sys.stdout.flush()
    raise
`;
}

// Process button handler
document.getElementById('processBtn').addEventListener('click', async () => {
    if (!pyodide) {
        showStatus('Python environment not ready. Please wait...', 'error');
        return;
    }

    if (!inputFileData) {
        showStatus('Please select an input file first.', 'error');
        return;
    }

    document.getElementById('processBtn').disabled = true;
    showStatus('Starting processing...', 'loading');
    
    // Give UI time to update
    await new Promise(resolve => setTimeout(resolve, 100));

    try {
        // Parse sample mapping from UI
        const sampleMapping = parseMappingText(document.getElementById('sampleMapping').value);
        
        // Hardcoded axis mapping (not user-configurable)
        const axisMapping = {1: "XY", 2: "YZ", 3: "XZ"};
        
        console.log('Sample mapping:', sampleMapping);
        console.log('Axis mapping:', axisMapping);

        showStatus('Writing input file to memory...', 'loading');
        logToConsole('Writing input file to virtual filesystem...');
        await new Promise(resolve => setTimeout(resolve, 100));

        // Write input file to Pyodide filesystem
        pyodide.FS.writeFile('Book1.xlsx', inputFileData);
        logToConsole('Input file written successfully');

        showStatus('Executing Python script... This may take 30-60 seconds for large files.', 'loading');
        logToConsole('Starting Python execution...');
        await new Promise(resolve => setTimeout(resolve, 100));

        // Capture Python stdout and redirect to UI console
        pyodide.setStdout({
            batched: (msg) => {
                logToConsole(msg);
            }
        });

        // Prepare the Python script based on selected tab
        const pythonScript = getPythonScriptForTab(currentTab, sampleMapping, axisMapping);

        // Run the Python script
        console.log('Starting Python execution...');
        const result = await pyodide.runPythonAsync(pythonScript);
        console.log('Python execution completed:', result);

        showStatus('Reading output file...', 'loading');
        await new Promise(resolve => setTimeout(resolve, 100));

        // Read the output file from Pyodide filesystem
        logToConsole('Reading output file...');
        const outputData = pyodide.FS.readFile('output.xlsx');
        logToConsole(`Output file created (${outputData.length} bytes)`);

        // Create download blob and URL
        const blob = new Blob([outputData], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
        
        console.log('Creating download button...');
        
        // Create download button
        const downloadBtn = document.createElement('button');
        downloadBtn.className = 'download-btn';
        downloadBtn.textContent = 'Download Processed File';
        downloadBtn.style.display = 'inline-block';
        downloadBtn.onclick = () => {
            const a = document.createElement('a');
            a.href = url;
            a.download = 'processed_output.xlsx';
            a.click();
            logToConsole('File downloaded!');
        };
        
        console.log('Download button created:', downloadBtn);
        
        // Show success message with download button
        const statusDiv = document.getElementById('status');
        statusDiv.className = 'status success';
        statusDiv.innerHTML = 'Processing complete! Click the button below to download your file.<br><br>';
        statusDiv.appendChild(downloadBtn);
        
        console.log('Download button appended to status div');
        console.log('Status div HTML:', statusDiv.innerHTML);
        
        // Make absolutely sure the status div is visible
        statusDiv.style.display = 'block';
        
        logToConsole('Ready to download!');

        // Clean up Pyodide filesystem
        try {
            pyodide.FS.unlink('Book1.xlsx');
            pyodide.FS.unlink('output.xlsx');
        } catch (e) {
            console.log('Cleanup error (non-critical):', e);
        }

    } catch (error) {
        console.error('Processing error:', error);
        console.error('Error stack:', error.stack);
        
        // Show more detailed error message
        let errorMsg = 'Error processing file: ' + error.message;
        if (error.message.includes('PythonError')) {
            errorMsg += '<br><br>Check browser console (F12) for Python traceback.';
        }
        
        // Try to get Python traceback if available
        try {
            const traceback = pyodide.runPython(`
import sys
import traceback
traceback.format_exc()
            `);
            console.error('Python traceback:', traceback);
            errorMsg += '<br><br><pre style="text-align: left; font-size: 11px; max-height: 200px; overflow: auto;">' + traceback + '</pre>';
        } catch (e) {
            console.log('Could not get Python traceback');
        }
        
        showStatus(errorMsg, 'error');
    } finally {
        document.getElementById('processBtn').disabled = false;
    }
});

function parseMappingText(text) {
    const mapping = {};
    const lines = text.split('\n');
    for (const line of lines) {
        const trimmed = line.trim();
        if (trimmed && trimmed.includes(':')) {
            const [key, value] = trimmed.split(':', 2);
            const numKey = parseInt(key.trim());
            if (!isNaN(numKey)) {
                mapping[numKey] = value.trim();
            }
        }
    }
    return mapping;
}

function showStatus(message, type) {
    const statusDiv = document.getElementById('status');
    statusDiv.className = 'status ' + type;
    
    if (type === 'loading') {
        statusDiv.innerHTML = '<span class="spinner"></span>' + message;
    } else {
        statusDiv.innerHTML = message;
    }
}

function hideStatus() {
    document.getElementById('status').style.display = 'none';
}

// Initialize on page load
window.addEventListener('load', initPyodide);
