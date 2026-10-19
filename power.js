/**
 * power.js - 下肢爆發力 (w) 與 SMMI 線性迴歸與帶狀比較分析模組
 * 具備年級與學校選單切換，包含三色線條 (+1SD 橘線 / 平均藍線 / -1SD 綠線)
 */

(function () {
    let globalWorkbook = null;

    document.addEventListener('DOMContentLoaded', function () {
        const fileInput = document.getElementById('excelFile');
        const modeSelect = document.getElementById('analysisMode');
        const sheetWrapper = document.getElementById('sheetSelectorWrapper');
        const gradeWrapper = document.getElementById('gradeSelectorWrapper');
        const schoolWrapper = document.getElementById('schoolSelectorWrapper');

        const gradeSelect = document.getElementById('gradeSelector');
        const schoolSelect = document.getElementById('schoolSelector');

        // 監聽模式切換
        modeSelect.addEventListener('change', function () {
            if (this.value === 'power') {
                if (sheetWrapper) sheetWrapper.classList.add('hidden');
                if (gradeWrapper) gradeWrapper.classList.remove('hidden');
                if (schoolWrapper) schoolWrapper.classList.remove('hidden');

                if (globalWorkbook) {
                    initPowerMode(globalWorkbook);
                }
            }
        });

        // 監聽年級或學校切換
        const triggerUpdate = () => {
            if (modeSelect.value === 'power' && globalWorkbook) {
                renderPowerAnalysis(globalWorkbook, gradeSelect.value, schoolSelect.value);
            }
        };

        if (gradeSelect) gradeSelect.addEventListener('change', triggerUpdate);
        if (schoolSelect) schoolSelect.addEventListener('change', triggerUpdate);

        // 監聽檔案上傳
        if (fileInput) {
            fileInput.addEventListener('change', function (e) {
                const file = e.target.files[0];
                if (!file) return;

                const reader = new FileReader();
                reader.onload = function (evt) {
                    try {
                        const data = new Uint8Array(evt.target.result);
                        globalWorkbook = XLSX.read(data, { type: 'array' });

                        if (modeSelect.value === 'power') {
                            if (gradeWrapper) gradeWrapper.classList.remove('hidden');
                            if (schoolWrapper) schoolWrapper.classList.remove('hidden');
                            initPowerMode(globalWorkbook);
                        }
                    } catch (err) {
                        showPowerError('Excel 解析失敗：' + err.message);
                    }
                };
                reader.readAsArrayBuffer(file);
            });
        }
    });

    /**
     * 初始化模式選單：年級與學校
     */
    function initPowerMode(workbook) {
        const availableGrades = getAllAvailableGrades(workbook);
        const gradeSelect = document.getElementById('gradeSelector');
        const schoolSelect = document.getElementById('schoolSelector');

        if (availableGrades.length === 0) {
            showPowerError('在上傳的 Excel 檔案中未發現年級資料！');
            return;
        }

        // 1. 填充年級選單
        if (gradeSelect) {
            gradeSelect.innerHTML = availableGrades.map(g => `<option value="${g}">${g}</option>`).join('');
            gradeSelect.value = availableGrades.includes('國三') ? '國三' : availableGrades[0];
        }

        // 2. 填充學校選單 (SheetNames)
        if (schoolSelect) {
            let options = `<option value="ALL">全體學校 (完整散佈點)</option>`;
            options += workbook.SheetNames.map(s => `<option value="${s}">${s}</option>`).join('');
            schoolSelect.innerHTML = options;
            schoolSelect.value = "ALL";
        }

        renderPowerAnalysis(workbook, gradeSelect.value, schoolSelect.value);
    }

    /**
     * 獲取所有年級
     */
    function getAllAvailableGrades(workbook) {
        const gradesSet = new Set();
        workbook.SheetNames.forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

            for (let i = 3; i < rows.length; i++) {
                const row = rows[i];
                if (row && row[2]) {
                    const g = String(row[2]).trim();
                    if (g) gradesSet.add(g);
                }
            }
        });
        return Array.from(gradesSet);
    }

    /**
     * 繪製分析頁面與圖表
     */
    function renderPowerAnalysis(workbook, selectedGrade, selectedSchool) {
        const wrapper = document.getElementById('chartsWrapper');
        const errorDiv = document.getElementById('errorMessage');
        if (errorDiv) errorDiv.classList.add('hidden');

        // 1. 抓取該年級「全體學校」學生數據 (計算基準迴歸線)
        const allStudents = extractStudents(workbook, selectedGrade, "ALL");
        if (allStudents.length < 2) {
            showPowerError(`「${selectedGrade}」的全體學生資料不足，無法繪製迴歸線！`);
            return;
        }

        // 2. 抓取特定學校學生數據 (作圖標記)
        const displayStudents = selectedSchool === "ALL" 
            ? allStudents 
            : extractStudents(workbook, selectedGrade, selectedSchool);

        // 渲染 HTML 結構
        const titleText = selectedSchool === "ALL"
            ? `⚡ ${selectedGrade} SMMI vs 下肢爆發力 (w) 線性迴歸帶狀圖`
            : `⚡ ${selectedSchool} (${selectedGrade}) vs 全體下肢爆發力迴歸帶狀圖`;

        wrapper.innerHTML = `
            <div class="bg-white p-6 rounded-xl border border-gray-200 shadow-sm space-y-4">
                <h2 class="text-xl font-bold text-gray-800 text-center border-b pb-3">${titleText}</h2>
                <div id="powerPlotlyDiv" class="w-full h-[500px]"></div>
            </div>
        `;

        // 繪製 Plotly 三色線與散佈圖
        drawPlotlyRegression(allStudents, displayStudents, selectedGrade, selectedSchool);
    }

    /**
     * 擷取學生 SMMI 與下肢爆發力 (w) 數據 (Col 9 & Col 14)
     */
    function extractStudents(workbook, targetGrade, targetSchool) {
        const list = [];
        const sheetsToScan = targetSchool === "ALL" ? workbook.SheetNames : [targetSchool];

        sheetsToScan.forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            if (!sheet) return;
            const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

            for (let i = 3; i < rows.length; i++) {
                const row = rows[i];
                if (!row) continue;

                if (String(row[2]).trim() === targetGrade) {
                    const smmi = parseFloat(row[9]);   // Col 9: SMMI
                    const power = parseFloat(row[14]); // Col 14: 下肢爆發力 (w)
                    const name = row[1] || '未命名';

                    // 過濾非數字、0 及負數
                    if (!isNaN(smmi) && !isNaN(power) && smmi > 0 && power > 0) {
                        list.push({
                            x: smmi,
                            y: power,
                            name: name,
                            school: sheetName
                        });
                    }
                }
            }
        });
        return list;
    }

    /**
     * 用 Plotly 繪製迴歸線與散佈點
     */
    function drawPlotlyRegression(allStudents, displayStudents, gradeName, schoolName) {
        // 全體年級數據 (計算迴歸參數)
        const xAll = allStudents.map(p => p.x);
        const yAll = allStudents.map(p => p.y);

        const n = xAll.length;
        const sumX = xAll.reduce((a, b) => a + b, 0);
        const sumY = yAll.reduce((a, b) => a + b, 0);
        const sumXY = xAll.reduce((acc, x, i) => acc + x * yAll[i], 0);
        const sumXX = xAll.reduce((acc, x) => acc + x * x, 0);

        // 迴歸斜率與截距
        const slope = (n * sumXY - sumX * sumY) / (n * sumXX - sumX * sumX);
        const intercept = (sumY - slope * sumX) / n;

        // 計算殘差與殘差標準差 (Residual SD)
        const residuals = yAll.map((y, i) => y - (slope * xAll[i] + intercept));
        const residualSD = Math.sqrt(residuals.reduce((acc, r) => acc + r * r, 0) / (n > 2 ? n - 2 : 1));

        // 產生帶狀線延伸 X 軸座標
        const minX = Math.min(...xAll) * 0.92;
        const maxX = Math.max(...xAll) * 1.08;
        const lineX = [minX, maxX];

        const lineY = lineX.map(x => slope * x + intercept);
        const upperY = lineX.map(x => slope * x + intercept + residualSD);
        const lowerY = lineX.map(x => slope * x + intercept - residualSD);

        // 待顯示的個案點座標與提示文字
        const xDisp = displayStudents.map(p => p.x);
        const yDisp = displayStudents.map(p => p.y);
        const textDisp = displayStudents.map(p => `${p.school} - ${p.name}`);

        const traces = [
            // 1. -1 SD 預測下限線 (綠色虛線)
            {
                x: lineX, y: lowerY, mode: 'lines',
                line: { color: '#2e7d32', dash: 'dash', width: 2 },
                name: `-1 SD 下限 (低於預期)`
            },
            // 2. +1 SD 預測上限線 (橘色虛線 + 藍色半透明帶狀填色)
            {
                x: lineX, y: upperY, mode: 'lines',
                fill: 'tonexty', fillcolor: 'rgba(187, 222, 251, 0.25)',
                line: { color: '#e65100', dash: 'dash', width: 2 },
                name: `+1 SD 上限 (高於預期)`
            },
            // 3. 全體年級迴歸平均線 (藍色實線)
            {
                x: lineX, y: lineY, mode: 'lines',
                line: { color: '#1976d2', width: 3 },
                name: `${gradeName}迴歸平均線 (y = ${slope.toFixed(2)}x + ${intercept.toFixed(2)})`
            },
            // 4. 學生實測個案點
            {
                x: xDisp, y: yDisp, text: textDisp, mode: 'markers',
                marker: {
                    size: schoolName === "ALL" ? 8 : 10,
                    color: schoolName === "ALL" ? '#6c757d' : '#d97706',
                    opacity: 0.85,
                    line: { width: 1, color: '#1f2937' }
                },
                hovertemplate: '<b>%{text}</b><br>SMMI: %{x}<br>爆發力: %{y} w<extra></extra>',
                name: schoolName === "ALL" ? `${gradeName}全體學生點` : `${schoolName} (${gradeName}) 學生點`
            }
        ];

        const layout = {
            title: { text: `${gradeName} SMMI vs 下肢爆發力 (w) 迴歸帶狀圖` },
            xaxis: { title: { text: 'SMMI (骨骼肌質量指數)' } },
            yaxis: { title: { text: '下肢爆發力 (w)' } },
            legend: { orientation: 'h', y: -0.2 }
        };

        Plotly.newPlot('powerPlotlyDiv', traces, layout, { responsive: true, displayModeBar: false });
    }

    function showPowerError(msg) {
        const errorDiv = document.getElementById('errorMessage');
        if (errorDiv) {
            errorDiv.textContent = msg;
            errorDiv.classList.remove('hidden');
        }
    }
})();