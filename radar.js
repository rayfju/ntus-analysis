/**
 * radar.js - 正五角形分析圖 (單項獨立採計與缺失名單獨立顯示版)
 * 規格：
 * 1. 只要五項中有任意一項有數據，該項目就正常納入計算，並於表格顯示實測平均數。
 * 2. 缺失欄位不作 0 計算，不影響該生其他有測量項目的計算。
 * 3. 底下獨立顯示缺失人姓名及缺失項目。
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
            if (this.value === 'radar') {
                if (sheetWrapper) sheetWrapper.classList.add('hidden');
                if (gradeWrapper) gradeWrapper.classList.remove('hidden');
                if (schoolWrapper) schoolWrapper.classList.remove('hidden');

                if (globalWorkbook) {
                    initRadarMode(globalWorkbook);
                }
            } else {
                if (gradeWrapper) gradeWrapper.classList.add('hidden');
                if (schoolWrapper) schoolWrapper.classList.add('hidden');
            }
        });

        // 監聽年級或學校切換
        const triggerUpdate = () => {
            if (modeSelect.value === 'radar' && globalWorkbook) {
                renderRadarAnalysis(globalWorkbook, gradeSelect.value, schoolSelect.value);
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

                        if (modeSelect.value === 'radar') {
                            if (gradeWrapper) gradeWrapper.classList.remove('hidden');
                            if (schoolWrapper) schoolWrapper.classList.remove('hidden');
                            initRadarMode(globalWorkbook);
                        }
                    } catch (err) {
                        showRadarError('Excel 解析失敗：' + err.message);
                    }
                };
                reader.readAsArrayBuffer(file);
            });
        }
    });

    /**
     * 初始化雷達圖選單：載入年級與學校
     */
    function initRadarMode(workbook) {
        const availableGrades = getAllAvailableGrades(workbook);
        const gradeSelect = document.getElementById('gradeSelector');
        const schoolSelect = document.getElementById('schoolSelector');

        if (availableGrades.length === 0) {
            showRadarError('在上傳的 Excel 檔案中未發現年級資料！');
            return;
        }

        // 1. 填充年級選單
        if (gradeSelect) {
            gradeSelect.innerHTML = availableGrades.map(g => `<option value="${g}">${g}</option>`).join('');
            gradeSelect.value = availableGrades.includes('國三') ? '國三' : availableGrades[0];
        }

        // 2. 填充學校選單
        if (schoolSelect) {
            let options = `<option value="ALL">全體學校 (僅呈現標竿)</option>`;
            options += workbook.SheetNames.map(s => `<option value="${s}">${s}</option>`).join('');
            schoolSelect.innerHTML = options;
            schoolSelect.value = "ALL";
        }

        renderRadarAnalysis(workbook, gradeSelect.value, schoolSelect.value);
    }

    /**
     * 獲取所有存在的年級
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
     * 繪製重疊比較雷達圖與對照表
     */
    function renderRadarAnalysis(workbook, selectedGrade, selectedSchool) {
        const wrapper = document.getElementById('chartsWrapper');
        const errorDiv = document.getElementById('errorMessage');
        if (errorDiv) errorDiv.classList.add('hidden');

        // 1. 讀取全體該年級資料與缺失記錄
        const overallData = extractStudentsData(workbook, selectedGrade, "ALL");
        if (overallData.students.length === 0) {
            showRadarError(`未找到「${selectedGrade}」的學生資料！`);
            return;
        }

        // 計算全體五大指標平均數與標準差 (單項獨立採計)
        const overallStats = calculateFiveMetricsPairwise(overallData.students);

        // 2. 讀取特定學校該年級數據
        let schoolStats = null;
        let schoolZScores = null;
        let displayMissingList = overallData.missingRecords; // 預設呈現全體缺失名單

        if (selectedSchool !== "ALL") {
            const schoolData = extractStudentsData(workbook, selectedGrade, selectedSchool);
            if (schoolData.students.length > 0) {
                schoolStats = calculateFiveMetricsPairwise(schoolData.students);
                schoolZScores = calculateSchoolZScores(schoolStats, overallStats);
            }
            displayMissingList = schoolData.missingRecords; // 切換為該校缺失名單
        }

        // 渲染 UI 框架
        const titleText = selectedSchool === "ALL" 
            ? `🎯 ${selectedGrade} 正五角形目標標竿分析圖 (±1 SD)` 
            : `🎯 ${selectedSchool} (${selectedGrade}) vs 全體目標標竿 重疊比較圖`;

        wrapper.innerHTML = `
            <div class="bg-white p-6 rounded-xl border border-gray-200 shadow-sm space-y-6">
                <h2 class="text-xl font-bold text-gray-800 text-center border-b pb-3">${titleText}</h2>
                
                <!-- Plotly 雷達圖容器 -->
                <div id="radarPlotlyDiv" class="w-full h-[450px]"></div>

                <!-- 統計說明表格 -->
                <div class="overflow-x-auto mt-4">
                    <table class="min-w-full text-sm text-center text-gray-600 border border-gray-200 rounded-lg overflow-hidden">
                        <thead class="bg-blue-600 text-white font-semibold">
                            <tr>
                                <th class="py-2.5 px-4">五角指標</th>
                                <th class="py-2.5 px-4">全體有效樣本數</th>
                                <th class="py-2.5 px-4">全體平均 (Mean)</th>
                                <th class="py-2.5 px-4">計算公式</th>
                                <th class="py-2.5 px-4">全體標竿 (±1 SD)</th>
                                ${selectedSchool !== "ALL" ? `<th class="py-2.5 px-4 bg-orange-500">${selectedSchool} 實測平均</th>` : ''}
                            </tr>
                        </thead>
                        <tbody id="radarTableBody" class="divide-y divide-gray-200 bg-gray-50">
                        </tbody>
                    </table>
                </div>

                <!-- 缺失資料提示區塊 -->
                <div id="missingDataSection" class="mt-6 p-4 bg-amber-50 border border-amber-200 rounded-lg text-sm text-amber-800">
                </div>
            </div>
        `;

        // 填入對照表格
        const tbody = document.getElementById('radarTableBody');
        let tableHtml = '';
        for (let key in overallStats) {
            const item = overallStats[key];
            const formulaText = key === 'bodyFat' ? 'Mean - 1 × SD' : 'Mean + 1 × SD';
            
            // 學校實測平均文字呈現
            let schoolValText = '未測量';
            if (schoolStats && schoolStats[key] && schoolStats[key].mean !== null && !isNaN(schoolStats[key].mean)) {
                schoolValText = schoolStats[key].mean.toFixed(2);
            }

            tableHtml += `
                <tr class="hover:bg-gray-100 transition">
                    <td class="py-2.5 px-4 font-bold text-gray-800">${item.label}</td>
                    <td class="py-2.5 px-4">${item.count}</td>
                    <td class="py-2.5 px-4">${item.mean !== null ? item.mean.toFixed(2) : '未測量'}</td>
                    <td class="py-2.5 px-4 text-gray-500">${formulaText}</td>
                    <td class="py-2.5 px-4 font-bold text-blue-700">${item.target !== null ? item.target.toFixed(2) : '未測量'}</td>
                    ${selectedSchool !== "ALL" ? `<td class="py-2.5 px-4 font-bold ${schoolValText === '未測量' ? 'text-gray-400' : 'text-orange-600'}">${schoolValText}</td>` : ''}
                </tr>
            `;
        }
        tbody.innerHTML = tableHtml;

        // 渲染缺失學生清單
        renderMissingNotice(displayMissingList, selectedGrade, selectedSchool);

        // 繪製 Plotly 重疊雷達圖
        drawPlotlyRadar(selectedGrade, selectedSchool, schoolZScores);
    }

    /**
     * 讀取學生數據：有任意一項有數據即列入學生清單；缺失項紀錄於 missingRecords
     */
    function extractStudentsData(workbook, targetGrade, targetSchool) {
        const students = [];
        const missingRecords = [];
        const sheetsToScan = targetSchool === "ALL" ? workbook.SheetNames : [targetSchool];

        sheetsToScan.forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            if (!sheet) return;
            const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

            for (let i = 3; i < rows.length; i++) {
                const row = rows[i];
                if (!row || row.length === 0) continue;

                if (String(row[2]).trim() === targetGrade) {
                    const name = row[1] ? String(row[1]).trim() : '未命名';
                    
                    // 數據解析：必須為有效正數，否則填為 null（不作為 0 計算）
                    const parseVal = (v) => {
                        const num = parseFloat(v);
                        return (!isNaN(num) && num > 0) ? num : null;
                    };

                    const height = parseVal(row[4]);   // Col 4: 身高
                    const bodyFat = parseVal(row[7]);  // Col 7: 體脂
                    const smmi = parseVal(row[9]);     // Col 9: SMMI
                    const grip = parseVal(row[10]);    // Col 10: 握力
                    const power = parseVal(row[14]);    // Col 14: 爆發力

                    // 記錄缺測項目
                    const missingItems = [];
                    if (height === null) missingItems.push('身高');
                    if (bodyFat === null) missingItems.push('體脂肪率');
                    if (smmi === null) missingItems.push('SMMI');
                    if (grip === null) missingItems.push('慣用手握力');
                    if (power === null) missingItems.push('下肢爆發力');

                    // 只要「五項中至少有一項有數據」(missingItems 長度 < 5)，即加入學生名單
                    if (missingItems.length < 5) {
                        students.push({ height, bodyFat, smmi, grip, power, name, school: sheetName });
                    }

                    // 若有缺測項目，獨立記錄至 missingRecords 明細中
                    if (missingItems.length > 0) {
                        missingRecords.push({
                            name,
                            school: sheetName,
                            missingItems
                        });
                    }
                }
            }
        });

        return { students, missingRecords };
    }

    /**
     * 單項獨立採計 (Pairwise Calculation)：有數據的項目正常計算，不補 0
     */
    function calculateFiveMetricsPairwise(students) {
        const calcSingleMetric = (arr) => {
            // 過濾掉 null 與無效值
            const valid = arr.filter(v => v !== null && !isNaN(v) && v > 0);
            if (valid.length === 0) return { mean: null, sd: null, count: 0 };

            const mean = valid.reduce((a, b) => a + b, 0) / valid.length;
            const variance = valid.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / (valid.length > 1 ? valid.length - 1 : 1);
            return { mean, sd: Math.sqrt(variance), count: valid.length };
        };

        const h = calcSingleMetric(students.map(s => s.height));
        const bf = calcSingleMetric(students.map(s => s.bodyFat));
        const s = calcSingleMetric(students.map(s => s.smmi));
        const g = calcSingleMetric(students.map(s => s.grip));
        const p = calcSingleMetric(students.map(s => s.power));

        return {
            height: { ...h, target: h.mean !== null ? h.mean + 1 * h.sd : null, label: '身高 (cm)' },
            bodyFat: { ...bf, target: bf.mean !== null ? bf.mean - 1 * bf.sd : null, label: '體脂肪率 (%)' },
            smmi: { ...s, target: s.mean !== null ? s.mean + 1 * s.sd : null, label: 'SMMI' },
            grip: { ...g, target: g.mean !== null ? g.mean + 1 * g.sd : null, label: '慣用手握力 (kg)' },
            power: { ...p, target: p.mean !== null ? p.mean + 1 * p.sd : null, label: '下肢爆發力 (w)' }
        };
    }

    /**
     * 計算學校實測相對於全體常模的 Z 分數 (未測量者回傳 null，避免雷達圖暴衝)
     */
    function calculateSchoolZScores(schoolStats, overallStats) {
        const getZ = (schoolItem, overallItem, isReversed = false) => {
            if (!schoolItem || schoolItem.mean === null || !overallItem.sd || overallItem.sd === 0) {
                return null;
            }
            return isReversed 
                ? (overallItem.mean - schoolItem.mean) / overallItem.sd
                : (schoolItem.mean - overallItem.mean) / overallItem.sd;
        };

        return [
            getZ(schoolStats.height, overallStats.height),
            getZ(schoolStats.bodyFat, overallStats.bodyFat, true),
            getZ(schoolStats.smmi, overallStats.smmi),
            getZ(schoolStats.grip, overallStats.grip),
            getZ(schoolStats.power, overallStats.power),
            getZ(schoolStats.height, overallStats.height)
        ];
    }

    /**
     * 渲染缺失明細至圖表下方
     */
    function renderMissingNotice(missingList, gradeName, schoolName) {
        const container = document.getElementById('missingDataSection');
        if (!container) return;

        const schoolText = schoolName === "ALL" ? "全體學校" : schoolName;

        if (missingList.length === 0) {
            container.innerHTML = `
                <div class="flex items-center space-x-2 text-emerald-700 font-medium">
                    <span>✅ <strong>【數據完整】</strong> ${schoolText} 之 ${gradeName} 學生之各項體能數據皆完整填寫。</span>
                </div>
            `;
            return;
        }

        let listHtml = `
            <div class="font-bold text-amber-900 mb-2">
                ⚠️ 【缺失欄位提示】以下為 ${schoolText} (${gradeName}) 部分欄位未檢測/缺失之學生明細（缺失項目未採計於該項平均數）：
            </div>
            <ul class="list-disc pl-5 space-y-1 text-xs md:text-sm text-amber-800">
        `;

        missingList.forEach(item => {
            const itemsText = item.missingItems.join('、');
            listHtml += `
                <li>
                    <strong>${item.school} - ${item.name}</strong>：缺失項目為【<span class="text-red-600 font-semibold">${itemsText}</span>】
                </li>
            `;
        });

        listHtml += `</ul>`;
        container.innerHTML = listHtml;
    }

    /**
     * 繪製 Plotly 重疊雷達圖
     */
    function drawPlotlyRadar(gradeName, schoolName, schoolZScores) {
        const categories = ['身高 (cm)', '體脂肪率 (%)', 'SMMI', '慣用手握力 (kg)', '下肢爆發力 (w)', '身高 (cm)'];
        const targetValues = [1, 1, 1, 1, 1, 1];

        const traces = [
            {
                type: 'scatterpolar',
                r: targetValues,
                theta: categories,
                fill: 'toself',
                fillcolor: 'rgba(59, 130, 246, 0.15)',
                line: { color: '#2563eb', width: 2.5 },
                name: `${gradeName}目標標竿 (+1 SD)`
            }
        ];

        if (schoolName !== "ALL" && schoolZScores) {
            traces.push({
                type: 'scatterpolar',
                r: schoolZScores,
                theta: categories,
                fill: 'toself',
                fillcolor: 'rgba(249, 115, 22, 0.3)',
                line: { color: '#ea580c', width: 3 },
                name: `${schoolName} (${gradeName}) 實測平均`
            });
        }

        const layout = {
            polar: {
                angularaxis: {
                    rotation: 90,
                    direction: 'clockwise',
                    tickfont: { size: 13, color: '#1f2937', weight: 'bold' }
                },
                radialaxis: {
                    visible: true,
                    range: [-1, 2],
                    tickvals: [0, 1, 2],
                    ticktext: ['平均數 (0)', '+1 SD (目標)', '+2 SD'],
                    color: '#6b7280'
                }
            },
            margin: { t: 40, b: 40, l: 60, r: 60 },
            legend: { orientation: 'h', y: -0.15, x: 0.2 }
        };

        Plotly.newPlot('radarPlotlyDiv', traces, layout, { responsive: true, displayModeBar: false });
    }

    function showRadarError(msg) {
        const errorDiv = document.getElementById('errorMessage');
        if (errorDiv) {
            errorDiv.textContent = msg;
            errorDiv.classList.remove('hidden');
        }
    }
})();