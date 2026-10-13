// 初始化各項數據分析模式（當切換到此模式或更換工作表時觸發）
function runStatsAnalysis(workbook) {
    const sheetSelectorWrapper = document.getElementById('sheetSelectorWrapper');
    const sheetSelector = document.getElementById('sheetSelector');
    
    // 顯示工作表下拉選單
    sheetSelectorWrapper.classList.remove('hidden');

    // 動態填入所有工作表名稱到選單中
    sheetSelector.innerHTML = '';
    workbook.SheetNames.forEach(sheetName => {
        const option = document.createElement('option');
        option.value = sheetName;
        option.textContent = sheetName;
        sheetSelector.appendChild(option);
    });

    // 預設載入第一個工作表的數據
    renderStatsCharts(workbook, workbook.SheetNames[0]);

    // 監聽工作表選單切換事件
    sheetSelector.onchange = function() {
        renderStatsCharts(workbook, this.value);
    };
}

function renderStatsCharts(workbook, sheetName) {
    const wrapper = document.getElementById('chartsWrapper');
    wrapper.innerHTML = ''; // 清空畫面

    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    if (rows.length < 4) {
        wrapper.innerHTML = `<p class="text-gray-400 text-center py-10">工作表「${sheetName}」的資料列數不足。</p>`;
        return;
    }

    // 1. 抓取前三列作為多層次標題來源
    let row0 = rows[0] ? rows[0].map(h => String(h || "").trim()) : [];
    let row1 = rows[1] ? rows[1].map(h => String(h || "").trim()) : [];
    let row2 = rows[2] ? rows[2].map(h => String(h || "").trim()) : [];
    
    let finalHeaders = [];
    let maxCols = Math.max(row0.length, row1.length, row2.length);
    let cat1 = "";
    let cat2 = "";

    for (let c = 0; c < maxCols; c++) {
        let val0 = row0[c] || "";
        let val1 = row1[c] || "";
        let val2 = row2[c] || "";

        if (val0 !== "" && !val0.includes("編號") && !val0.includes("姓名")) {
            cat1 = val0;
        }
        if (val1 !== "" && !val1.includes("編號") && !val1.includes("姓名")) {
            cat2 = val1;
        }

        let parts = [];
        if (cat2 && cat2 !== cat1) parts.push(cat2);
        if (val2 && !parts.includes(val2)) parts.push(val2);
        else if (!val2 && val1 && !parts.includes(val1)) parts.push(val1);

        let combined = parts.join(" - ");
        if (!combined) combined = val0 || `欄位_${c+1}`;
        finalHeaders.push(combined);
    }

    let dataRowStart = 3; // 從第 4 列開始為真實受試者數據
    let colData = {}; 

    // 2. 略過 A 到 D 欄（索引 0, 1, 2, 3），從欄位索引 4 (E欄) 開始全部抓取
    for (let c = 4; c < finalHeaders.length; c++) {
        const headerName = finalHeaders[c];
        if (!headerName || headerName.includes("undefined")) continue;

        let columnValues = [];
        for (let r = dataRowStart; r < rows.length; r++) {
            const row = rows[r];
            if (!row) continue;
            const val = parseFloat(row[c]);
            if (!isNaN(val) && val !== 0) {
                columnValues.push(val);
            }
        }

        if (columnValues.length > 2) {
            colData[c] = {
                name: headerName,
                values: columnValues
            };
        }
    }

    const validIndices = Object.keys(colData);
    if (validIndices.length === 0) {
        wrapper.innerHTML = `<p class="text-gray-400 text-center py-10">工作表「${sheetName}」中，從 E 欄開始找不到可供分析的數值欄位。</p>`;
        return;
    }

    // 頂端提示說明區塊
    const topDesc = document.createElement('div');
    topDesc.className = "mb-6 p-4 bg-gray-900 text-gray-200 border border-gray-700 rounded-xl shadow-md text-sm";
    topDesc.innerHTML = `目前正在檢視工作表：<span class="font-bold text-blue-400 underline">${sheetName}</span></b>。`;
    wrapper.appendChild(topDesc);

    // 3. 針對每一個有效欄位，各自建立獨立的黑色卡片
    validIndices.forEach(colIdx => {
        const fieldInfo = colData[colIdx];
        const fieldName = fieldInfo.name;
        const values = fieldInfo.values;

        // 計算平均數與標準差
        const mean = values.reduce((a, b) => a + b, 0) / values.length;
        const variance = values.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / (values.length > 1 ? values.length - 1 : 1);
        const std = Math.sqrt(variance);

        const section = document.createElement('div');
        section.className = "bg-gray-900 text-white p-6 rounded-2xl border border-gray-800 shadow-xl mb-8 space-y-4";
        
        const headerDiv = document.createElement('div');
        headerDiv.innerHTML = `
            <h3 class="text-lg font-bold text-gray-100">${fieldName}</h3>
            <p class="text-xs text-gray-400 mt-1">Mean = <span class="text-blue-400 font-semibold">${mean.toFixed(2)}</span> | SD = <span class="text-red-400 font-semibold">±${std.toFixed(2)}</span> (樣本數: ${values.length})</p>
        `;
        section.appendChild(headerDiv);

        const chartDiv = document.createElement('div');
        const chartId = `chart_hist_${Math.random().toString(36).substring(2, 9)}`;
        chartDiv.id = chartId;
        chartDiv.className = "w-full h-[350px]";
        section.appendChild(chartDiv);
        wrapper.appendChild(section);

        // 4. 建立 Plotly 藍色直方圖 (Histogram)
        const histTrace = {
            x: values,
            type: 'histogram',
            name: '次數分佈 (直方圖)',
            marker: { color: '#3b82f6', opacity: 0.85, line: { color: '#1d4ed8', width: 1 } },
            xbins: { size: std > 0 ? std * 0.8 : undefined } // 依照標準差自動調整組距
        };

        // 5. 計算對應的紅色常態分佈平滑曲線座標
        const minX = Math.min(...values);
        const maxX = Math.max(...values);
        const rangeSpan = maxX - minX === 0 ? 1 : maxX - minX;
        const numPoints = 100;
        const step = rangeSpan / (numPoints - 1);

        const xCurve = [];
        const yCurve = [];

        // 估算長條圖總面積 (Total Area = 樣本數 * 組距)
        const nBins = 7; // 預估組數
        const binWidth = rangeSpan / nBins || 1;

        for (let i = 0; i < numPoints; i++) {
            const xVal = minX + i * step;
            xCurve.push(xVal);
            
            let yVal = 0;
            if (std > 0) {
                const exponent = -Math.pow(xVal - mean, 2) / (2 * Math.pow(std, 2));
                // 乘上樣本數與組距，讓曲線高度與直方圖的量級對齊
                yVal = (values.length * binWidth / (std * Math.sqrt(2 * Math.PI))) * Math.exp(exponent);
            }
            yCurve.push(yVal);
        }

        const curveTrace = {
            x: xCurve,
            y: yCurve,
            mode: 'lines',
            type: 'scatter',
            name: '常態分佈曲線',
            line: { color: '#ef4444', width: 3 } // 紅色平滑曲線
        };

        const layout = {
            title: { text: `${fieldName} - 直方圖與常態分佈曲線`, font: { size: 14, color: '#f3f4f6' } },
            paper_bgcolor: 'rgba(0,0,0,0)',
            plot_bgcolor: 'rgba(0,0,0,0)',
            font: { color: '#9ca3af' },
            xaxis: { 
                title: fieldName, 
                zeroline: false,
                gridcolor: '#374151'
            },
            yaxis: { 
                title: '頻數 (Count)', 
                zeroline: false,
                gridcolor: '#374151'
            },
            showlegend: true,
            legend: { orientation: 'h', y: -0.25, x: 0.15 },
            margin: { t: 40, b: 40, l: 50, r: 20 }
        };

        Plotly.newPlot(chartId, [histTrace, curveTrace], layout, {responsive: true, displayModeBar: false});
    });
}