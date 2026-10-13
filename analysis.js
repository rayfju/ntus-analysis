document.getElementById('excelFile').addEventListener('change', function(e) {
    const file = e.target.files[0];
    const errorDiv = document.getElementById('errorMessage');
    errorDiv.classList.add('hidden');
    errorDiv.textContent = '';

    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            const wrapper = document.getElementById('chartsWrapper');
            wrapper.innerHTML = ''; // 清空先前的圖表

            let successCount = 0;
            let allSheetsData = [];
            let allGlobalX = [];
            let allGlobalY = [];
            let allGlobalNames = [];

            // 迴圈遍歷 Excel 中的每一個分頁 (Worksheet)
            for (let sheetName of workbook.SheetNames) {
                const worksheet = workbook.Sheets[sheetName];
                const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                let smmiCol = -1;
                let jumpCol = -1;
                let nameCol = -1;
                let headerRowIdx = -1;

                // 1. 全域掃描所有儲存格，尋找欄位名稱
                for (let r = 0; r < rows.length; r++) {
                    const row = rows[r];
                    if (!row || row.length === 0) continue;

                    for (let c = 0; c < row.length; c++) {
                        const cellValue = String(row[c] || "").trim();
                        if (cellValue === 'SMMI') {
                            smmiCol = c;
                            headerRowIdx = Math.max(headerRowIdx, r);
                        }
                        if (cellValue === '垂直跳高度' || cellValue === '垂直跳高度(m)') {
                            jumpCol = c;
                            headerRowIdx = Math.max(headerRowIdx, r);
                        }
                        if (cellValue === '姓名' || cellValue === '名字' || cellValue.toLowerCase() === 'name') {
                            nameCol = c;
                        }
                    }
                }

                // 2. 如果成功找到必要欄位
                if (smmiCol !== -1 && jumpCol !== -1) {
                    const xArray = [];
                    const yArray = [];
                    const nameArray = [];

                    const startRow = headerRowIdx !== -1 ? headerRowIdx + 1 : 0;

                    for (let dataRowIdx = startRow; dataRowIdx < rows.length; dataRowIdx++) {
                        const dataRow = rows[dataRowIdx];
                        if (!dataRow) continue;

                        const xVal = parseFloat(dataRow[smmiCol]);
                        const yVal = parseFloat(dataRow[jumpCol]);
                        
                        let nameVal = `受試者 ${dataRowIdx}`;
                        if (nameCol !== -1 && dataRow[nameCol] !== null && dataRow[nameCol] !== undefined) {
                            const trimmedName = String(dataRow[nameCol]).trim();
                            if (trimmedName !== "") {
                                nameVal = trimmedName;
                            }
                        }

                        // 條件過濾：若數值為 NaN、等於 0 或缺失/空白，則不列入計算
                        if (
                            !isNaN(xVal) && !isNaN(yVal) &&
                            xVal !== 0 && yVal !== 0 &&
                            dataRow[smmiCol] !== null && dataRow[smmiCol] !== undefined && String(dataRow[smmiCol]).trim() !== "" &&
                            dataRow[jumpCol] !== null && dataRow[jumpCol] !== undefined && String(dataRow[jumpCol]).trim() !== ""
                        ) {
                            xArray.push(xVal);
                            yArray.push(yVal);
                            nameArray.push(nameVal);
                            
                            allGlobalX.push(xVal);
                            allGlobalY.push(yVal);
                            allGlobalNames.push(`[${sheetName}] ${nameVal}`);
                        }
                    }

                    if (xArray.length > 2) {
                        successCount++;
                        
                        // 建立各別工作表的獨立分析圖表區塊
                        const section = document.createElement('div');
                        section.className = "bg-white p-4 rounded-lg border border-gray-200 shadow-sm";
                        
                        const title = document.createElement('h2');
                        title.className = "text-lg font-bold text-gray-700 mb-3 border-b pb-2";
                        title.textContent = `工作表：${sheetName} (有效樣本數: ${xArray.length})`;
                        section.appendChild(title);

                        const chartDiv = document.createElement('div');
                        const chartId = `chart_${Math.random().toString(36).substring(2, 9)}`;
                        chartDiv.id = chartId;
                        chartDiv.className = "w-full h-[450px]";
                        section.appendChild(chartDiv);

                        wrapper.appendChild(section);

                        // 繪製該工作表的獨立圖表（含散佈點、回歸線、95% 信賴區間）
                        drawIndividualChart(chartId, sheetName, xArray, yArray, nameArray);

                        // 收集資料供最下方的多線疊加比較圖使用
                        allSheetsData.push({
                            sheetName: sheetName,
                            x: xArray,
                            y: yArray
                        });
                    }
                }
            }

            if (successCount === 0) {
                showError("在所有的工作表中，都找不到同時包含 'SMMI' 與 '垂直跳高度(m)' 且具備有效數值的資料。");
                return;
            }

            // 3. 在最下方新增：【各工作表迴歸趨勢線多線疊加比較圖】（類似你提供的範例圖片）
            if (allSheetsData.length > 0) {
                const trendSection = document.createElement('div');
                trendSection.className = "bg-gray-50 p-6 rounded-lg border-2 border-gray-300 shadow-md mt-10";
                
                const trendTitle = document.createElement('h2');
                trendTitle.className = "text-xl font-bold text-gray-800 mb-2 border-b border-gray-300 pb-2";
                trendTitle.textContent = `📈 各工作表迴歸趨勢線多線疊加比較`;
                trendSection.appendChild(trendTitle);

                const trendDesc = document.createElement('p');
                trendDesc.className = "text-sm text-gray-600 mb-4";
                trendDesc.textContent = "此圖表將所有工作表的線性回歸趨勢線以多線疊加方式呈現於同一張圖中，方便直接對比各組趨勢走向。（滑鼠移至線條上可顯示該工作表名稱）";
                trendSection.appendChild(trendDesc);

                const trendChartDiv = document.createElement('div');
                const trendChartId = `chart_trend_${Math.random().toString(36).substring(2, 9)}`;
                trendChartDiv.id = trendChartId;
                trendChartDiv.className = "w-full h-[500px] bg-white rounded-lg p-2";
                trendSection.appendChild(trendChartDiv);

                wrapper.appendChild(trendSection);

                // 繪製多線疊加趨勢圖
                drawMultiTrendChart(trendChartId, allSheetsData);
            }

            // 4. 同時保留最底下的【全體總合比較分析圖】
            if (allGlobalX.length > 2) {
                const globalSection = document.createElement('div');
                const globalSectionDiv = document.createElement('div');
                globalSectionDiv.className = "bg-indigo-50 p-6 rounded-lg border-2 border-indigo-200 shadow-md mt-6";
                
                const globalTitle = document.createElement('h2');
                globalTitle.className = "text-xl font-bold text-indigo-900 mb-2 border-b border-indigo-200 pb-2";
                globalTitle.textContent = `📊 全體總合比較分析 (所有工作表合併, 總有效樣本數: ${allGlobalX.length})`;
                globalSectionDiv.appendChild(globalTitle);

                const globalDesc = document.createElement('p');
                globalDesc.className = "text-sm text-indigo-700 mb-4";
                globalDesc.textContent = "此圖表為獨立的總體分析，將所有工作表的有效數據合併，呈現整體的線性回歸趨勢與 95% 信賴區間。";
                globalSectionDiv.appendChild(globalDesc);

                const globalChartDiv = document.createElement('div');
                const globalChartId = `chart_global_${Math.random().toString(36).substring(2, 9)}`;
                globalChartDiv.id = globalChartId;
                globalChartDiv.className = "w-full h-[500px] bg-white rounded-lg p-2";
                globalSectionDiv.appendChild(globalChartDiv);

                wrapper.appendChild(globalSectionDiv);

                drawIndividualChart(globalChartId, "全體總合比較分析", allGlobalX, allGlobalY, allGlobalNames, '#4f46e5');
            }

        } catch (error) {
            console.error(error);
            showError("讀取 Excel 發生錯誤，請確認檔案格式是否正確。");
        }
    };
    reader.readAsArrayBuffer(file);
});

function showError(message) {
    const errorDiv = document.getElementById('errorMessage');
    errorDiv.textContent = message;
    errorDiv.classList.remove('hidden');
    document.getElementById('chartsWrapper').innerHTML = '<p class="text-gray-400">無法顯示圖表</p>';
}

// 繪製「各別工作表獨立圖表」或「全體總合圖」的函式
function drawIndividualChart(containerId, titleText, x, y, names, markerColor = 'royalblue') {
    const n = x.length;
    
    const meanX = x.reduce((a, b) => a + b, 0) / n;
    const meanY = y.reduce((a, b) => a + b, 0) / n;
    
    let ssXX = 0;
    let ssXY = 0;
    for (let i = 0; i < n; i++) {
        ssXX += Math.pow(x[i] - meanX, 2);
        ssXY += (x[i] - meanX) * (y[i] - meanY);
    }
    
    const slope = ssXY / ssXX;
    const intercept = meanY - slope * meanX;

    let sse = 0;
    for (let i = 0; i < n; i++) {
        const yPred = slope * x[i] + intercept;
        sse += Math.pow(y[i] - yPred, 2);
    }
    const s = Math.sqrt(sse / (n - 2));

    const minX = Math.min(...x);
    const maxX = Math.max(...x);
    const numPoints = 100;
    const step = (maxX - minX) / (numPoints - 1);
    
    const xLine = [];
    const yUpper = [];
    const yLower = [];
    const yFit = [];
    const tVal = 1.96; 

    for (let i = 0; i < numPoints; i++) {
        const xi = minX + i * step;
        xLine.push(xi);
        
        const yi = slope * xi + intercept;
        yFit.push(yi);

        const se = s * Math.sqrt((1 / n) + Math.pow(xi - meanX, 2) / ssXX);
        yUpper.push(yi + tVal * se);
        yLower.push(yi - tVal * se);
    }

    const ciTraceUpper = {
        x: xLine,
        y: yUpper,
        mode: 'lines',
        type: 'scatter',
        line: { color: 'transparent' },
        showlegend: false,
        hoverinfo: 'skip'
    };

    const ciTraceLower = {
        x: xLine,
        y: yLower,
        mode: 'lines',
        type: 'scatter',
        fill: 'tonexty',
        fillcolor: markerColor === '#4f46e5' ? 'rgba(79, 70, 229, 0.15)' : 'rgba(37, 99, 235, 0.15)',
        line: { color: 'transparent' },
        name: '95% 信賴區間',
        hoverinfo: 'skip'
    };

    const regLine = {
        x: xLine,
        y: yFit,
        mode: 'lines',
        type: 'scatter',
        name: '線性回歸線',
        line: { color: markerColor === '#4f46e5' ? '#312e81' : 'crimson', width: 2.5 }
    };

    const scatterTrace = {
        x: x,
        y: y,
        text: names,
        mode: 'markers',
        type: 'scatter',
        name: '受試者數據',
        marker: { color: markerColor, size: 9, opacity: 0.75 },
        hoovertemplate: '<b>%{text}</b><br>SMMI: %{x}<br>垂直跳高度(m): %{y}<extra></extra>'
    };

    const layout = {
        title: { text: titleText.includes('工作表') ? `工作表：${titleText} (SMMI vs 垂直跳高度(m))` : titleText, font: { size: 16 } },
        xaxis: { title: 'SMMI', zeroline: false },
        yaxis: { title: '垂直跳高度(m)', zeroline: false },
        hovermode: 'closest',
        margin: { t: 40, b: 40, l: 60, r: 20 }
    };

    Plotly.newPlot(containerId, [ciTraceUpper, ciTraceLower, regLine, scatterTrace], layout, {responsive: true});
}

// 繪製「各工作表迴歸趨勢線多線疊加比較圖」的專屬函式（純線條多線比較，類似提供的範例圖）
function drawMultiTrendChart(containerId, allSheetsData) {
    let plotData = [];
    
    const colorPalette = [
        '#2563eb', // 藍
        '#dc2626', // 紅
        '#16a34a', // 綠
        '#9333ea', // 紫
        '#d97706', // 橙
        '#0891b2'  // 青
    ];

    allSheetsData.forEach((sheet, index) => {
        const color = colorPalette[index % colorPalette.length];
        const x = sheet.x;
        const y = sheet.y;
        const sheetName = sheet.sheetName;
        const n = x.length;

        // 計算線性回歸線
        const meanX = x.reduce((a, b) => a + b, 0) / n;
        const meanY = y.reduce((a, b) => a + b, 0) / n;
        
        let ssXX = 0;
        let ssXY = 0;
        for (let i = 0; i < n; i++) {
            ssXX += Math.pow(x[i] - meanX, 2);
            ssXY += (x[i] - meanX) * (y[i] - meanY);
        }
        
        const slope = ssXY / ssXX;
        const intercept = meanY - slope * meanX;

        const minX = Math.min(...x);
        const maxX = Math.max(...x);
        
        // 為了讓趨勢線平滑，取頭尾兩點或密集點連線
        const trendLine = {
            x: [minX, maxX],
            y: [minX * slope + intercept, maxX * slope + intercept],
            mode: 'lines',
            type: 'scatter',
            name: sheetName,
            line: { color: color, width: 3 },
            hovertemplate: `<b>工作表: ${sheetName}</b><br>迴歸趨勢線<extra></extra>`
        };

        plotData.push(trendLine);
    });

    const layout = {
        title: { text: '各工作表迴歸趨勢線多線疊加比較', font: { size: 16 } },
        xaxis: { title: 'SMMI', zeroline: false },
        yaxis: { title: '垂直跳高度(m)', zeroline: false },
        hovermode: 'closest',
        margin: { t: 40, b: 40, l: 60, r: 20 },
        legend: { orientation: 'h', y: -0.25 } // 讓圖例整齊排列在下方
    };

    Plotly.newPlot(containerId, plotData, layout, {responsive: true});
}