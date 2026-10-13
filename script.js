let currentWorkbook = null;

document.getElementById('excelFile').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            currentWorkbook = XLSX.read(data, { type: 'array' });
            runAnalysis();
        } catch (error) {
            console.error(error);
            showError("讀取 Excel 發生錯誤，請確認檔案格式是否正確。");
        }
    };
    reader.readAsArrayBuffer(file);
});

document.getElementById('analysisMode').addEventListener('change', function() {
    runAnalysis();
});

function runAnalysis() {
    if (!currentWorkbook) return;

    const mode = document.getElementById('analysisMode').value;
    const errorDiv = document.getElementById('errorMessage');
    const sheetSelectorWrapper = document.getElementById('sheetSelectorWrapper');
    
    errorDiv.classList.add('hidden');
    errorDiv.textContent = '';

    if (mode === 'smmi') {
        sheetSelectorWrapper.classList.add('hidden'); // SMMI 模式不需要工作表切換選單
        runSmmiAnalysis(currentWorkbook);
    } else if (mode === 'stats') {
        runStatsAnalysis(currentWorkbook); // 各項數據分析模式會啟動工作表選單與各欄位獨立拆分圖表
    }
}

function showError(message) {
    const errorDiv = document.getElementById('errorMessage');
    errorDiv.textContent = message;
    errorDiv.classList.remove('hidden');
    document.getElementById('chartsWrapper').innerHTML = '<p class="text-gray-400">無法顯示圖表</p>';
}