<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>沃尔玛商品贴换标加工转换工具</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
</head>
<body class="bg-gray-50 p-6 font-sans">

<div class="max-w-6xl mx-auto bg-white p-8 rounded-lg shadow-md">
    <h1 class="text-2xl font-bold mb-6 text-gray-800">沃尔玛商品贴换标加工转换工具</h1>

    <div class="mb-8 p-4 border border-blue-200 bg-blue-50 rounded-md">
        <h2 class="text-lg font-semibold mb-3 text-blue-800">第一步：上传《在库库存明细表》</h2>
        <input type="file" id="inventoryFile" accept=".xlsx, .xls, .csv" class="block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded file:border-0 file:text-sm file:font-semibold file:bg-blue-100 file:text-blue-700 hover:file:bg-blue-200"/>
        <p class="text-xs text-gray-500 mt-2">系统会自动筛选：库存主体=深圳市德拉姆供应链有限公司，仓库名称=DLM供应链沃尔玛深圳仓-SZ，且可用库存>0的数据。</p>
    </div>

    <div class="mb-8 p-4 border border-gray-200 rounded-md">
        <div class="flex justify-between items-center mb-3">
            <h2 class="text-lg font-semibold text-gray-800">第二步：录入转换需求</h2>
            <button onclick="addRow()" class="bg-green-500 hover:bg-green-600 text-white px-4 py-2 rounded text-sm font-medium transition">+ 添加一行需求</button>
        </div>
        
        <table class="w-full text-left border-collapse" id="reqTable">
            <thead>
                <tr class="bg-gray-100 text-gray-700 text-sm">
                    <th class="p-2 border">SKU (必填)</th>
                    <th class="p-2 border">目标 FNSKU (必填)</th>
                    <th class="p-2 border">需求数量 (必填)</th>
                    <th class="p-2 border">备注 (选填)</th>
                    <th class="p-2 border w-16 text-center">操作</th>
                </tr>
            </thead>
            <tbody id="reqBody">
                <tr>
                    <td class="p-2 border"><input type="text" class="w-full p-1 border rounded req-sku" placeholder="输入SKU"></td>
                    <td class="p-2 border"><input type="text" class="w-full p-1 border rounded req-fnsku" placeholder="输入目标FNSKU"></td>
                    <td class="p-2 border"><input type="number" class="w-full p-1 border rounded req-qty" placeholder="输入数量" min="1"></td>
                    <td class="p-2 border"><input type="text" class="w-full p-1 border rounded req-note" placeholder="选填"></td>
                    <td class="p-2 border text-center"><button onclick="removeRow(this)" class="text-red-500 hover:text-red-700 font-bold">删除</button></td>
                </tr>
            </tbody>
        </table>
    </div>

    <div class="text-center">
        <button onclick="processData()" class="bg-blue-600 hover:bg-blue-700 text-white px-8 py-3 rounded-lg font-bold text-lg shadow-lg transition">执行匹配并导出转换表</button>
    </div>

    <div id="resultArea" class="mt-6 hidden">
        <h3 class="text-lg font-semibold mb-2">执行结果预览：</h3>
        <ul id="summaryList" class="text-sm space-y-1 bg-gray-100 p-4 rounded-md"></ul>
    </div>
</div>

<script>
    let inventoryData = [];

    // 监听文件上传并解析 Excel
    document.getElementById('inventoryFile').addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            inventoryData = XLSX.utils.sheet_to_json(firstSheet);
            alert(`成功读取库存表，共加载 ${inventoryData.length} 行数据。请继续录入需求。`);
        };
        reader.readAsArrayBuffer(file);
    });

    // 动态添加需求行
    function addRow() {
        const tbody = document.getElementById('reqBody');
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td class="p-2 border"><input type="text" class="w-full p-1 border rounded req-sku" placeholder="输入SKU"></td>
            <td class="p-2 border"><input type="text" class="w-full p-1 border rounded req-fnsku" placeholder="输入目标FNSKU"></td>
            <td class="p-2 border"><input type="number" class="w-full p-1 border rounded req-qty" placeholder="输入数量" min="1"></td>
            <td class="p-2 border"><input type="text" class="w-full p-1 border rounded req-note" placeholder="选填"></td>
            <td class="p-2 border text-center"><button onclick="removeRow(this)" class="text-red-500 hover:text-red-700 font-bold">删除</button></td>
        `;
        tbody.appendChild(tr);
    }

    // 动态删除需求行
    function removeRow(btn) {
        const row = btn.parentNode.parentNode;
        if (document.getElementById('reqBody').children.length > 1) {
            row.parentNode.removeChild(row);
        } else {
            alert("至少保留一行需求！");
        }
    }

    // 核心处理逻辑
    function processData() {
        if (inventoryData.length === 0) {
            alert("请先上传《在库库存明细表》！");
            return;
        }

        const reqRows = document.querySelectorAll('#reqBody tr');
        let requirements = [];
        reqRows.forEach(row => {
            const sku = row.querySelector('.req-sku').value.trim();
            const fnsku = row.querySelector('.req-fnsku').value.trim();
            const qty = parseInt(row.querySelector('.req-qty').value);
            const note = row.querySelector('.req-note').value.trim();
            if (sku && fnsku && qty > 0) {
                requirements.push({ sku, targetFnsku: fnsku, qty, note });
            }
        });

        if (requirements.length === 0) {
            alert("请正确填写至少一条完整需求（SKU、FNSKU、数量必填）！");
            return;
        }

        // 严格根据截图匹配列名：库存主体、仓库名称、可用库存
        let validPool = inventoryData.filter(row => 
            row['库存主体'] === '深圳市德拉姆供应链有限公司' &&
            row['仓库名称'] === 'DLM供应链沃尔玛深圳仓-SZ' &&
            Number(row['可用库存']) > 0
        );

        let outputExcelData = [];
        let summaryMsg = [];

        // 瀑布流匹配逻辑
        requirements.forEach(req => {
            let remainQty = req.qty;
            let currentFulfilled = 0;
            let rowResults = [];

            // 根据截图，SKU 列名是 SKU，目标FNSKU 匹配列是 FnSKU
            let skuInventory = validPool.filter(r => r['SKU'] === req.sku);
            
            // 优先级排序：FnSKU 一致的排在最前面
            skuInventory.sort((a, b) => {
                if (a['FnSKU'] === req.targetFnsku && b['FnSKU'] !== req.targetFnsku) return -1;
                if (a['FnSKU'] !== req.targetFnsku && b['FnSKU'] === req.targetFnsku) return 1;
                return 0;
            });

            for (let i = 0; i < skuInventory.length; i++) {
                if (remainQty <= 0) break;

                let item = skuInventory[i];
                let itemAvail = Number(item['可用库存']);
                if (itemAvail <= 0) continue;

                let deductQty = Math.min(remainQty, itemAvail);
                remainQty -= deductQty;
                currentFulfilled += deductQty;
                item['可用库存'] -= deductQty;

                // 如果原始 FnSKU 和目标不一致，才生成“贴换标”数据
                if (item['FnSKU'] !== req.targetFnsku) {
                    rowResults.push({
                        '转换类型': '贴换标',
                        '法人主体': '深圳市德拉姆供应链有限公司',
                        '仓库': 'DLM供应链沃尔玛深圳仓-SZ',
                        'SKU1': req.sku,
                        '库区1': '成品-存储1区',
                        'FNSKU1': item['FnSKU'],  // 提取原始的 FnSKU
                        '数量': deductQty,
                        'SKU2': req.sku,
                        'FNSKU2': req.targetFnsku,
                        '备注': req.note
                    });
                }
            }

            // 状态判断：使用 ✅ 和 ❌
            let statusMark = remainQty === 0 
                ? '✅ 满足' 
                : `❌ 不满足，只有${currentFulfilled}，缺${req.qty - currentFulfilled}`;

            // 写入生成的行
            rowResults.forEach(row => {
                row['状态'] = statusMark;
                outputExcelData.push(row);
            });

            // 界面结果总结
            summaryMsg.push(`需求 SKU: ${req.sku} | 目标 FNSKU: ${req.targetFnsku} | 需 ${req.qty} 个 -> <strong>${statusMark}</strong>`);
        });

        // 页面预览更新
        const resultArea = document.getElementById('resultArea');
        const summaryList = document.getElementById('summaryList');
        resultArea.classList.remove('hidden');
        summaryList.innerHTML = summaryMsg.map(msg => `<li class="border-b py-2">${msg}</li>`).join('');

        exportToExcel(outputExcelData);
    }

    function exportToExcel(dataRows) {
        if(dataRows.length === 0) {
            alert("匹配完成，但所有满足条件的库存均与目标FNSKU一致，无需生成贴换标单据。");
            return;
        }

        // 构建要求的表头结构
        const headerInfo = [
            ['注释：'],
            ['1. 黄色背景字段为必填'],
            ['2. 针对转换类型为贴换标和换品牌，其中仓库1、SKU1、库区1、FNSKU1、数量为加工对象，仓库2、SKU2、库区2、FNSKU2为加工结果'],
            ['', '', '', '', '', '', '', '', '', '', ''], // 空行，可根据截图需求微调
            ['转换类型', '法人主体', '仓库', 'SKU1', '库区1', 'FNSKU1', '数量', 'SKU2', 'FNSKU2', '备注', '状态']
        ];

        // 删除空行，严格适配您之前发的输出表格格式 (前三行是注释和空头，第四行是字段)
        headerInfo.splice(3, 1); 

        const bodyRows = dataRows.map(row => [
            row['转换类型'], row['法人主体'], row['仓库'], row['SKU1'], row['库区1'], 
            row['FNSKU1'], row['数量'], row['SKU2'], row['FNSKU2'], row['备注'], row['状态']
        ]);

        const finalSheetData = headerInfo.concat(bodyRows);
        const worksheet = XLSX.utils.aoa_to_sheet(finalSheetData);
        
        // 设置列宽
        worksheet['!cols'] = Array(11).fill({ wch: 18 });

        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "转换单");
        
        XLSX.writeFile(workbook, `贴换标转换单_${new Date().getTime()}.xlsx`);
    }
</script>
</body>
</html>
