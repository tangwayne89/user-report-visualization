// 使用 SheetJS 來讀取 Excel 文件
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.0/xlsx.full.min.js"></script>
<script>
    function handleFile(e) {
        var file = e.target.files[0];
        var reader = new FileReader();
        
        reader.onload = function(e) {
            var data = new Uint8Array(e.target.result);
            var workbook = XLSX.read(data, {type: 'array'});
            
            // 讀取第一個工作表
            var firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            // 將工作表轉換為 JSON 數據
            var excelData = XLSX.utils.sheet_to_json(firstSheet, {header: 1});
            
            console.log(excelData);  // 這裡可以查看讀取到的數據
        };
        
        reader.readAsArrayBuffer(file);
    }
</script>

<input type="file" onchange="handleFile(event)">
