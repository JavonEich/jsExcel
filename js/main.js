/**
 * Created by paux on 15/6/1.
 */
!function(){
    var files = document.getElementById("files"),
        exImport = document.getElementById("import"),
        exExport = document.getElementById("export"),
        colName = document.getElementById("col"),
        find = document.getElementById("find"),
        tip = document.getElementById("tip"),
        rowFind = document.getElementById("rowFind"),
        startFind = document.getElementById("startFind"),
        re=/.(xlsx|XLSX)$/,
        file, f, colNameVal, findVal, exportData;

    files.addEventListener("change", function(e){
        file = e.target.files;
        f = file[0];
    }, false);
    exImport.addEventListener("click", function(){
        colNameVal = colName.value;
        findVal = find.value;
        if(re.test(f.name)){
            console.log("import:"+(new Date()));
            tip.innerHTML = "导入......";
            excelImport(f);
        }else{
            alert("文件格式须为xlsx");
        }
    }, false);
    function excelImport(f){
        var reader = new FileReader();
        reader.onload = function(e){
            var data = e.target.result,
                arr = fixdata(data),
                wb = XLSX.read(btoa(arr), {type: "base64"}),
                output = exportData = JSON.stringify(to_json(wb), 2, 2);
            if(out.innerText === undefined) {
                out.textContent = output;
            }
            else{
                out.innerText = output;
            }
            tip.innerHTML = "导入完成";
        }
        reader.readAsArrayBuffer(f);
    }
    function to_json(workbook) {
        var result = {};
        //遍历工作表
        workbook.SheetNames.forEach(function(sheetName) {
            var roa = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
            if (roa.length > 0) {
                var tem = [];
                if (!colNameVal || !findVal) {
                    result[sheetName] = roa;
                } else {
                    //筛选
                    roa.forEach(function (value, index) {
                        roa[index][colNameVal].indexOf(findVal) > -1 ? tem.push(roa[index]) : !1;
                    });
                    result[sheetName] = tem;
                }
                console.log(sheetName + ":" + result[sheetName].length);
            }
        });
        return result;
    }
    function fixdata(data) {
        var o = "", l = 0, w = 10240;
        for (; l < data.byteLength / w; ++l)
            o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w, l * w + w)));
        o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w)));
        return o;
    }
    //*********************************************
    exExport.addEventListener("click", function(){
        console.log("export:"+(new Date()));
        tip.innerHTML = "导出......";
        excelExport(exportData);
    }, false);
    function excelExport(data){
        var data = JSON.parse(data),
            wb = new Workbook(),
            wbOut;
        for(var obj in data){
            if(data.hasOwnProperty(obj)){
                wb.SheetNames.push(obj);
                wb.Sheets[obj] = sheet_from_array_of_arrays(data[obj]);
            }
        }
        wbOut = XLSX.write(wb, {bookType: "xlsx", bookSST: true, type: "binary"});
        saveAs(new Blob([s2ab(wbOut)], {type: "application/octet-stream"}), (+new Date()) + ".xlsx");
        tip.innerHTML = "导出完成";
        //保存为json
        //saveAs(new Blob([JSON.stringify(data)], {type: "text/plain;charset=UTF-8"}), "json"+(+new Date())+".txt");
    }
    function Workbook(){
        if(!(this instanceof Workbook))
            return new Workbook();
        this.SheetNames = [];
        this.Sheets = {};
    }
    function sheet_from_array_of_arrays(data){
        var ws = {};
        var range = {s: {c: 10000000, r: 10000000}, e: {c: 0, r: 0}};
        for(var R = 0; R != data.length; ++R){
            var C = 0;
            for(var val in data[R]) {
                if (range.s.r > R)
                    range.s.r = R;
                if (range.s.c > C)
                    range.s.c = C;
                if (range.e.r < R)
                    range.e.r = R;
                if (range.e.c < C)
                    range.e.c = C;
                var cell = {v: data[R][val]};
                if (cell.v == null) continue;
                if (typeof cell.v === "number") {
                    cell.t = "n";
                } else if (typeof cell.v === "boolean") {
                    cell.t = "b";
                }else if(cell.v instanceof Date){
                    cell.t = "n";
                    cell.z = XLSX.SSF._table[14];
                    cell.v = datenum(cell.v);
                }else{
                    cell.t = "s";
                }
                var cellRef = XLSX.utils.encode_cell({c: C, r: R});
                ws[cellRef] = cell;

                ++C;
            }
        }
        if (range.s.c < 10000000)
            ws['!ref'] = XLSX.utils.encode_range(range);
        return ws;
    }
    function s2ab(s){
        var len = s.length,
            buf = new ArrayBuffer(len),
            view = new Uint8Array(buf),
            i = 0;
        for(; i != len; ++i){
            view[i] = s.charCodeAt(i) & 0xFF;
        }
        return buf;
    }
    function datenum(v, date1904) {
        if(date1904)
            v+=1462;
        var epoch = Date.parse(v);
        return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
    }
}();
