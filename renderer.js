// This file is required by the index.html file and will
// be executed in the renderer process for that window.
// All of the Node.js APIs are available in this process.

// const electron = require('electron');

// const dialog = electron.remote.dialog;

// dialog.showOpenDialog({
//     properties: ['openFile'],
//     filters: [{ name: 'excel', extensions: ['xlsx'] },]
// }, callback);

var fs = require('fs');
var path=require('path');
var xlsx = require('node-xlsx');


var cvtFlag = false
var excelPath = ''
var excelDir = ''
var excelBaseName =''

function convert(){
    var list= xlsx.parse(excelPath)
    var arr = list[0].data
    //console.log(arr)
    var title = arr.shift()
    arr = arr.filter(a=>a[3].trim()=='孵化器'||a[3].trim()=='入孵企业')
    for(let i in arr){
        arr[i][12]=new Date(1900, 0, arr[i][12] )//日期格式转换
    }
    console.log(arr)
    var fuhuaqi = arr.filter(a=>a[1].toString().indexOf("-")==-1).map(a=>[a[1].toString(),a[2],a[0]])
    var fuhuaqival = fuhuaqi.map(a=>a[1]) 
    for(let index in fuhuaqival){
        if(fuhuaqival.indexOf(fuhuaqival[index])!=fuhuaqival.lastIndexOf(fuhuaqival[index])){
            fuhuaqi[index][1]=fuhuaqi[index][2]+fuhuaqi[index][1]  //孵化器重名处理
        }
    }
    for(let i in fuhuaqi){
        fuhuaqi[i][1]=fuhuaqi[i][1].slice(0,30) //防止表名过长报错
    }
    var ct = []
    for(let item of fuhuaqi){
        let temp = arr.filter(a=>a[1].toString().indexOf(item[0])>-1) //
        let obj = {
            name:item[1],
            data:[title,...temp]
        }
        ct.push(obj)
    }
    fs.writeFileSync(excelDir + '/' + new Date().getTime()+'导出'+ excelBaseName,xlsx.build(ct),{'flag':'w'});
    alert('导出完毕!')

    excelPath = ''
    excelDir = ''
    excelBaseName =''
    btn.value=''
    cvtFlag = false
    
}
var btn = document.getElementById('btn')
var cvt = document.getElementById('cvt')

cvt.addEventListener('click',function(){
    if(cvtFlag){
        convert()
    }else{
        alert('请先选择文件!')
    }
})
btn.addEventListener("change",function (e) {
    //console.log(e.target.files[0].path);
    excelPath = e.target.files[0].path
    var extname = path.extname(excelPath);
    //console.log(extname)
    if(extname!=='.xlsx'){
        alert('文件格式不正确!')
        btn.value=''
        return
    }
    fs.stat(excelPath, function(err, stats) { 
        if (err) { 
        throw err; 
        } 
        console.log(stats); 
        excelDir = path.dirname(excelPath)
        excelBaseName = path.basename(excelPath)
        cvtFlag = true
        //console.log(excelDir)
    });
});