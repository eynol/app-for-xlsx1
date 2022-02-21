import xlsx from 'xlsx'
import fs from 'fs'
import path from 'path'



const workbook = xlsx.readFile('so.xlsx')
const newWorkbook = xlsx.utils.book_new();

const sheet = workbook.Sheets[workbook.SheetNames[0]]
console.log(sheet,sheet['!merges'])
const data = xlsx.utils.sheet_to_json<any>(sheet,{header:1})
console.log(data)
const title = data.shift();
const headers = data.shift();


// 从基本工资开始计算 加法 到 应发合计，然后后面是相减的值，直到实发工资
const start = headers.indexOf('基本工资')
const middle = headers.indexOf('应发合计')
const end = headers.indexOf('实发工资');

const sheetsToNewList = [];

for(let arr of data){
    if(!arr || arr.length===0){
        continue
    }
    let sum = 0;
    for(let i=start;i<middle;i++){
        if(typeof arr[i]==='number'){
            sum+=arr[i]
        }
    }
    if(arr[middle]!== sum){
        arr[middle] = sum;
    }
    for(let i= middle+1;i<end;i++){
        if(typeof arr[i]==='number'){
            arr[middle] -= arr[i]
        }
    }
    if(arr[end]!== sum){
        arr[end] = sum;
    }

    const userSheet:any[][]  = [
        [],
        []
    ] ;
    for(let i=0;i<headers.length;i++){
        if(arr[i]!==undefined){
            userSheet[0].push(headers[i])
            userSheet[1].push(arr[i])
        }
    }
    sheetsToNewList.push(userSheet)

}
const newSheet = xlsx.utils.aoa_to_sheet([
    title.concat(Array.from(headers).fill('').filter((_,index)=>index!==2)),
    headers,
    ...data
],)

// 合并单元格
newSheet['!merges'] = sheet['!merges']
newSheet.A1.s = { alignment: { wrapText: true, vertical: 'center', horizontal: 'center' } };
xlsx.utils.book_append_sheet(newWorkbook, newSheet,'生成的总表')



xlsx.writeFileXLSX(newWorkbook,'生成的总表.xlsx')



