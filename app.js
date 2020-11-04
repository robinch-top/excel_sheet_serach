const fs = require('fs');
const path = require('path');
const readlineSync = require('readline-sync');
const xlsx = require('node-xlsx');

start()
/**
 * 开始函数
 */
async function start() {
    //等待输入
    var filePath = readlineSync.question("输入查找路径:");
    var Str = readlineSync.question("输入查找关键词:");
    //大写
    let findStr = Str.toUpperCase()
    let list = await findDirectory(filePath, findStr)
    if (list.length) {
        let header = [["url", "sheetname"]]
        let xlsxobj = [
            {
                name: 'sheet1',
                data: header.concat(list)
            }
        ]
        let dir = `./serach-${Str}-${Date.now()}.xlsx`
        //写入当前目录excel
        fs.writeFileSync(dir, xlsx.build(xlsxobj), "binary");
        console.log(`执行完毕,结果已写入  ${dir}`);
    } else {
        console.log(`未查询到！`);
    }
}
/**
 * 查找当前路径所有excel，返回符合关键字的url
 * @param {String} filePath 查找路径
 * @param {String} findStr 查找关键词
 */
async function findDirectory(filePath, findStr) {
    return new Promise(async (resolve, rejects) => {
        let files = await fileDisplay(filePath)
        let findlist = []
        for (let item of files) {
            var filedir = path.join(filePath, item);
            let stat = await fileIsFile(filedir)
            if (stat.isFile()) {
                let Suffix=filedir.split('.')[1]
                if(Suffix==='xlsx'||Suffix==='xls'){
                    let data = await findFile(filedir, findStr)
                    if (data.length) {
                        findlist = findlist.concat(data)
                    }
                }
            } else if (stat.isDirectory()) {
                let data = await findDirectory(filedir)
                findlist = findlist.concat(data)
            }
        }
        resolve(findlist)
    })
}
/**
 * 查找指定路径的单个excel，对符合关键字的表，进行记录
 * @param {String} path 
 * @param {String} str 
 */
function findFile(path, str) {
    return new Promise((resolve, rejects) => {
        let excel= xlsx.parse(path)
        let list = []
        for (let item of excel) {
            if (item.name.toUpperCase().indexOf(str) != -1) {
                console.log("符合的表路径:" + path);
                console.log("符合的表sheet名:" + item.name);
                list.push([path, item.name])
            }
        }
        resolve(list)
    })
};
/**
 * 读取当前路径的所有文件
 * @param {String} path 
 */
function fileDisplay(path) {
    return new Promise((resolve, rejects) => {
        fs.readdir(path, function (err, files) {
            if (err) {
                rejects(err)
            } else {
                resolve(files)
            }
        })
    })
}
/**
 * 判断文件的stat
 * @param {String} filedir 
 */
function fileIsFile(filedir) {
    return new Promise((resolve, rejects) => {
        fs.stat(filedir, function (eror, stats) {
            if (eror) {
                rejects(eror)
            } else {
                resolve(stats)
            }
        })
    })
}