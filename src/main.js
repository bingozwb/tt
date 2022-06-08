const xlsx = require('node-xlsx')
const fs = require('fs')

async function main() {
  // source: M + X
  // filter: C-DELIVER + L-外协加工费&实验检验费
  // compare: S + AB[2] + M[last]
  const file1 = 'source.xlsx'
  // target: BB + BC
  // compare: AI + D + O
  const file2 = 'target.xlsx'
  const data1 = getXlsx(file1).data
  // console.log('raw number: ', data1.length)
  // console.log('ceil number: ', data1[2].length)
  // console.log('title \n', data1[2])
  const data2 = getXlsx(file2).data
  // console.log('raw number: ', data2.length)
  // console.log('ceil number: ', data2[0].length)
  // console.log('title \n', data2[0])

  const filterData = getFilterData(data1)
  console.log(data2[0][53], data2[0][54])
  // match data
  for (let i = 1; i < data2.length; i++) {
    // skip if 'AI === 0'
    if (!data2[i][34]) {
      console.log('manual match ---------------------------------', i + 1)
      continue
    }
    const matchedItem = getMatchItem(data2[i], filterData)
    console.log(i, 'matchedItem', matchedItem)
    data2[i][53] = matchedItem[0]
    data2[i][54] = matchedItem[1]
  }

  // fill data
  const buffer = xlsx.build([{name: 'test', data: data2}])
  fs.writeFileSync('result.xlsx', buffer, {'flag':'w'})
}

function getFilterData(sourceData) {
  let filterData = []
  // console.log('sourceData.length', sourceData.length)
  // console.log(sourceData[2][2], sourceData[2][11])
  for (let i = 0; i < sourceData.length; i++) {
    if (sourceData[i][2] === 'DELIVER' && (sourceData[i][11] === '外协加工费' || sourceData[i][11] === '实验检验费' || sourceData[i][11] === '人员外包服务.其他类')) {
      filterData.push(sourceData[i])
    }
  }
  // console.log('filterData.length', filterData.length)
  return filterData
}

function getMatchItem(tItem, sourceData) {
  // console.log('sourceData.length', sourceData.length)
  const t1 = tItem[34]
  const t2 = tItem[3]
  const t3 = tItem[14]
  // console.log('t1, t2, t3', t1, t2, t3)
  let temRes = []
  let temRawNum
  for (let i = 0; i < sourceData.length; i++) {
    const s1 = sourceData[i][18]
    const s2 = sourceData[i][27].split('.')[1]
    let match = sourceData[i][12].match(/(?<=-)[A-Z,0-9]{14}/)
    const s3 = match ? match[0] : null
    // console.log('s1, s2, s3', s1, s2, s3)
    if (t1 === s1 && t2 === s2) {
      temRes = [sourceData[i][12], sourceData[i][23]]
      if (t3 === s3) {
        sourceData.splice(i, 1)
        return temRes
      } else {
        temRawNum = i
      }
    }
  }
  // every item MUST matched
  if (temRes.length === 0) {
    throw new Error('mismatch !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!')
  }
  sourceData.splice(temRawNum, 1)
  return temRes
}

function getXlsx(file) {
  console.log(file)
  //表格解析
  let sheetList = xlsx.parse(file)
  //对数据进行处理
  console.log('sheetList.length', sheetList.length)
  let sheet = sheetList[0]
  sheet.data.forEach((row, index) => {
    let rowIndex = index
    row.forEach((cell, index) => {
      let colIndex = index
      if (cell !== undefined) {
        // console.log('cell', cell)
        sheet.data[rowIndex][colIndex] = cell
        // sheet.data[rowIndex][colIndex] = cell.replace(/replaced text1/g, '').replace(/replaced text2/g, '')
        let reg = /\{([\u4e00-\u9fa5\.\w\:\、\/\d\s《》-]*)\|[\u4e00-\u9fa5\.\w\:\、\/\d\s《》-]*\}/
        let tempStr = sheet.data[rowIndex][colIndex]
        while (reg.test(tempStr)) {
          tempStr = tempStr.replace(reg, RegExp.$1)
        }
        sheet.data[rowIndex][colIndex] = tempStr
      }
    })
  })
  return sheet
}

// We recommend this pattern to be able to use async/await everywhere
// and properly handle errors.
main()
  .then(() => process.exit(0))
  .catch(error => {
    console.error(error)
    process.exit(1)
  })
