'use strict'

let order = ''
const pallets = {}
let currentPallet = ""
const boxes = {}
const result = []
const mixes = {}
const positions = {}
const withoutDM = ['былинная', 'гурьевская', 'суворовская', 'богатырская', 'княжеская', 'сухой', 'кг']

                                      // ЗАГРУЗКА EXCEL
async function handleFileAsync(e) {
  const file = e.target.files[0];
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data);
  const a = workbook.Sheets[workbook.SheetNames[0]];
  order = a['A1'] ? a['A1']['w'] : ""
  if (a['B3'] && a['B3']['w'].toLowerCase().includes('сборка')) {
    order = 'сборка'
    let currentArray = []
    Object.entries(a).forEach((i, c) => {
      if (i[0].includes("B")) {
        let currentAdress = i[0].match( /\d+/g )[0]
        if (i[1]["w"].toLowerCase().includes('комплект:')) {
          let nameMix = getNameMix(a['F' + currentAdress]["w"])
          mixes[nameMix] = []
          currentArray = mixes[nameMix]
          currentArray.push([a['B' + (currentAdress - 2)]["w"]])
          currentArray.push([" "])
        }
        if (a['D' + currentAdress]) {
          currentArray.push([i[1]["w"], a['D' + currentAdress]["w"], a['P' + currentAdress] ? a['P' + currentAdress]["w"] : '', a['W' + currentAdress]["w"], a['Z' + currentAdress]["w"]])
        }
        if (a['F' + currentAdress]) {
          currentArray.push([i[1]["w"], a['F' + currentAdress]["w"]])
        }
        if (i[1]["w"].toLowerCase().includes('ответственный')) {
          currentArray.push([" "])
          currentArray.push([i[1]["w"], "___________________", a['M' + currentAdress]["w"]])
        }
      }
    })
  }
  if (order.toLowerCase().includes('номенклатура')) {
    Object.entries(a).forEach((i, c) => {
      if (i[0].includes("C")) {
        if (i[1]["w"].length == 24) {
          result.push(boxes[i[1]["w"]])
          delete boxes[i[1]["w"]]
          return
        }
        let currentAdress = i[0].match( /\d+/g )[0]
        if (i[0] === 'C1') return
        let taste = a['A' + currentAdress]["w"]
        let mark = a['D' + currentAdress]["w"]
        boxes[mark] = [taste, "1", i[1]["w"], mark]
      }
    })
  }
  if (order.toLowerCase().includes('вб')) {
      Object.entries(a).forEach((i, c) => {
        if (i[0].includes("B")) {
          if (isNaN(i[1]["w"])) return
          let currentAdress = i[0].match( /\d+/g )[0]
          if (a['A' + currentAdress]["w"].includes('палет')) {
            pallets[a['A' + currentAdress]["w"]] = [[],[],[]]
            currentPallet = a['A' + currentAdress]["w"]
            return
          }
          pallets[currentPallet][getPath(a['A' + currentAdress]["w"],a['D' + currentAdress]["w"].slice(2,12))].push([a['A' + currentAdress]["w"], a['D' + currentAdress]["w"].slice(2,12), +i[1]["w"], +i[1]["w"], "-", +getCount(a['A' + currentAdress]["w"]) * +i[1]["w"]])
          if (getPath(a['A' + currentAdress]["w"],a['D' + currentAdress]["w"].slice(2,12)) == 1) {
            pallets[`${currentPallet} ${getNameMix(a['A' + currentAdress]["w"])}`] = mixes[getNameMix(a['A' + currentAdress]["w"])]
          }
        }  
      });
  }
  if (!order) {
      Object.entries(a).forEach((i, c) => {
    if (i[0].includes("C")) {
      if (i[1]["w"].toLowerCase().includes('вб')) order = i[1]["w"]
      if (i[1]["w"].toLowerCase().includes('озон')) order = i[1]["w"]
      if (i[1]["w"].toLowerCase().includes('маркет')) order = i[1]["w"] 
      let currentAdress = i[0].match( /\d+/g )[0]
      if (!a['A'+ currentAdress]?.["w"]) return
      if (isNaN(a['A'+ currentAdress]["w"])) return
      let c = a['C' + currentAdress]["w"]
      let d = a['D' + currentAdress]["w"]
      let f = a['F' + currentAdress]["w"]
      if (f.length < 3) f = 'палет ' + f
      let g = a['G' + currentAdress]["w"]
      let h = a['H' + currentAdress] ? a['H' + currentAdress]["w"] : 0
      let rowI = a['I' + currentAdress] ? a['I' + currentAdress]["w"] : ''
      if (pallets[f]) {
        pallets[f][getPath(c, d)].push([c, d, +g, +h, rowI, +getCount(c)*h])
      } else {
        pallets[f] = [[],[],[]]
        pallets[f][getPath(c, d)].push([c, d, +g, +h, rowI, +getCount(c)*h])
      }
      if (getPath(c, d) == 1) {
        pallets[`${f} ${getNameMix(c)}`] = mixes[getNameMix(c)]
      }
    }
  });
  }
  console.log(pallets)
  console.log(mixes)
}

function getNameMix(fullNameMix) {
  let indexOfMix = fullNameMix.indexOf('МИКС')
  return fullNameMix.slice(indexOfMix, indexOfMix + 8).trim()
}

function getPath(position, date) {
  if (withoutDM.some(i => position.toLowerCase().includes(i))) return 2
  if (position.includes('МИКС')) return 1
  const [day, month, year] = date.split('.')
  if ((year == 25) && (month < 3)) return 2
  if (year < 25) return 2
  return 0
}

function getCount(position) {
  const index = position.lastIndexOf('шт')
  const digit = position.substring(index - 4, index)
  if (digit) return digit.match( /\d+/g )[0]
  return 0
}

function sum(arr, element) {
  return arr.reduce((p,i)=>p+i[element],0)
}

                                      // СКАЧИВАНИЕ EXCEL

const te2 = document.querySelector(".test");
te2.addEventListener("change", handleFileAsync, false);

const downloadEx = document.querySelector(".downloadEx");

downloadEx.addEventListener("click", () => {
  var workbook = XLSX.utils.book_new();
  if (result.length) {
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet([...result]), "лист 1");
    const resArray = []
    for (let key in boxes) {
      resArray.push(boxes[key])
    }
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet([...resArray]), "лист 2");
  } else {
    const totalBoxes = []
    for (let key in pallets) {
      if (key.includes('МИКС')) {
        if (Array.isArray(pallets[key])) XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet([...pallets[key]]), key);
      } else {
        totalBoxes.push([sum(pallets[key][0], 3), sum(pallets[key][1], 3), sum(pallets[key][2], 3)])
        XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet([[order, "серия", "количество", "", "", "макри, шт."], [key, "", "шт", "уп", "ост"], ...pallets[key][0], [],...pallets[key][1], [], ...pallets[key][2], [], ['', 'количество', 'МОНО', 'МИКС', 'БЕЗ ЧЗ', 'ВСЕГО'], ['', 'коробки', sum(pallets[key][0], 3), sum(pallets[key][1], 3), sum(pallets[key][2], 3), sum(pallets[key][0], 3) + sum(pallets[key][1], 3) + sum(pallets[key][2], 3)], ['', 'марки', sum(pallets[key][0], 5), sum(pallets[key][1], 5)]]), key);
      }
}
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet([...totalBoxes]), "статус");
  }
  XLSX.writeFile(workbook, "Report.xlsx");
});