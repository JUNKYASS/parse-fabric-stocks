const axios = require('axios');
const XLSX = require('xlsx');
const { XMLParser } = require('fast-xml-parser');

const STOCKS_FILE_PATH = './stocks.xlsx';
const GALTEX_OZON_RESULT_FILE_PATH = './ready_stocks/galtex-ozon-stocks-updated.xlsx';
const GALTEX_WB_RESULT_FILE_PATH = './ready_stocks/galtex-wb-stocks-updated.xlsx';
const TD_OZON_RESULT_FILE_PATH = './ready_stocks/td-ozon-stocks-updated.xlsx';
const TD_WB_RESULT_FILE_PATH = './ready_stocks/td-wb-stocks-updated.xlsx';
const COMPARE_FILE_PATH = './compare.xlsx';

const BYAZ220_SHEETNAME = 'GT_Byaz_220';
const BYAZ150_SHEETNAME = 'GT_Byaz_150';
const BYAZ150GOST_SHEETNAME = 'GT_Byaz_150_Gost';
const POPLIN220_SHEETNAME = 'GT_Poplin_220';
const TD_SHEETNAME = 'TD';

const WAREHOUSE_ID = 'СЦ (Коляново) (1020002072018000)';
const NAME_POSTFIX = '+2% к прайсу';
const TD_DATA_EXPORT_URL = 'https://texdesign.ru/bitrix/catalog_export/cloth.xml';

const parseGaltexStocks = async () => {
  try {
    const workbook1 = XLSX.readFile(STOCKS_FILE_PATH);
    const sheetName1 = workbook1.SheetNames[0];
    const sheet1 = workbook1.Sheets[sheetName1];
    const data1 = XLSX.utils.sheet_to_json(sheet1, { header: 1, });

    const materialNameRow = data1[5][0] || undefined; // Определяем название материала
    if (!materialNameRow) return console.log('Material name empty');

    const sheetName = (
      materialNameRow.includes('Бязь') && materialNameRow.includes('220') ? BYAZ220_SHEETNAME :
        materialNameRow.includes('Бязь') && materialNameRow.includes('150см/120гр') ? BYAZ150_SHEETNAME :
          materialNameRow.includes('Бязь') && materialNameRow.includes('150см/140гр') ? BYAZ150GOST_SHEETNAME :
            materialNameRow.includes('Поплин') ? POPLIN220_SHEETNAME :
              undefined
    ); // Определяем название листа в зависимости от названия материала
    if (!sheetName) return console.log('Material name not found');

    const workbook2 = XLSX.readFile(COMPARE_FILE_PATH);
    const sheet2 = workbook2.Sheets[sheetName];
    if (!sheet2) return console.log('Sheet not found');
    const data2 = XLSX.utils.sheet_to_json(sheet2, { header: 1 });

    const stocksFileValues = data1.slice(6);
    const compareFileValues = data2.map(row => row[1]);

    // Формируем остатки
    const result = compareFileValues.filter(value => value).map((value, i) => { // В файле compare берём каждое значение и ищем его в файле stocks
      const valueMatch = stocksFileValues.filter(value2 => (value2[0] + NAME_POSTFIX).includes(value)); // Поиск всех совпадений (может быть одно или два)
      const greaterValue = valueMatch.length > 1 ? valueMatch[0][3] > valueMatch[1][3] ? valueMatch[0] : valueMatch[1] : valueMatch[0]; // Если одно совпадение, то берем его, если два, то берём то, в котором больше остаток

      const remain = greaterValue && greaterValue.length > 0 && greaterValue[3] > 600 ? 5 : 0;

      return [data2[i][0], data2[i][2], remain];
    });

    return result;
  } catch (error) {
    console.error(error);
  }
};

const parseTexdesignStocks = async () => {
  try {
    const response = await axios.get(TD_DATA_EXPORT_URL);
    if (response.statusText !== 'OK') throw new Error(`Ошибка загрузки XML: ${response.statusText}`);
    if (!response.headers['content-type']?.includes('xml')) throw new Error('Ответ не является XML-документом');

    const xmlRawData = await response.data.toString();
    if (!xmlRawData || xmlRawData.trim() === '') throw new Error('Загруженный XML-файл пустой или повреждён'); // Проверяем, что данные не пустые

    try {
      const parser = new XMLParser({
        ignoreAttributes: false,
      });

      xmlToJsonData = parser.parse(xmlRawData);
    } catch (parseError) {
      throw new Error(`Не удалось распарсить XML: ${parseError.message}`);
    }

    const allItems = xmlToJsonData.yml_catalog?.shop?.offers?.offer;
    if (!allItems || allItems.length === 0) throw new Error('XML-файл не содержит ни одного товара');

    const workbook = XLSX.readFile(COMPARE_FILE_PATH); // Получаем данные соответствия
    const sheet = workbook.Sheets[TD_SHEETNAME];
    const articlesData = XLSX.utils.sheet_to_json(sheet, { header: 1, });

    // Формируем остатки
    const result = articlesData.filter(value => value[1]).map((article, i) => {
      const matchedItem = allItems.find(item => item.param.find(param => param['@_name'] == 'Артикул')?.['#text'] == article[1]); // Ищем среди всех товаров совпадающий артикул

      const qty = matchedItem?.param?.find(param => param['@_name'] == 'Количество')?.['#text']; // Выбираем параметр "Количество"
      const remain = qty && qty > 600 ? 5 : 0;

      return [article[0], article[2], remain];
    });

    return result;
  } catch (error) {
    console.error(error);
  }
};

const createXLSXFile = (data, fileName) => { // Сохранение XLSX файла
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Результаты");
  XLSX.writeFile(wb, fileName);

  console.log(`Результат успешно записан в файл: ${fileName}`);
};

const main = async () => {
  try {
    const arg = process.argv.slice(2)[0];
    let data, resultOzon, resultWb;

    switch (arg) {
      case 'galtex': // Сохраняем остатки для Galtex
        data = await parseGaltexStocks();

        resultOzon = data.map(item => ({
          'Название склада (идентификатор склада)': WAREHOUSE_ID,
          'Артикул': item[0],
          'Название товара': '',
          'Доступно на складе, шт': item[2]
        }));

        resultWb = data.filter(elem => elem[1]).map(item => ({
          'Баркод': item[1],
          'Количество': item[2],
        }));

        createXLSXFile(resultOzon, GALTEX_OZON_RESULT_FILE_PATH);
        createXLSXFile(resultWb, GALTEX_WB_RESULT_FILE_PATH);

        break;
      case 'td': // Сохраняем остатки для Texdesign
        data = await parseTexdesignStocks();
        resultOzon = data.map(item => ({
          'Название склада (идентификатор склада)': WAREHOUSE_ID,
          'Артикул': item[0],
          'Название товара': '',
          'Доступно на складе, шт': item[2]
        }));

        resultWb = data.filter(elem => elem[1]).map(item => ({
          'Баркод': item[1],
          'Количество': item[2],
        }));

        createXLSXFile(resultOzon, TD_OZON_RESULT_FILE_PATH);
        createXLSXFile(resultWb, TD_WB_RESULT_FILE_PATH);

        break;
      default:
        console.log('Unknown argument');
    }
  } catch (error) {
    console.error(error);
  }
};

main();


// // Копирование баркодов из файла с баркодами в файл соответствия
// const workbook1 = XLSX.readFile(BARCODES_FILE_PATH);
// const sheetName1 = workbook1.SheetNames[2];
// const sheet1 = workbook1.Sheets[sheetName1];
// const barcodesData = XLSX.utils.sheet_to_json(sheet1, { header: 1, });

// const workbook2 = XLSX.readFile(COMPARE_FILE_PATH);
// const sheetName2 = workbook2.SheetNames[2];
// const sheet2 = workbook2.Sheets[sheetName2];
// const compareData = XLSX.utils.sheet_to_json(sheet2, { header: 1, });

// const result = compareData.map((compareValue, compareIndex) => {
//   const valueMatch = barcodesData.find(value2 => value2[1] === compareValue[0]);
//   return [...compareValue, valueMatch ? valueMatch[0] : ''];
// });

// // Сохраняем файл
// const ws = XLSX.utils.json_to_sheet(result);
// const wb = XLSX.utils.book_new();
// XLSX.utils.book_append_sheet(wb, ws, "Результаты");
// XLSX.writeFile(wb, './new.xlsx');