const fs = require('fs');
const axios = require('axios');
const https = require('https');
const XLSX = require('xlsx');
const { XMLParser } = require('fast-xml-parser');

const GT_BYAZ_220_120_SHEETNAME = 'GT_Byaz_220_120';
const GT_BYAZ_220_140_SHEETNAME = 'GT_Byaz_220_140';
const GT_BYAZ_150_120_SHEETNAME = 'GT_Byaz_150_120';
const GT_BYAZ_150_140_SHEETNAME = 'GT_Byaz_150_140';
const GT_BYAZ_150_120_SOLID_SHEETNAME = 'GT_Byaz_150_120_Solid';
const GT_BYAZ_150_140_SOLID_SHEETNAME = 'GT_Byaz_150_140_Solid';
const GT_POPLIN_220_SHEETNAME = 'GT_Poplin_220';
const TD_SHEETNAME = 'TD';
const AD_SHEETNAME = 'AD';
const TDL_BYAZ_220_SOLID_SHEETNAME = 'TDL_Byaz_220_solid';
const LOGOS_SHEETNAME = 'LB';
const TT_SHEETNAME = 'TT';

const WAREHOUSE_ID = 'СЦ (Коляново) (1020002072018000)';
const NAME_POSTFIX = '+2% к прайсу';
const TD_DATA_EXPORT_URL = 'https://texdesign.ru/bitrix/catalog_export/cloth.xml';

const GALTEX_OZON_RESULT_FILE_PATH = './ready_stocks/galtex-ozon-stocks-updated.xlsx';
const GALTEX_WB_RESULT_FILE_PATH = './ready_stocks/galtex-wb-stocks-updated.xlsx';
const TD_OZON_RESULT_FILE_PATH = './ready_stocks/td-ozon-stocks-updated.xlsx';
const TD_WB_RESULT_FILE_PATH = './ready_stocks/td-wb-stocks-updated.xlsx';
const AD_OZON_RESULT_FILE_PATH = './ready_stocks/ad-ozon-stocks-updated.xlsx';
const AD_WB_RESULT_FILE_PATH = './ready_stocks/ad-wb-stocks-updated.xlsx';
const TDL_OZON_RESULT_FILE_PATH = './ready_stocks/tdl-ozon-stocks-updated.xlsx';
const TDL_WB_RESULT_FILE_PATH = './ready_stocks/tdl-wb-stocks-updated.xlsx';
const LOGOS_OZON_RESULT_FILE_PATH = './ready_stocks/logos-ozon-stocks-updated.xlsx';
const LOGOS_WB_RESULT_FILE_PATH = './ready_stocks/logos-wb-stocks-updated.xlsx';
const TT_OZON_RESULT_FILE_PATH = './ready_stocks/tt-ozon-stocks-updated.xlsx';
const TT_WB_RESULT_FILE_PATH = './ready_stocks/tt-wb-stocks-updated.xlsx';
const MAPPING_FILE_PATH = './mapping.xlsx';
const XLSX_STOCKS_FILE_PATH = './stocks.xlsx';
const XLS_STOCKS_FILE_PATH = './stocks.xls';
const stocksWorkbook = fs.existsSync(XLSX_STOCKS_FILE_PATH)
  ? XLSX.readFile(XLSX_STOCKS_FILE_PATH)
  : fs.existsSync(XLS_STOCKS_FILE_PATH)
    ? XLSX.readFile(XLS_STOCKS_FILE_PATH)
    : false;

const stringToInt = (str) => parseFloat(str.toString().replace(/\s/g, '')); // Преобразование строки в число

const parseTDLStocks = async () => { // TDL (сделано только для бязи 220 однотонной, в будущем можно расширить)
  try {
    const stocksSheetName = stocksWorkbook.SheetNames[0];
    const stocksSheet = stocksWorkbook.Sheets[stocksSheetName];
    const stocksData = XLSX.utils.sheet_to_json(stocksSheet, { header: 1, });
    // console.log(stocksData);

    const mappingWorkBook = XLSX.readFile(MAPPING_FILE_PATH);
    const mappingSheet = mappingWorkBook.Sheets[TDL_BYAZ_220_SOLID_SHEETNAME];
    if (!mappingSheet) return console.log('Sheet not found');
    const mappingData = XLSX.utils.sheet_to_json(mappingSheet, { header: 1 });
    const filteredMappingData = mappingData.filter(row => row[3] !== 0); // Берём из файла mapping только те строки, в которых в 4 столбце не указано 0

    // console.log(mappingData);

    const result = filteredMappingData.map((mappingValue, i) => { // В файле mapping берём каждое значение и ищем его в файле stocks
      const valueMatch = stocksData.find(stocksValue => stocksValue[3] && (stocksValue[3].trim() == mappingValue[1])); // Поиск совпадения по артмкулу поставщика
      const remain = valueMatch && valueMatch.length > 0 && valueMatch[4] > 350 ? 5 : 0;

      return [mappingValue[0], mappingValue[3], remain]; // Возвращаем [артикул Озон, артикул ВБ, остаток]
    });
    // console.log(result);

    return result;
  } catch (error) {
    console.error(error);
  }
};

const parseLogosStocks = async () => {
  try {
    const stocksSheetName = stocksWorkbook.SheetNames[0];
    const stocksSheet = stocksWorkbook.Sheets[stocksSheetName];
    const stocksData = XLSX.utils.sheet_to_json(stocksSheet, { header: 1, });
    // console.log(stocksData);

    const mappingWorkBook = XLSX.readFile(MAPPING_FILE_PATH);
    const mappingSheet = mappingWorkBook.Sheets[LOGOS_SHEETNAME];
    if (!mappingSheet) return console.log('Sheet not found');
    const mappingData = XLSX.utils.sheet_to_json(mappingSheet, { header: 1 });
    const filteredMappingData = mappingData.filter(row => row[3] !== 0); // Берём из файла mapping только те строки, в которых в 4 столбце не указано 0
    // console.log(mappingData);

    const result = filteredMappingData.map((mappingValue, i) => { // В файле mapping берём каждое значение и ищем его в файле stocks
      const valueMatch = stocksData.find(stocksValue => stocksValue[0] && (stocksValue[0].trim() == mappingValue[1])); // Поиск совпадения по артмкулу поставщика
      const remain = valueMatch && valueMatch.length > 0 && valueMatch[15] > 350 ? 10 : 0;
      return [mappingValue[0], mappingValue[2], remain]; // Возвращаем [артикул Озон, артикул ВБ, остаток]
    });
    // console.log(result);

    return result;
  } catch (error) {
    console.error(error);
  }
};

const parseTTStocks = async () => {
  try {
    const stocksSheetName = stocksWorkbook.SheetNames[0];
    const stocksSheet = stocksWorkbook.Sheets[stocksSheetName];
    const stocksData = XLSX.utils.sheet_to_json(stocksSheet, { header: 1, });
    const poplinRowIndex = stocksData.findIndex(element => element.length === 1 && element[0] === 'Поплин'); // Ищем строку Поплин и берём её индекс далее разделяем массив на две части, оставляем только ту, которая идёт до строки Поплин (т.е. берём только бязь) (в остатках есть только бязь и поплин в одной таблице, поэтому нужно разделить их на две части)
    const poplinStocks = poplinRowIndex !== -1 ? stocksData.slice(poplinRowIndex) : stocksData; // Берём из файла stocks только те строки, которые идут после строки Поплин
    const byazStocks = poplinRowIndex !== -1 ? stocksData.slice(0, poplinRowIndex) : stocksData; // Берём из файла stocks только те строки, которые идут до строки Поплин (т.е. берём только бязь)
    // console.log('poplinStocks', poplinStocks)
    // console.log('byazStocks', byazStocks);

    const mappingWorkBook = XLSX.readFile(MAPPING_FILE_PATH);
    const mappingSheet = mappingWorkBook.Sheets[TT_SHEETNAME];
    if (!mappingSheet) return console.log('Sheet not found');
    const mappingData = XLSX.utils.sheet_to_json(mappingSheet, { header: 1 });
    const filteredMappingData = mappingData.filter(row => row[3] !== 0); // Берём из файла mapping только те строки, в которых в 4 столбце не указано 0
    // console.log(mappingData);

    const result = filteredMappingData.map((mappingValue, i) => { // В файле mapping берём каждое значение и ищем его в файле stocks
      let valueMatch;

      if (mappingValue[0] && mappingValue[0].includes('-BZ-')) { // Если в артикуле поставщика есть -BZ-, значит это бязь, ищем совпадение в массиве byazStocks
        valueMatch = byazStocks.find(stocksValue => stocksValue[2] && (stocksValue[2].trim() == mappingValue[1])); // Поиск совпадения по артмкулу поставщика
      } else if (mappingValue[0] && mappingValue[0].includes('-PP-')) { // Если в артикуле поставщика есть -PP-, значит это поплин, ищем совпадение в массиве poplinStocks
        valueMatch = poplinStocks.find(stocksValue => stocksValue[2] && (stocksValue[2].trim() == mappingValue[1])); // Поиск совпадения по артмкулу поставщика
      } else { // Если в артикуле поставщика нет -BZ- или -PP-, значит это что-то другое, ищем совпадение в массиве stocksData
        valueMatch = stocksData.find(stocksValue => stocksValue[2] && (stocksValue[2].trim() == mappingValue[1])); // Поиск совпадения по артмкулу поставщика
      }

      const remain = valueMatch && valueMatch.length > 0 && stringToInt(valueMatch[3]) > 250 ? 10 : 0;
      return [mappingValue[0], mappingValue[2], remain]; // Возвращаем [артикул Озон, артикул ВБ, остаток]
    });
    // console.log(result);

    return result;
  } catch (error) {
    console.error(error);
  }
};

const parseArtdesignStocks = async () => {
  try {
    const stocksSheetName = stocksWorkbook.SheetNames[0];
    const stocksSheet = stocksWorkbook.Sheets[stocksSheetName];
    const stocksData = XLSX.utils.sheet_to_json(stocksSheet, { header: 1, });

    const mappingWorkBook = XLSX.readFile(MAPPING_FILE_PATH);
    const mappingSheet = mappingWorkBook.Sheets[AD_SHEETNAME];
    if (!mappingSheet) return console.log('Sheet not found');
    const mappingData = XLSX.utils.sheet_to_json(mappingSheet, { header: 1 });
    const filteredMappingData = mappingData.filter(row => row[4] !== 0); // Берём из файла mapping только те строки, в которых в 4 столбце не указано 0

    const result = filteredMappingData.map((mappingValue, i) => { // В файле mapping берём каждое значение и ищем его в файле stocks
      const valueMatch = stocksData.find(stocksValue => stocksValue[1] && (stocksValue[1].trim() == mappingValue[2])); // Поиск совпадения по артмкулу поставщика
      const remain = valueMatch && valueMatch.length > 0 && valueMatch[4] > 600 ? 5 : 0;

      return [mappingValue[0], mappingValue[3], remain]; // Возвращаем [артикул Озон, артикул ВБ, остаток]
    });

    return result;
  } catch (error) {
    console.error(error);
  }
}

const parseGaltexStocks = async () => {
  try {
    const stocksSheetName = stocksWorkbook.SheetNames[0];
    const stocksSheet = stocksWorkbook.Sheets[stocksSheetName];
    const stocksData = XLSX.utils.sheet_to_json(stocksSheet, { header: 1, });

    const materialNameRowIndex = stocksData.findIndex(value => value[0] === 'Характеристика') + 1; // Ищем строку Характеристика номенклатуры и берём следующую за ней строку
    const materialNameRow = stocksData[materialNameRowIndex][0] || undefined; // Определяем название материала
    if (!materialNameRow) return console.log('Material name empty');

    const sheetName = (
      materialNameRow.includes('Бязь') && materialNameRow.includes('(220см/120гр) наб') ? GT_BYAZ_220_120_SHEETNAME :
        materialNameRow.includes('Бязь') && materialNameRow.includes('(220см/140гр) наб') ? GT_BYAZ_220_140_SHEETNAME :
          materialNameRow.includes('Бязь') && materialNameRow.includes('(150см/120гр) наб') ? GT_BYAZ_150_120_SHEETNAME :
            materialNameRow.includes('Бязь') && materialNameRow.includes('(150см/140гр) наб') ? GT_BYAZ_150_140_SHEETNAME :
              materialNameRow.includes('Бязь') && materialNameRow.includes('(150см/120гр) гл/кр') ? GT_BYAZ_150_120_SOLID_SHEETNAME :
                materialNameRow.includes('Бязь') && materialNameRow.includes('(150см/140гр) гл/кр') ? GT_BYAZ_150_140_SOLID_SHEETNAME :
                  materialNameRow.includes('Поплин') ? GT_POPLIN_220_SHEETNAME :
                    undefined
    ); // Определяем название листа в зависимости от названия материала
    if (!sheetName) return console.log('Material name not found');

    // Определяем номер столбца из которого брать количество остатков
    const kharakteristikaRow = stocksData.find(value => value[0] === 'Характеристика');
    if (!kharakteristikaRow) return console.log('kharakteristika row not found');
    const stocksCountHeadingIndex = kharakteristikaRow.findIndex(value => value === 'Остаток');

    const mappingWorkBook = XLSX.readFile(MAPPING_FILE_PATH);
    const mappingSheet = mappingWorkBook.Sheets[sheetName];
    if (!mappingSheet) return console.log('Sheet not found');
    const mappingData = XLSX.utils.sheet_to_json(mappingSheet, { header: 1 });
    const filteredMappingData = mappingData.filter(row => row[0] && row[1] && row[3] !== 0); // Берём из файла mapping только те строки, в которых в 4 столбце не указано 0
    const stocksFileValues = stocksData.slice(5); // Берём из файла stocks только те строки, которые идут после технических строк

    // Формируем остатки
    const result = filteredMappingData.map((value, i) => { // В файле mapping берём каждое значение и ищем его в файле stocks

      const valueMatch = stocksFileValues.filter(value2 => (value2[0] + NAME_POSTFIX).includes(value[1])); // Поиск всех совпадений (может быть одно или два)
      const greaterValue = valueMatch.length > 1 ? valueMatch[0][stocksCountHeadingIndex] > valueMatch[1][stocksCountHeadingIndex] ? valueMatch[0] : valueMatch[1] : valueMatch[0]; // Если одно совпадение, то берем его, если два, то берём то, в котором больше остаток
      const remain = greaterValue && greaterValue.length > 0 && stringToInt(greaterValue[stocksCountHeadingIndex]) > 400 ? 10 : 0;
      return [filteredMappingData[i][0], filteredMappingData[i][2], remain];
    });

    return result;
  } catch (error) {
    console.error(error);
  }
};

const parseTexdesignStocks = async () => {
  try {
    const agent = new https.Agent({
      rejectUnauthorized: false
    }); // Отключаем проверку сертификата

    const response = await axios.get(TD_DATA_EXPORT_URL, { httpsAgent: agent }); // Загружаем XML
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

    const workbook = XLSX.readFile(MAPPING_FILE_PATH); // Получаем данные соответствия
    const sheet = workbook.Sheets[TD_SHEETNAME];
    const mappingData = XLSX.utils.sheet_to_json(sheet, { header: 1, });
    const filteredMappingData = mappingData.filter(row => row[3] !== 0); // Берём из файла mapping только те строки, в которых в 4 столбце не указано 0

    // Формируем остатки
    const result = filteredMappingData.filter(value => value[1]).map((article, i) => {
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
      case 'ad': // Сохраняем остатки для ArtDesign
        data = await parseArtdesignStocks();

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

        createXLSXFile(resultOzon, AD_OZON_RESULT_FILE_PATH);
        createXLSXFile(resultWb, AD_WB_RESULT_FILE_PATH);

        break;
      case 'TDL': // Сохраняем остатки для TDL
        data = await parseTDLStocks();

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

        createXLSXFile(resultOzon, TDL_OZON_RESULT_FILE_PATH);
        createXLSXFile(resultWb, TDL_WB_RESULT_FILE_PATH);

        break;
      case 'logos': // Сохраняем остатки для Logos
        data = await parseLogosStocks();

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

        createXLSXFile(resultOzon, LOGOS_OZON_RESULT_FILE_PATH);
        createXLSXFile(resultWb, LOGOS_WB_RESULT_FILE_PATH);

        break;
      case 'TT': // Сохраняем остатки для Традиции Текстиля
        data = await parseTTStocks();

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

        createXLSXFile(resultOzon, TT_OZON_RESULT_FILE_PATH);
        createXLSXFile(resultWb, TT_WB_RESULT_FILE_PATH);

        break;
      default:
        console.log('Unknown argument');
    }
  } catch (error) {
    console.error(error);
  }
};

main();