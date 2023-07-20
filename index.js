const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');
const readline = require('readline');
const moment = require('moment');

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});

(async () => {
  // Prompt for the URL
  rl.question('Enter the URL to scrape: ', async (url) => {
    const browser = await puppeteer.launch({ headless: false });
    const page = await browser.newPage();

    await page.goto(url);

    // Prompt for the price range
    rl.question('Enter the minimum price: ', async (minPrice) => {
      rl.question('Enter the maximum price: ', async (maxPrice) => {
        let combinedData = [];

        while (true) {
          // Получаем названия товаров
          const titles = await page.$$eval('.styles-module-root-TWVKW', elements => elements.map(el => el.innerText));

          // Получаем цены товаров
          const prices = await page.$$eval('.styles-module-root-_KFFt.styles-module-size_l-_oGDF.styles-module-size_l_dense-Wae_G.styles-module-size_l-hruVE.styles-module-size_dense-z56yO.stylesMarningNormal-module-root-OSCNq.stylesMarningNormal-module-paragraph-l-dense-TTLmp', elements => elements.map(el => el.innerText));

          // Получаем ссылки на товары
          const links = await page.$$eval('a[data-marker="item-title"]', elements => elements.map(el => el.getAttribute('href')));

          // Получаем описания товаров
          const descriptions = await page.$$eval('p.styles-module-root-_KFFt[style="-webkit-line-clamp:4"]', elements => elements.map(el => el.innerText));

          // Получаем города
          const cities = await page.$$eval('div.geo-root-zPwRk', elements => elements.map(el => el.innerText));

          // Получаем время публикации
          const dates = await page.$$eval('p[data-marker="item-date"]', elements => elements.map(el => el.innerText));

          const pageData = titles.map((title, index) => ({
            title,
            price: prices[index],
            link: 'https://www.avito.ru' + links[index],
            description: descriptions[index],
            city: cities[index],
            date: dates[index],
          }));

          // Filter data based on the price range
          const filteredData = pageData.filter(item => {
            const priceValue = parseFloat(item.price.replace(/\s+/g, '').replace(/,/g, '.')); // Convert the price to a number
            return priceValue >= minPrice && priceValue <= maxPrice;
          });

          combinedData = combinedData.concat(filteredData);

          // Вызываем JavaScript-обработчик клика на кнопку "Дальше"
          const nextButton = await page.$('.styles-module-item_arrow-sxBqe[aria-label="Следующая страница"]');
          if (nextButton) {
            await page.evaluate(element => element.click(), nextButton);
            await page.waitForTimeout(2000); // Добавляем задержку 2 секунды перед парсингом следующей страницы
          } else {
            break;
          }
        }

        // Prompt for the output filename
        rl.question('Enter the output filename (e.g., output.xlsx): ', async (filename) => {
          // Ensure the filename has the .xlsx extension
          if (!filename.endsWith('.xlsx')) {
            filename += '.xlsx';
          }

          // Создаем новый документ Excel
          const workbook = new ExcelJS.Workbook();
          const worksheet = workbook.addWorksheet(moment().format('DD.MM.YYYY')); // Use the current date as the worksheet name

          // Заголовки столбцов
          worksheet.addRow(['Название', 'Цена', 'Ссылка', 'Описание', 'Город', 'Время публикации']);

          // Записываем данные в таблицу Excel
          combinedData.forEach(({ title, price, link, description, city, date }) => {
            worksheet.addRow([title, price, link, description, city, date]);
          });

          // Автоматически изменяем ширину столбцов для лучшего отображения данных
          worksheet.columns.forEach(column => {
            if (column.header) {
              column.width = Math.max(15, column.header.length, ...column.values.map(value => (value ? value.toString().length : 0)));
            } else {
              column.width = 15;
            }
          });

          // Сохраняем документ Excel в файл с указанным именем
          await workbook.xlsx.writeFile(filename);

          console.log(`Данные успешно записаны в файл ${filename}.`);

          await browser.close();
          rl.close(); // Close the readline interface
        });
      });
    });
  });
})();
