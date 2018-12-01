/**
 * Scrapper for sports.fr calendar pages
 * It extracts data from one calendar page to the last page of calendar
 * Extracted data are written in an Excel file (one sheet by page) *
 */

const request = require('request-promise');
const cheerio = require('cheerio');
const cheerioTableparser = require('cheerio-tableparser');
const Excel = require('exceljs');

const workbook = new Excel.Workbook();
const baseUrl = 'http://www.sports.fr';
const firstPage = baseUrl + '/nba/2019/journees/journee2018-10-16.html';
const tableClass = '.nwResultats';
const nextClass = '.nwBtn.next';
const xlsxName = 'nba-calendrier.xlsx';
const endLink = '/nba/2019/journees/journee.html';

/**
 * Extracts data from a page
 *
 * @param {string} page page URL
 * @returns {Promise<string>} resolved with next page URL, rejected on error
 */
async function getDataFromPage(page) {
  return new Promise((resolve, reject) => {
    try {
      // Get page to parse
      const html = await request(page);
      // Extract table
      cheerioTableparser(cheerio);
      const data = cheerio(tableClass, html).parsetable(false, false, true);
      // Create new sheet with name extracted from table title
      const worksheet = workbook.addWorksheet(
        data[0][0]
          .split('-')[1]
          .trim()
          .replace(/\//g, '-')
      );
      // Sheet columns configuration
      worksheet.columns = [
        { key: 'heure', header: 'Heure' },
        { key: 'domicile', header: 'Equipe à domicile' },
        { key: 'scoreDomicile', header: 'Score' },
        { key: 'scoreExterieur', header: 'Score' },
        { key: 'exterieur', header: "Equipe à l'extérieur" },
      ];
      // Loop through table lines to create sheet rows
      for (let i = 1; i < data[0].length; i++) {
        worksheet.addRow({
          heure: data[1][i],
          domicile: data[2][i],
          scoreDomicile: data[3][i].split('-')[0],
          scoreExterieur: data[3][i].split('-')[1],
          exterieur: data[4][i],
        });
      }
      // Set sheet to be visible
      worksheet.state = 'visible';
      // Extract next page URL with next button
      const next = cheerio(nextClass, html).attr('href');
      if (next === endLink) return resolve();
      // A next page exists, resolve with it to continue
      resolve(baseUrl + next);
    } catch (err) {
      reject(err);
    }
  });
}

/**
 * Extract data from site
 * It loops from the first page to the last page through link
 * between them
 *
 * @param {string} page page URL to scrap
 */
async function scrapp(page) {
  if (page) {
    try {
      // Get data and next page
      return scrapp(await getDataFromPage(page));
    } catch (err) {
      console.error(err);
    }
  }
  // No more page, write excel file
  else workbook.xlsx.writeFile(xlsxName).then(() => console.log('done'));
}

// Launch scrapper
scrapp(firstPage);
