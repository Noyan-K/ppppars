const exceljs = require('exceljs');
const path = require('path');
const { PrismaClient } = require("@prisma/client");
const prisma = new PrismaClient()

async function parser() {
  const workBook = new exceljs.Workbook();
  const file = await workBook.xlsx.readFile(path.join(__dirname, "../", "vendor", "xerox.xlsx"));
  const sheet = file.getWorksheet('Documate Scanners');

  const scannerCategory = await prisma.category.findFirst({ where: { name: { startsWith: 'Scanner', mode: 'insensitive' } } }) ?? await prisma.category.create({ data: { name: 'Scanner' } });
  const xeroxVendor = await prisma.vendor.findFirst({ where: { name: { startsWith: 'Xerox', mode: 'insensitive' } } }) ?? await prisma.vendor.create({ data: { name: 'Xerox' } });
  const currency = await prisma.currency.findFirst({ where: { name: { startsWith: 'usd', mode: 'insensitive' } } }) ?? await prisma.currency.create({ data: { name: 'usd' } });
  const supplier = await prisma.supplier.findFirst({ where: { name: { startsWith: 'Xerox partner only', mode: 'insensitive' } } }) ?? await prisma.supplier.create({ data: { name: 'Xerox partner only' } }); // !!!!!!!! RENAME SUPPLIER


  let childCategory;

  sheet.eachRow({ includeEmpty: true }, async function(row, rowNumber) {
    if (typeof row.values[2] === 'string' && row.values[3] === undefined) {
      childCategory = await prisma.category.create({ data: { name: row.values[2], parent_id: scannerCategory.id } });
    }

    if (typeof row.values[12] === 'number' && typeof row.values[14] === 'number') {
      const product = await prisma.product.create({ data: {
        name: row.values[3],
        category_id: childCategory?.id ?? scannerCategory.id,
        vendor_partnumber: row.values[6],
        vendor_id: xeroxVendor.id
      } });

      await prisma.supplierProductPrice.create({ data: {
        price: row.values[12],
        price_date: new Date(),
        product_id: product.id,
        currency_id: currency.id,
        currency_id: currency.id,
        supplier_id: supplier.id,
      } });
    }
  });
}

async function parser1() {
  const workBook = new exceljs.Workbook();
  const file = await workBook.xlsx.readFile(path.join(__dirname, "../", "vendor", "xerox.xlsx"));
  const sheet = file.getWorksheet('Mono printers');

  for (let i = 3; i <= sheet.rowCount; i++) {
    let row = sheet.getRow(i);

    if (
      row.getCell(2).style.font.size === 12 &&
      row.getCell(2).style.fill.fgColor?.indexed === 12
    ) {
      const productName = row.values[2];

      for (let j = i + 1; j <= sheet.rowCount; j++) {
        if (sheet.getRow(j).values === []) {
          break;
        }

        if (
          sheet.getRow(j).values[2] === 'MANDATORY' &&
          sheet.getRow(j).getCell(2).style.fill.fgColor?.indexed === 11
        ) {
          for (let k = j + 1; k <= sheet.rowCount; k++) {
            if (sheet.getRow(k).getCell(2).style.fill.fgColor?.indexed === 11) {
              break;
            }

            if (sheet.getRow(k).getCell(2).style.fill.fgColor?.indexed === 42) {
            }
          }
        }

        if (
          sheet.getRow(j).getCell(2).style.font.size === 10 &&
          sheet.getRow(j).getCell(2).style.fill.fgColor?.indexed === 12
        ) {
          for (let k = j + 1; k <= sheet.rowCount; k++) {
            if (sheet.getRow(k).values.length === 0) {
              break;
            }

            const product = sheet.getRow(k).values;
            console.log(product[3], '======');
          }
        }
      }
    }
  }
}

async function parser2(sheet) {

  const category = await prisma.category.findFirst({ where: { name: { startsWith: sheet.name, mode: 'insensitive' } } }) ?? await prisma.category.create({ data: { name: sheet.name } });
  const vendor = await prisma.vendor.findFirst({ where: { name: { startsWith: 'Xerox', mode: 'insensitive' } } }) ?? await prisma.vendor.create({ data: { name: 'Xerox' } });
  const language = await prisma.language.findFirst({ where: { language: { startsWith: 'en', mode: 'insensitive' } } }) ?? await prisma.language.create({ data: { language: 'en' }});
  const currency = await prisma.currency.findFirst({ where: { name: { startsWith: 'usd', mode: 'insensitive' } } }) ?? await prisma.currency.create({ data: { name: 'usd' } });
  const supplier = await prisma.supplier.findFirst({ where: { name: { startsWith: 'Xerox partner only', mode: 'insensitive' } } }) ?? await prisma.supplier.create({ data: { name: 'Xerox partner only' } });

  scanProducts: for (let i = 3; i <= sheet.rowCount; i++) {
    let row = sheet.getRow(i);

    if (
      row.getCell(2).style.font?.size === 12 &&
      (row.getCell(2).style.fill.fgColor?.indexed === 12 ||
      row.getCell(2).style.fill.fgColor?.argb === 'FFFF0000')
    ) {
      const productName = row.values[2].includes('- STOP ORDER') || row.values[2].includes('- NO LONGER') ? row.values[2].slice(0, row.values[2].indexOf('- STOP ORDER')).slice(0, row.values[2].indexOf('- NO LONGER')) : row.values[2];

      scanProductsInner: for (let j = i + 1; j <= sheet.rowCount; j++) {
        let row = sheet.getRow(j);

        if (
          row.values[2] === 'MANDATORY' &&
          row.getCell(2).style.fill.fgColor?.indexed === 11
        ) {
          scanProductsMandatory: for (let m = j + 1; m <= sheet.rowCount; m++) {
            let row = sheet.getRow(m);

            if (
              row.getCell(2).style.fill.fgColor?.indexed === 11 ||
              row.getCell(2).style.font.size !== 10
            ) {
              j = m;
              continue scanProductsInner;
            }

            if (typeof row.values[11] === 'number') {
              const description = await prisma.description.create({
                data: {
                  text: typeof row.values[3] === 'string' ? row.values[3] : row.values[3].richText[0].text + row.values[3].richText[1].text,
                  language_id: language.id,
                },
              });

              const product = await prisma.product.create({
                data: {
                  name: productName,
                  category_id: category.id,
                  vendor_partnumber: row.values[5],
                  vendor_id: vendor.id,
                  description_id: description?.id ?? null,
                },
              });

              await prisma.supplierProductPrice.create({ data: {
                price: row.values[11],
                price_date: new Date(),
                product_id: product.id,
                currency_id: currency.id,
                supplier_id: supplier.id,
              } });
            }
          }
        }

        if (
          row.getCell(2).style.font.size === 10 &&
          row.getCell(2).style.fill.fgColor?.indexed === 12
        ) {
          scanProductsOthers: for (let o = j + 1; o <= sheet.rowCount; o++) {
            let row = sheet.getRow(o);

            if (
              row.getCell(2).style.font.size === 12
            ) {
              i = o - 1;
              continue scanProducts;
            }

            if (
              typeof row.values[3] !== 'undefined' &&
              typeof row.getCell(2).style.fill.fgColor === 'undefined' &&
              typeof row.values[11] === 'number'
            ) {
              const descText = Array.isArray(row.values[3]?.richText) ? row.values[3].richText.reduce((acc, curr) => acc.text + curr.text) : row.values[3];

              const description = await prisma.description.create({
                data: {
                  text: descText,
                  language_id: language.id,
                },
              });

              const product = await prisma.product.create({
                data: {
                  name: productName,
                  category_id: category.id,
                  vendor_partnumber: row.values[5],
                  vendor_id: vendor.id,
                  description_id: description?.id ?? null,
                },
              });

              await prisma.supplierProductPrice.create({ data: {
                price: row.values[11],
                price_date: new Date(),
                product_id: product.id,
                currency_id: currency.id,
                supplier_id: supplier.id,
              } });
            }
          }
        }
      }
    }
  }
}


async function parser4(sheet) {
  const category = await prisma.category.findFirst({ where: { name: { startsWith: sheet.name, mode: 'insensitive' } } }) ?? await prisma.category.create({ data: { name: sheet.name } });
  const vendor = await prisma.vendor.findFirst({ where: { name: { startsWith: 'Xerox', mode: 'insensitive' } } }) ?? await prisma.vendor.create({ data: { name: 'Xerox' } });
  const currency = await prisma.currency.findFirst({ where: { name: { startsWith: 'usd', mode: 'insensitive' } } }) ?? await prisma.currency.create({ data: { name: 'usd' } });
  const supplier = await prisma.supplier.findFirst({ where: { name: { startsWith: 'Xerox partner only', mode: 'insensitive' } } }) ?? await prisma.supplier.create({ data: { name: 'Xerox partner only' } }); // !!!!!!!! RENAME SUPPLIER

  for (let i = 3; i <= sheet.rowCount; i++) {
    const row = sheet.getRow(i);

    if (
      row.values[2] !== undefined &&
      (
        row.getCell(2).style.font?.size === 10 ||
        row.getCell(2).style.font?.size === 11
      ) &&
      typeof row.values[12] === 'number'
    ) {
      const product = await prisma.product.create({
        data: {
          name: row.values[2],
          category_id: category.id,
          vendor_partnumber: row.values[3],
          vendor_id: vendor.id
        },
      });

      await prisma.supplierProductPrice.create({ data: {
        price: row.values[12],
        price_date: new Date(),
        product_id: product.id,
        currency_id: currency.id,
        supplier_id: supplier.id,
      } });
    }
  }
}

async function parser5(sheet) {
  const category = await prisma.category.findFirst({ where: { name: { startsWith: sheet.name, mode: 'insensitive' } } }) ?? await prisma.category.create({ data: { name: sheet.name } });
  const vendor = await prisma.vendor.findFirst({ where: { name: { startsWith: 'Xerox', mode: 'insensitive' } } }) ?? await prisma.vendor.create({ data: { name: 'Xerox' } });
  const currency = await prisma.currency.findFirst({ where: { name: { startsWith: 'usd', mode: 'insensitive' } } }) ?? await prisma.currency.create({ data: { name: 'usd' } });
  const supplier = await prisma.supplier.findFirst({ where: { name: { startsWith: 'Xerox partner only', mode: 'insensitive' } } }) ?? await prisma.supplier.create({ data: { name: 'Xerox partner only' } }); // !!!!!!!! RENAME SUPPLIER

  mainLoop: for (let i = 3; i <= sheet.rowCount; i++) {
    const row = sheet.getRow(i);

    if (
      (
        row.getCell(2).style?.fill?.fgColor?.indexed === 11 ||
        row.getCell(2).style?.fill?.fgColor?.argb === 'FF00FF00'
      ) &&
      (
        row.values[2] === 'MANDATORY' ||
        row.values[2]?.includes('MANDATORY - ')
      ) &&
      row.getCell(2).style?.fill?.fgColor?.argb !== 'FFFF0000'
    ) {
      const productName = sheet.name.includes('_') ? null : row.values[2].replace('MANDATORY - ', '');

      for (let j = i + 1; j <= sheet.rowCount; j++) {
        const row = sheet.getRow(j);

        if (
          row.getCell(2).style?.fill?.fgColor?.argb !== 'FFCCFFCC' &&
          row.getCell(2).style?.fill?.fgColor?.indexed !== 42 &&
          typeof row.values[11] !== 'number'
        ) {
          i = j;
          continue mainLoop;
        }

        const product = await prisma.product.create({
          data: {
            name: productName ??
                  row.values[3]?.richText?.[0]?.text ??
                  typeof row.values[3] === 'string' ? row.values[3] : row.values[3].richText[0].text + row.values[3].richText[1].text,
            category_id: category.id,
            vendor_partnumber: row.values[5],
            vendor_id: vendor.id
          },
        });
  
        await prisma.supplierProductPrice.create({ data: {
          price: row.values[11],
          price_date: new Date(),
          product_id: product.id,
          currency_id: currency.id,
          supplier_id: supplier.id,
        } });
      }
    }

    if (
      row.values[3] !== undefined &&
      row.values[2] !== undefined &&
      typeof row.values[11] === 'number' &&
      row.getCell(2).style.fill.fgColor?.argb !== 'FFFF0000'
    ) {
      const product = await prisma.product.create({
        data: {
          name: typeof row.values[3] === 'string' ? row.values[3] : row.values[3].richText[0].text + row.values[3].richText[1].text,
          category_id: category.id,
          vendor_partnumber: row.values[5],
          vendor_id: vendor.id
        },
      });

      await prisma.supplierProductPrice.create({ data: {
        price: row.values[11],
        price_date: new Date(),
        product_id: product.id,
        currency_id: currency.id,
        supplier_id: supplier.id,
      } });
    }
  }
}

async function main() {
  const workBook = new exceljs.Workbook();
  const file = await workBook.xlsx.readFile(path.join(__dirname, "../", "vendor", "xerox.xlsx"));

  for (let sheetId = 3; sheetId < 5; sheetId++) {
    parser2(file.worksheets[sheetId]);
  }

  for (let sheetId = 4; sheetId < file.worksheets.rowCount; sheetId++) {
    parser5(file.worksheets[sheetId]);
  }

  parser();
  parser4(file.getWorksheet('Office Software'));
}

main();