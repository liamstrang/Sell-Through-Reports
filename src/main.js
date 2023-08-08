const axios = require("axios");
const readline = require("readline");
const ExcelJS = require("exceljs");
const ProgressBar = require("progress");
const fs = require("fs");
require("dotenv").config();

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});

const API_HASH = process.env.API_HASH;
const API_TOKEN = process.env.API_TOKEN;

// Fetch order data from BigCommerce API
async function fetchOrders(dateRangeStart, dateRangeEnd) {
  // Set up the API headers
  const headers = {
    "X-Auth-Token": API_TOKEN,
    "Content-Type": "application/json",
  };
  let orders = [];
  let page = 1;
  const limit = 200; // Max limit per page allowed by BigCommerce V2 API

  while (true) {
    const response = await axios.get(
      `https://api.bigcommerce.com/stores/${API_HASH}/v2/orders?status_id=2&min_date_created=${dateRangeStart}&max_date_created=${dateRangeEnd}&limit=${limit}&page=${page}`,
      { headers }
    );

    // If there's no data or less data than the limit, break out of the loop
    if (!response.data.length) break;

    orders = orders.concat(response.data);

    // If we've fetched less than the limit, it means we're on the last page
    if (response.data.length < limit) break;

    page++;
  }

  return orders;
}

async function fetchOrderProducts(order) {
  const headers = {
    "X-Auth-Token": API_TOKEN,
    "Content-Type": "application/json",
  };

  const response = await axios.get(order.products.url, { headers });
  return response.data;
}

// Generate the sell-through report
async function generateReport(brand, dateRangeStart, dateRangeEnd) {
  const orders = await fetchOrders(dateRangeStart, dateRangeEnd);

  const bar = new ProgressBar("Generating report [:bar] :percent :etas", {
    complete: "=",
    incomplete: " ",
    width: 40,
    total: orders.length,
  });

  // Create a new workbook and worksheet
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Sell Through Report");

  // Define some background colors
  const headerBgColor = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFC0C0C0" },
  };
  const dataBgColor = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFFFF" },
  };

  // Define and set column headers
  worksheet.columns = [
    { header: "Brand", key: "Brand", width: 15 },
    { header: "SKU", key: "SKU", width: 20 },
    { header: "Title", key: "Title", width: 30 },
    { header: "Quantity", key: "Quantity", width: 10 },
  ];

  // Style the headers
  worksheet.getRow(1).eachCell((cell) => {
    cell.fill = headerBgColor;
  });

  let rowIndex = 2; // Start from the second row as the first row is for headers

  const ordersProducts = await Promise.all(
    orders.map((order) => fetchOrderProducts(order))
  );

  let productAggregates = {};

  for (let i = 0; i < orders.length; i++) {
    const products = ordersProducts[i];

    // Filter products based on the brand (if provided)
    const filteredProducts = brand
      ? products.filter((product) => product.brand === brand)
      : products;

    for (let product of filteredProducts) {
      // Use the product's SKU as a key to aggregate data
      if (!productAggregates[product.sku]) {
        productAggregates[product.sku] = {
          Brand: product.brand,
          SKU: product.sku,
          Title: product.name,
          Quantity: 0, // Initialize with zero
        };
      }

      // Increment the quantity
      productAggregates[product.sku].Quantity += product.quantity;
    }
    bar.tick();
  }
  const productArray = Object.values(productAggregates);

  // Sort the array based on Quantity in descending order
  productArray.sort((a, b) => b.Quantity - a.Quantity);

  // Now, write the aggregated data to the worksheet
  for (const productData of productArray) {
    worksheet.addRow({
      Brand: productData.Brand,
      SKU: productData.SKU,
      Title: productData.Title,
      Quantity: productData.Quantity,
    });

    // Style the data rows
    worksheet.getRow(rowIndex).eachCell((cell) => {
      cell.fill = dataBgColor;
    });
    rowIndex++;
  }

  // Save the workbook to a file
  await workbook.xlsx.writeFile("SellThroughReport.xlsx");
  console.log("Report generated as 'SellThroughReport.xlsx'");
}

// Get user input for brand and date range
rl.question("Enter the brand (or press enter to skip): ", (brand) => {
  rl.question(
    "Enter the date range (past week, past month, past quarter, past year, custom): ",
    async (dateRange) => {
      let dateRangeStart = new Date();
      let dateRangeEnd = new Date();

      switch (dateRange) {
        case "past week":
          dateRangeStart.setDate(dateRangeEnd.getDate() - 7);
          break;
        case "past month":
          dateRangeStart.setMonth(dateRangeEnd.getMonth() - 1);
          break;
        case "past quarter":
          dateRangeStart.setMonth(dateRangeEnd.getMonth() - 3);
          break;
        case "past year":
          dateRangeStart.setFullYear(dateRangeEnd.getFullYear() - 1);
          break;
        case "custom":
          rl.question(
            "Enter the start date (DD/MM/YYYY): ",
            (startDateInput) => {
              rl.question(
                "Enter the end date (DD/MM/YYYY): ",
                async (endDateInput) => {
                  dateRangeStart = parseCustomDate(startDateInput);
                  dateRangeEnd = parseCustomDate(endDateInput);

                  if (!dateRangeStart || !dateRangeEnd) {
                    console.log("Invalid date format.");
                    rl.close();
                    return;
                  }

                  if (fs.existsSync("SellThroughReport.xlsx")) {
                    fs.unlinkSync("SellThroughReport.xlsx");
                    console.log("Existing report deleted.");
                  }

                  await generateReport(
                    brand,
                    dateRangeStart.toISOString(),
                    dateRangeEnd.toISOString()
                  );

                  rl.close();
                }
              );
            }
          );
          return;

        default:
          console.log("Invalid date range.");
          rl.close();
          return;
      }

      function parseCustomDate(dateString) {
        const [day, month, year] = dateString.split("/");
        if (day && month && year) {
          return new Date(`${year}-${month}-${day}T00:00:00Z`);
        }
        return null;
      }
    }
  );
});
