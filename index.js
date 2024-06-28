// index.js
const axios = require("axios");
const XLSX = require("xlsx");
const fs = require("fs");
// require('dotenv').config();

const pageNumber = 1;
const pageSize = 25;
// const letter = "o"

const API_URL = `https://api.messefrankfurt.com/service/esb_api/exhibitor-service/api/2.1/public/exhibitor/search?language=en-GB&q=&orderBy=name&pageNumber=${pageNumber}&pageSize=${pageSize}&orSearchFallback=false&showJumpLabels=false&jumpLabelId=o&findEventVariable=HEIMTEXTIL`;

const fetchData = async () => {
  try {
    const response = await axios.get(API_URL, {
      headers: {
        // Add necessary headers here if needed
        Apikey: "LXnMWcYQhipLAS7rImEzmZ3CkrU033FMha9cwVSngG4vbufTsAOCQQ==",
      },
    });
    // console.log(response.data.result.hits);
    return response.data.result.hits;
  } catch (error) {
    console.error("Error fetching data:", error);
    return null;
  }
};

const removeHtmlTags = (str) => {
  return str.replace(/<br \/>/g, " ").trim();
};

const removeDuplicates = (data, key) => {
  const uniqueData = [];
  const emailSet = new Set();

  data.forEach((item) => {
    if (!emailSet.has(item[key])) {
      emailSet.add(item[key]);
      uniqueData.push(item);
    }
  });

  return uniqueData;
};

const filterData = (data) => {
  let filteredData = [];

  data.forEach((item) => {
    const exhibitor = item.exhibitor;
    const exhibitionHall = exhibitor.exhibition.exhibitionHall.length
      ? exhibitor.exhibition.exhibitionHall[0]
      : {};
    const stand = exhibitionHall.stand.length ? exhibitionHall.stand[0] : {};

    if (exhibitor.contacts && exhibitor.contacts.length > 0) {
      exhibitor.contacts.forEach((contact) => {
        filteredData.push({
          companyName: exhibitor.name,
          contactPerson: `${contact.firstname} ${contact.lastname}`,
          designation: contact.position,
          country: exhibitor.address.country.label,
          email: contact.email,
          website: exhibitor.homepage,
          telephone: contact.phone,
          mobile: contact.mobile,
          address: removeHtmlTags(exhibitor.addressrdm.formatedAddress),
          hall: exhibitionHall.id,
          booth: stand.name,
        });
      });
    } else {
      filteredData.push({
        companyName: exhibitor.name,
        contactPerson: "",
        designation: "",
        country: exhibitor.address.country.label,
        email: exhibitor.address.email,
        website: exhibitor.homepage,
        telephone: exhibitor.address.tel,
        mobile: "",
        address: removeHtmlTags(exhibitor.addressrdm.formatedAddress),
        hall: exhibitionHall.id,
        booth: stand.name,
      });
    }
  });

  const uniqueData = removeDuplicates(filteredData, "email");
  //   console.log(uniqueData);
  return uniqueData;
};

// const saveToExcel = (data) => {
//   const worksheet = XLSX.utils.json_to_sheet(data);
//   const workbook = XLSX.utils.book_new();
//   XLSX.utils.book_append_sheet(workbook, worksheet, "Exhibitors");
//   XLSX.writeFile(workbook, "exhibitors.xlsx");
// };

const appendToExcel = (data, filePath) => {
  let workbook;
  const sheetName = "Exhibitors";
  let worksheet;

  if (fs.existsSync(filePath)) {
    // If the file exists, read the existing workbook
    workbook = XLSX.readFile(filePath);
    worksheet = workbook.Sheets[sheetName];
    let existingData = [];

    if (worksheet) {
      // If the sheet exists, append new data to existing data
      const existingData = XLSX.utils.sheet_to_json(worksheet);
    }
    data = [...existingData, ...data];
  } else {
    // If the file doesn't exist, create a new workbook and sheet
    workbook = XLSX.utils.book_new();
  }

  // Create a new worksheet with the combined data
  worksheet = XLSX.utils.json_to_sheet(data, {
    header: [
      "companyName",
      "contactPerson",
      "designation",
      "country",
      "email",
      "website",
      "telephone",
      "mobile",
      "address",
      "hall",
      "booth",
    ],
  });

  // Append the worksheet to the workbook
  //   XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
  if (workbook.SheetNames.includes(sheetName)) {
    workbook.Sheets[sheetName] = worksheet;
  } else {
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
  }

  // Write the workbook to the file
  XLSX.writeFile(workbook, filePath);
};

const saveToExcel = (data) => {
  const worksheet = XLSX.utils.json_to_sheet(data, {
    header: [
      "companyName",
      "contactPerson",
      "designation",
      "country",
      "email",
      "website",
      "telephone",
      "mobile",
      "address",
      "hall",
      "booth",
    ],
  });
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Exhibitors");
  XLSX.writeFile(workbook, "exhibitors1.xlsx");
};

const main = async () => {
  const data = await fetchData();
  if (data) {
    const filteredData = filterData(data);
    // console.log(filteredData);

    // saveToExcel(filteredData);
    appendToExcel(filteredData, "exhibitors.xlsx");
    console.log("Data saved to exhibitors.xlsx");
  }
};

main();
