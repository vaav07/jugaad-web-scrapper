// index.js
const axios = require("axios");
const XLSX = require("xlsx");
// require('dotenv').config();

const pageNumber = 1;

const API_URL = `https://api.messefrankfurt.com/service/esb_api/exhibitor-service/api/2.1/public/exhibitor/search?language=en-GB&q=&orderBy=name&pageNumber=${1}&pageSize=10&orSearchFallback=false&showJumpLabels=false&jumpLabelId=n&findEventVariable=HEIMTEXTIL`;

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

// const filterData = (data) => {
//   return data.map((item) => {
//     const exhibitor = item.exhibitor;
//     const exhibitionHall = exhibitor.exhibition.exhibitionHall.length
//       ? exhibitor.exhibition.exhibitionHall[0]
//       : {};
//     const stand = exhibitionHall.stand.length ? exhibitionHall.stand[0] : {};
//     const contacts = exhibitor.contacts.map((contact) => ({
//       firstname: contact.firstname,
//       lastname: contact.lastname,
//       department: contact.department,
//       position: contact.position,
//       phone: contact.phone,
//       mobile: contact.mobile,
//       email: contact.email,
//     }));

//     return {
//       name: exhibitor.name,
//       country: exhibitor.address.country.label,
//       email: exhibitor.address.email,
//       href: exhibitor.href,
//       homepage: exhibitor.homepage,
//       tel: exhibitor.address.tel,
//       formatedAddress: exhibitor.addressrdm.formatedAddress,
//       contacts: contacts,
//       hallAndLevel: exhibitionHall.id,
//       firstBoothNumber: stand.name,
//     };
//   });
// };

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

  return filteredData;
};

// const saveToExcel = (data) => {
//   const worksheet = XLSX.utils.json_to_sheet(data);
//   const workbook = XLSX.utils.book_new();
//   XLSX.utils.book_append_sheet(workbook, worksheet, "Exhibitors");
//   XLSX.writeFile(workbook, "exhibitors.xlsx");
// };

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
  XLSX.writeFile(workbook, "exhibitors.xlsx");
};

const main = async () => {
  const data = await fetchData();
  if (data) {
    const filteredData = filterData(data);
    // console.log(filteredData);

    saveToExcel(filteredData);
    console.log("Data saved to exhibitors.xlsx");
  }
};

main();
