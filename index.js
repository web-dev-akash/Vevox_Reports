const express = require("express");
const cors = require("cors");
const axios = require("axios");
const { google } = require("googleapis");
const path = require("path");
const xlsx = require("xlsx");
const multer = require("multer");
const fs = require("fs");
const { promisify } = require("util");
const unlinkAsync = promisify(fs.unlink);
require("dotenv").config();
const app = express();
app.use(express.urlencoded({ extended: true }));
app.use(cors());
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const REFRESH_TOKEN = process.env.REFRESH_TOKEN;
const PORT = process.env.PORT || 8080;

const storage = multer.diskStorage({
  destination: path.join(__dirname, "uploads"),
  filename: function (req, file, cb) {
    cb(null, file.fieldname);
  },
});
const upload = multer({ storage: storage });

app.use(express.static(path.join(__dirname, "template")));

app.get("/", (req, res) => {
  res.sendFile(`index.html`);
});

const getVevoxSheetData = async () => {
  const spreadsheetId = process.env.SPREADSHEET_ID;
  const auth = new google.auth.GoogleAuth({
    keyFile: "key.json", //the key file
    scopes: "https://www.googleapis.com/auth/spreadsheets",
  });
  const authClientObject = await auth.getClient();
  const sheet = google.sheets({
    version: "v4",
    auth: authClientObject,
  });
  const readData = await sheet.spreadsheets.values.get({
    auth, //auth object
    spreadsheetId, // spreadsheet id
    range: "Vevox Data!A:H", //range of cells to read from.
  });
  return readData.data.values;
};

const addDataToSheet = async (users) => {
  const spreadsheetId = process.env.SPREADSHEET_ID;
  const auth = new google.auth.GoogleAuth({
    keyFile: "key.json", //the key file
    scopes: "https://www.googleapis.com/auth/spreadsheets",
  });
  const authClientObject = await auth.getClient();
  const sheet = google.sheets({
    version: "v4",
    auth: authClientObject,
  });
  const rowsLength = await getVevoxSheetData();
  const writeData = await sheet.spreadsheets.values.update({
    auth, //auth object
    spreadsheetId, //spreadsheet id
    range: `Vevox Data!A${rowsLength.length + 1}:H${
      rowsLength.length + users.length
    }`, //sheet name and range of cells
    valueInputOption: "USER_ENTERED",
    resource: {
      values: users,
    },
  });
  return writeData.data;
};

const updateDataOnVevoxSheet = async (users) => {
  const vevoxData = await getVevoxSheetData();
  const addVevoxData = [];
  for (let i = 0; i < users.length; i++) {
    const foundExistingEntry = vevoxData.filter(
      (user) => user[3] === users[i].email && +user[5] === +users[i].sessionId
    );
    if (foundExistingEntry.length == 0) {
      addVevoxData.push([
        users[i].firstname,
        users[i].lastname,
        users[i].correct,
        users[i].email,
        users[i].date,
        users[i].sessionId,
        users[i].attempted,
        users[i].polled,
      ]);
    }
  }
  await addDataToSheet(addVevoxData);
  return addVevoxData;
};

app.post("/view", upload.single("file.xlsx"), async (req, res) => {
  try {
    const workbook = xlsx.readFile("./uploads/file.xlsx");
    const sheetName1 = workbook.SheetNames[0];
    const sheet1 = workbook.Sheets[sheetName1];
    const data1 = xlsx.utils.sheet_to_json(sheet1);
    const sessionId = data1[4]["__EMPTY"];
    const currentUsers = [];
    for (let i = 8; i < data1.length; i++) {
      const firstname = data1[i][""];
      const lastname = data1[i]["__EMPTY"];
      const email = data1[i]["__EMPTY_1"];
      const attemptDate = new Date(
        data1[i]["__EMPTY_2"].substring(0, 11)
      ).toDateString();
      const obj = { firstname, lastname, email, attemptDate };
      if (email && !email.includes("1234500")) {
        currentUsers.push(obj);
      }
    }

    const finalUsers = [];
    const sheetName2 = workbook.SheetNames[2];
    const sheet2 = workbook.Sheets[sheetName2];
    const data2 = xlsx.utils.sheet_to_json(sheet2);
    const totalPolled = Object.values(data2[2]).length - 3;
    for (let i = 7; i < data2.length - 2; i++) {
      const firstname = data2[i]["Polling Results"];
      const lastname = data2[i]["__EMPTY"];
      const correct = data2[i]["__EMPTY_1"] ? data2[i]["__EMPTY_1"] : 0;
      const totalNotEmpty = Object.values(data2[i]).filter(
        (val) => val === ""
      ).length;
      console.log(totalNotEmpty);
      const totalAttempted =
        totalNotEmpty === 0 ? 0 : totalPolled - totalNotEmpty;
      const userFound = currentUsers.find(
        (user) => user.firstname === firstname && user.lastname === lastname
      );
      if (userFound) {
        const obj = {
          firstname,
          lastname,
          correct,
          email: userFound.email,
          date: userFound.attemptDate,
          polled: totalPolled,
          attempted: totalAttempted,
          sessionId,
        };
        finalUsers.push(obj);
      }
    }
    // return res.send({ data1, data2, finalUsers });
    const data = await updateDataOnVevoxSheet(finalUsers);
    let table = "";
    data.map(
      (user) =>
        (table += `<tr>
          <td style="border:1px solid; padding:10px 20px;">${user[0]}</td>
          <td style="border:1px solid; padding:10px 20px;">${user[1]}</td>
          <td style="border:1px solid; padding:10px 20px;">${user[2]}</td>
          <td style="border:1px solid; padding:10px 20px;">${user[3]}</td>
          <td style="border:1px solid; padding:10px 20px;">${user[4]}</td>
          <td style="border:1px solid; padding:10px 20px;">${user[5]}</td>
          <td style="border:1px solid; padding:10px 20px;">${user[6]}</td>
          <td style="border:1px solid; padding:10px 20px;">${user[7]}</td>
        </tr>`)
    );
    res.send(`
    <div style="width : 80%; margin : 50px auto; text-align : center; display : grid; place-items:center;">
      <h1>Excel file uploaded and processed successfully.</h1>
      <Table style="text-align : center; font-size : 20px; margin-top : 20px; border-collapse: collapse; ">
        <Thead>
          <th style="border:1px solid; padding:10px 20px;">First Name</th>
          <th style="border:1px solid; padding:10px 20px;">Last Name</th>
          <th style="border:1px solid; padding:10px 20px;">Correct Answer</th>
          <th style="border:1px solid; padding:10px 20px;">Email</th>
          <th style="border:1px solid; padding:10px 20px;">Date</th>
          <th style="border:1px solid; padding:10px 20px;">Session ID</th>
          <th style="border:1px solid; padding:10px 20px;">Attempted</th>
          <th style="border:1px solid; padding:10px 20px;">Polled</th>
        </Thead>
        ${table}
      </Table>
    </div>
    `);
    await unlinkAsync(req.file.path);
    return;
  } catch (error) {
    console.error("Error reading Excel file:", error);
    res.status(500).send({ error: "Error reading Excel file." });
    await unlinkAsync(req.file.path);
    return;
  }
});

app.listen(PORT, () => {
  console.log(`http://localhost:${PORT}`);
});
