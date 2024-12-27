const fs = require("fs");
const readline = require("readline");
const { google } = require("googleapis");
const xlsx = require("xlsx");
// const http = require("http");

// Load client secrets from a local file
const port = 3000; // Change this if you need another port
const CREDENTIALS_PATH = "credentials.json";
const TOKEN_PATH = "token.json";
const EXCEL_FILE = "schedule.xlsx";
let year;
const YOUR_NAME = "Юсеф Карим"; // Replace with your name in the schedule
const monthNames = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];
const monthNumber = {
  January: 1,
  February: 2,
  March: 3,
  April: 4,
  May: 5,
  June: 6,
  July: 7,
  August: 8,
  September: 9,
  October: 10,
  November: 11,
  December: 12,
};

// Function to authenticate with Google Calendar API
function authorize(credentials, callback) {
  const { client_secret, client_id, redirect_uris } = credentials.installed;
  const oAuth2Client = new google.auth.OAuth2(
    client_id,
    client_secret,
    redirect_uris[0]
  );

  // Check if token already exists
  fs.readFile(TOKEN_PATH, (err, token) => {
    if (err) return getNewToken(oAuth2Client, callback);
    oAuth2Client.setCredentials(JSON.parse(token));
    callback(oAuth2Client);
  });
}

// Get new token if necessary
function getNewToken(oAuth2Client, callback) {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: "offline",
    scope: ["https://www.googleapis.com/auth/calendar"],
  });
  console.log("Authorize this app by visiting this URL:", authUrl);
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
  rl.question("Enter the code from that page here: ", (code) => {
    rl.close();
    oAuth2Client.getToken(code, (err, token) => {
      if (err) return console.error("Error retrieving access token", err);
      oAuth2Client.setCredentials(token);
      fs.writeFileSync(TOKEN_PATH, JSON.stringify(token));
      console.log("Token stored to", TOKEN_PATH);
      callback(oAuth2Client);
    });
  });
}

// Read Excel file and extract your schedule
function parseSchedule() {
  const workbook = xlsx.readFile(EXCEL_FILE);
  const sheetNames = workbook.SheetNames;
  const foundMonth = sheetNames.find((sheetName) =>
    monthNames.some((month) =>
      sheetName.toLowerCase().includes(month.toLowerCase())
    )
  );
  const sheet = workbook.Sheets[foundMonth];
  const data = xlsx.utils.sheet_to_json(sheet);

  function findObjectByValue(data, targetValue) {
    return data.find((obj) =>
      Object.entries(obj).some(([key, value]) => {
        if (value === targetValue) {
          year = extractYear(key);
          return true;
        }
      })
    );
  }

  function extractYear(title) {
    const yearMatch = title.match(/\b\d{4}\b/); // Match any 4-digit number
    return yearMatch ? parseInt(yearMatch[0], 10) : null; // Convert to number or return null
  }

  const foundObject = findObjectByValue(data, YOUR_NAME);
  if (foundObject) {
    console.log("Found object:", foundObject);
  } else {
    console.log(`No object found with value: ${targetValue}`);
  }

  function extractNumber(str) {
    const match = str.match(/\d+/); // Matches one or more digits
    return match ? parseInt(match[0], 10) : null; // Convert to a number or return null
  }

  const formatter = new Intl.NumberFormat("en-US", {
    minimumIntegerDigits: 2,
    useGrouping: false,
  });

  function extractSchedule(data) {
    const schedule = [];
    let day = 1; // Start from the 1st of December

    for (const key in data) {
      const keyNumber = extractNumber(key);
      const timePattern = /^\d{2}:\d{2}-\d{2}:\d{2}$/; // Matches 'HH:mm-HH:mm'
      if (key.startsWith("__EMPTY_") && !isNaN(keyNumber) && keyNumber >= 5) {
        if (typeof data[key] === "string" && timePattern.test(data[key]))
          schedule.push({
            date: `${year}-${monthNumber[foundMonth]}-${formatter.format(day)}`,
            startTime: data[key].split("-")[0],
            endTime: data[key].split("-")[1],
          });
        day++;
      }
    }

    return schedule;
  }
  return extractSchedule(foundObject);
}

// Add events to Google Calendar
function addEvents(auth) {
  const calendar = google.calendar({ version: "v3", auth });

  const events = parseSchedule();
  console.log(events);
  events.forEach((event) => {
    const calendarEvent = {
      summary: "life",
      colorId: "11",
      start: {
        dateTime: `${event.date}T${event.startTime}:00`,
        timeZone: "Europe/Minsk",
      },
      end: {
        dateTime: `${event.date}T${event.endTime}:00`,
        timeZone: "Europe/Minsk",
      },
    };

    calendar.events.insert(
      {
        calendarId: "primary",
        resource: calendarEvent,
      },
      (err, event) => {
        if (err) return console.error("Error adding event:", err);
        console.log("Event created:", event.data.htmlLink);
      }
    );
  });
}

// Main function
fs.readFile(CREDENTIALS_PATH, (err, content) => {
  if (err) return console.error("Error loading client secret file:", err);
  authorize(JSON.parse(content), addEvents);
});
