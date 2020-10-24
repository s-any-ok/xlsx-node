const xlsx = require("sheetjs-style");
const {
  sortByTeamName,
  setColmsLen,
  getAllTeams,
  appendAllSheets,
} = require("./helpers/functions");

const fileName = "CS_GO5Х5.xlsx";
const options = { cellDates: true };

const wb = xlsx.readFile(fileName, options);
const ws = wb.Sheets["Ответы на форму (1)"];

const jsonData = xlsx.utils.sheet_to_json(ws, { raw: false });
const sortJsonData = sortByTeamName(jsonData);

const allTeams = getAllTeams(sortJsonData);
const newWB = xlsx.utils.book_new();
appendAllSheets(allTeams, newWB);

xlsx.writeFile(newWB, "Members Table CSGO.xlsx");

console.log(allTeams);

// const newWS1 = xlsx.utils.json_to_sheet(team1);
// newWS1["!cols"] = setColmsLen();
// xlsx.utils.book_append_sheet(newWB, newWS1, "Night Raid");
