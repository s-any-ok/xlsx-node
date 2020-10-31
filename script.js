const xlsx = require("sheetjs-style");
const {
  sortByTeamName,
  setColmsLen,
  getAllTeams,
  appendAllSheets,
  getTeamsNames,
} = require("./helpers/functions");

const fileName = "tables/CS_GO5Х5new.xlsx";
const options = { cellDates: true };

const wb = xlsx.readFile(fileName, options);
const ws = wb.Sheets["Ответы на форму (1)"];

const jsonData = xlsx.utils.sheet_to_json(ws, { raw: false });
const sortJsonData = sortByTeamName(jsonData);

const allTeams = getAllTeams(sortJsonData);
const newWB = xlsx.utils.book_new();
appendAllSheets(allTeams, newWB);

xlsx.writeFile(newWB, "tables/Members Table CSGO.xlsx");
const teams = getTeamsNames(jsonData);
console.dir(teams);

// const newWS1 = xlsx.utils.json_to_sheet(team1);
// newWS1["!cols"] = setColmsLen();
// xlsx.utils.book_append_sheet(newWB, newWS1, "Night Raid");
