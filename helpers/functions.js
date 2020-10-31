const xlsx = require("sheetjs-style");

const teamName = "Маєш команду? Якщо так, то напиши її назву";
const fullName = "Прізвище та ім'я ";
const colmSize = [18, 20, 20, 20, 16, 28, 16, 18, 12];
const sortByTeamName = (array) =>
  [...array].sort((a, b) =>
    a[teamName] < b[teamName]
      ? 1
      : a[teamName] === b[teamName]
      ? a[fullName] < b[fullName]
        ? 1
        : -1
      : -1
  );

const getTeamsNames = (array) => {
  const teams = [];
  array.map((t) => !teams.includes(t[teamName]) && teams.push(t[teamName]));
  return teams;
};

const getTeamByName = (array, name) =>
  array.filter((member) => member[teamName] === name);

const setColmsLen = (arrLen = colmSize) => {
  const res = [];
  arrLen.map((l) => res.push({ wch: l }));
  return res;
};

const getAllTeams = (array) => {
  const teamsNames = getTeamsNames(array);
  const allTeams = [];
  for (let name of teamsNames) {
    let team = getTeamByName(array, name);
    allTeams.push(team);
  }
  return allTeams;
};

const appendAllSheets = (allTeams, wb) => {
  allTeams.map((team) => {
    let sheet = xlsx.utils.json_to_sheet(team);
    sheet["!cols"] = setColmsLen();
    sheet["!cols"][0] = { hidden: true };
    sheet["!cols"][6] = { hidden: true };
    delete sheet["B1"].w;
    sheet["B1"].v = "ВНЗ";
    delete sheet["I1"].w;
    sheet["I1"].v = "Команда";
    delete sheet["F1"].w;
    sheet["F1"].v = "Steam";
    xlsx.utils.sheet_add_aoa(sheet, [["Внесок, грн", 0]], {
      origin: "B10",
    });
    xlsx.utils.sheet_add_aoa(sheet, [["Порушення", ""]], {
      origin: "B11",
    });
    sheet["B10"].s = {
      font: {
        sz: 14,
        color: {
          rgb: "FF000000",
        },
      },
      fill: {
        fgColor: { rgb: "FFEA4D4D" },
      },
    };
    sheet["B11"].s = {
      font: {
        sz: 14,
        color: {
          rgb: "FF000000",
        },
      },
      fill: {
        fgColor: { rgb: "FFFFFF00" },
      },
    };
    colorSheetRows(sheet, ["A", "B", "C", "D", "E", "F", "G", "H", "I"]);
    xlsx.utils.book_append_sheet(
      wb,
      sheet,
      team[0]["Маєш команду? Якщо так, то напиши її назву"]
    );
  });
};

const colorSheetRows = (sheet, array) =>
  array.forEach(
    (l) =>
      (sheet[l + "1"].s = {
        font: {
          sz: 13,
          color: {
            rgb: "FF000000",
          },
        },
        fill: {
          fgColor: { rgb: "FFC9DAF8" },
        },
      })
  );

module.exports = {
  sortByTeamName,
  getTeamsNames,
  getTeamByName,
  setColmsLen,
  getAllTeams,
  appendAllSheets,
};
