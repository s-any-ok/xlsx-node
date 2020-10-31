const xlsx = require("sheetjs-style");

const teamName = "Назва команди";
const fullName = "Прізвище та ім'я ";
const colmSize = [18, 20, 20, 20, 16, 35, 20, 22, 12];
const headerLettersArr = ["A", "B", "C", "D", "E", "F", "G", "H", "I"];
const borders = {
  border: {
    top: { style: "thin", color: "FF000000" },
    bottom: { style: "thin", color: "FF000000" },
    left: { style: "thin", color: "FF000000" },
    right: { style: "thin", color: "FF000000" },
  },
};
const headerStyle = {
  font: {
    sz: 13,
    color: {
      rgb: "FF000000",
    },
  },
  fill: {
    fgColor: { rgb: "FFC9DAF8" },
  },
  alignment: {
    horizontal: "center",
    vertical: "center",
  },
  ...borders,
};
const centerClm = {
  alignment: {
    horizontal: "center",
    vertical: "center",
  },
  ...borders,
};

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

const deleteColumns = (ws, col, size = 100) => {
  for (let i = 1; i < size; i++) {
    if (ws[`${col}${i}`] == undefined) break;
    ws[`${col}${i}`].v = "";
  }
};

const changeClmnName = (ws, cl, newName = "") => {
  delete ws[`${cl}`].w;
  ws[`${cl}`].v = newName;
};
const addStyleToColm = (ws, colArr, style, size = 100) => {
  colArr.map((col) => {
    for (let i = 1; i <= size; i++) {
      if (ws[`${col}${i}`] == undefined) break;
      ws[`${col}${i}`].s = style;
    }
  });
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

    deleteColumns(sheet, "I", 10);

    addStyleToColm(sheet, headerLettersArr, borders);
    addStyleToColm(sheet, ["C", "D", "E", "H"], centerClm);
    addStyleToColm(sheet, headerLettersArr, headerStyle, 1);

    sheet["!cols"][8] = { wch: 20 };
    sheet["I1"].v = "Порушення";
    sheet["I1"].s = {
      ...headerStyle,
      fill: {
        fgColor: { rgb: "FFF7FF05" },
      },
    };
    xlsx.utils.sheet_add_aoa(sheet, [["Статус команди", ""]], {
      origin: "B11",
    });
    xlsx.utils.sheet_add_aoa(sheet, [["Внесок, грн", 0]], {
      origin: "B12",
    });
    sheet["C11"].s = {
      font: {
        sz: 14,
        color: {
          rgb: "FF000000",
        },
      },
      fill: {
        fgColor: { rgb: "FFFFFF00" },
      },
      ...borders,
    };
    sheet["C12"].s = {
      font: {
        sz: 14,
        color: {
          rgb: "FF000000",
        },
      },
      fill: {
        fgColor: { rgb: "FFEA4D4D" },
      },
      ...borders,
    };
    sheet["B11"].s = borders;
    sheet["B12"].s = borders;

    xlsx.utils.book_append_sheet(wb, sheet, team[0]["Назва команди"]);
  });
};

module.exports = {
  sortByTeamName,
  getTeamsNames,
  getTeamByName,
  setColmsLen,
  getAllTeams,
  appendAllSheets,
};
