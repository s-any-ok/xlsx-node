
# Sheetjs-style
support set cell style for sheetjs!
API is the same as sheetjs!


## install
```
npm install sheetjs-style
```

## How to Use?
Please read [SheetJs Documents](https://github.com/SheetJS/sheetjs/blob/3468395494c450ea8ba7e20afb1bd6127f516ccd/README.md)!

## How to set cell Style?
for example:
```js
ws["A1"].s = {									// set the style for target cell
  font: {
    name: '宋体',
    sz: 24,
    bold: true,
    color: { rgb: "FFFFAA00" }
  },
};
```

## Thanks
[sheetjs](https://github.com/SheetJS/sheetjs)
[js-xlsx](https://github.com/protobi/js-xlsx)