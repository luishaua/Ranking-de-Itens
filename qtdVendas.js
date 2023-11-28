const XLSX = require('xlsx');
const workbook = XLSX.readFile('C:\\Users\\lrhau\\Desktop\\narciso teste\\Lista de compras inativos lj 1 .xlsx');
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];
const data = XLSX.utils.sheet_to_json(worksheet);

const itemCounts = {};

data.forEach(row => {
    if (row.PRODUTO) {
        const items = row.PRODUTO.split(', ');
        items.forEach(item => {
            itemCounts[item] = (itemCounts[item] || 0) + 1;
        });
    }
});

const itemRanking = Object.keys(itemCounts)
    .map(item => ({ item, count: itemCounts[item] }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 100);

itemRanking.forEach((item, index) => {
    console.log(`${index + 1}. ${item.item}: ${item.count} vendas`);
});