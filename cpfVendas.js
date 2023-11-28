const XLSX = require('xlsx');

function getItemsByCPF(filePath, targetCPF) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet);

    const customerItems = {};

    data.forEach(row => {
        if (row['CPF/CNPJ'] && row.PRODUTO) {
            const cpf = row['CPF/CNPJ'].replace(/[^\d]/g, '');
            const cleanedCPF = targetCPF.replace(/[^\d]/g, '');
            if (cpf === cleanedCPF) {
                const productName = row.PRODUTO.trim();
                const quantity = parseFloat(row.QUANTIDADE) || 0;
                customerItems[productName] = (customerItems[productName] || 0) + quantity;
            }
        }
    });

    const itemsList = Object.keys(customerItems)
        .map(item => ({ item, count: customerItems[item] }))
        .sort((a, b) => b.count - a.count)
        .slice(0, 100);

    return itemsList;
}

const filePath = 'C:\\Users\\lrhau\\Desktop\\narciso teste\\Lista de compras inativos lj 1 .xlsx';
const targetCPF = '096.117.987-25';

const customerItems = getItemsByCPF(filePath, targetCPF);

if (customerItems.length === 0) {
    console.log('Nenhum item encontrado para o CPF especificado.');
} else {
    console.log(`O CLIENTE comprou os seguintes produtos:`);
    customerItems.forEach((item, index) => {
        console.log(`${index + 1}. ${item.item}: ${item.count} unidades`);
    });
}
