const XLSX = require("xlsx");

const arquivoXLSX = XLSX.readFile("rela_insc_homologadas-2.xlsx");

const arquivo = arquivoXLSX.SheetNames[0];
const arquivoConvertido = arquivoXLSX.Sheets[arquivo];

const dados = XLSX.utils.sheet_to_json(arquivoConvertido);

const inscritosEngenhariaEletronica = dados.filter(item => 
    (item.__EMPTY_3 && item.__EMPTY_3.includes("ENGENHARIA ELETRÔNICA") && 
    item.__EMPTY_2 === "CAMPO MOURÃO") ||
    (item.__EMPTY_5 && item.__EMPTY_5.includes("ENGENHARIA ELETRÔNICA") && 
    item.__EMPTY_4 === "CAMPO MOURÃO")
);

console.log(`Total de inscritos em Engenharia Eletrônica no Campus Campo Mourão: ${inscritosEngenhariaEletronica.length}`);
console.log(inscritosEngenhariaEletronica);