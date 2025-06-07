const XLSX = require("xlsx");

const arquivoXLSX = XLSX.readFile("rela_insc_homologadas-2.xlsx");

const arquivo = arquivoXLSX.SheetNames[0];
const arquivoConvertido = arquivoXLSX.Sheets[arquivo];

const dados = XLSX.utils.sheet_to_json(arquivoConvertido);

const cursoOpcoes = {
    opcaoA: "ENGENHARIA ELETRÔNICA",
    opcaoB: "ENGENHARIA QUÍMICA",
    opcaoC: "CIÊNCIA DA COMPUTAÇÃO",
}

const inscritosEngenhariaEletronica = dados.filter(item => 
    (item.__EMPTY_3 && item.__EMPTY_3.includes(cursoOpcoes.opcaoB) && 
    item.__EMPTY_2 === "CAMPO MOURÃO") ||
    (item.__EMPTY_5 && item.__EMPTY_5.includes(cursoOpcoes.opcaoB) && 
    item.__EMPTY_4 === "CAMPO MOURÃO")
);

const incritoPrimeiraOpcao = dados.filter(item =>
    (item.__EMPTY_3 && item.__EMPTY_3.includes(cursoOpcoes.opcaoB) && 
    item.__EMPTY_2 === "CAMPO MOURÃO")
);

const segundaPrimeiraOpcao = dados.filter(item =>
    (item.__EMPTY_5 && item.__EMPTY_5.includes(cursoOpcoes.opcaoB) && 
    item.__EMPTY_4 === "CAMPO MOURÃO")
);

console.log(`Total de inscritos em ${cursoOpcoes.opcaoB} no Campus Campo Mourão: ${inscritosEngenhariaEletronica.length}`);
console.log(`Inscritos Apenas na Primeira Opção em ${cursoOpcoes.opcaoB} no Campus Campo Mourão: ${incritoPrimeiraOpcao.length}`);
console.log(`Inscritos Apenas na Segunda Opção em ${cursoOpcoes.opcaoB} no Campus Campo Mourão: ${segundaPrimeiraOpcao.length}`);

//Traz o total tanto quem é da primeira opção com oseleção quanto da segunda opção como seleção
// console.log(inscritosEngenhariaEletronica);