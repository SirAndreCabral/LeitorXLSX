const XLSX = require("xlsx");

const arquivoXLSX = XLSX.readFile("rela_insc_homologadas_final.xlsx");

const planilhaDoExcel = arquivoXLSX.SheetNames[0];
const DadosDaPlanilha = arquivoXLSX.Sheets[planilhaDoExcel];

const dados = XLSX.utils.sheet_to_json(DadosDaPlanilha);

const cursoOpcoes = {
    opcaoA: "ENGENHARIA ELETRÔNICA",
    opcaoB: "ENGENHARIA QUÍMICA",
    opcaoC: "CIÊNCIA DA COMPUTAÇÃO",
    opcaoD: "ENGENHARIA AMBIENTAL",
    opcaoE: "ENGENHARIA CIVIL",
    opcaoF: "ENGENHARIA DE ALIMENTOS"
};

const CursoVaga = [
    { nome: "ENGENHARIA ELETRÔNICA", vagas: "24" },
    { nome: "ENGENHARIA QUÍMICA", vagas: "24" },
    { nome: "CIÊNCIA DA COMPUTAÇÃO", vagas: "44" },
    { nome: "ENGENHARIA AMBIENTAL", vagas: "55" },
    { nome: "ENGENHARIA CIVIL", vagas: "55" },
    { nome: "ENGENHARIA DE ALIMENTOS", vagas: "44" }
];

// vai selecionar o curso baseado no cursoOpcoes
const cursoSelecionado = CursoVaga.find(curso => curso.nome === cursoOpcoes.opcaoB);

// filtra inscritos que escolheram o curso e que são do campus de Campo Mourão
const inscritos = dados.filter(item => 
    cursoSelecionado && (
        (item.__EMPTY_3 && item.__EMPTY_3.includes(cursoSelecionado.nome) && item.__EMPTY_2 === "CAMPO MOURÃO") ||
        (item.__EMPTY_5 && item.__EMPTY_5.includes(cursoSelecionado.nome) && item.__EMPTY_4 === "CAMPO MOURÃO")
    )
);

console.log("--------------------------------------");
console.log(`Total de inscritos em ${cursoSelecionado.nome} no Campus Campo Mourão: ${inscritos.length}`);
console.log(`Total de vagas no curso ${cursoSelecionado.vagas}`);
console.log("--------------------------------------");

// pega só os alunos que selecionaram na primeira opção
const incritoPrimeiraOpcao = dados.filter(item =>
    (item.__EMPTY_3 && item.__EMPTY_3.includes(cursoSelecionado.nome) && 
    item.__EMPTY_2 === "CAMPO MOURÃO")
);

// pega só os alunos que selecionaram na segunda opção
const segundaPrimeiraOpcao = dados.filter(item =>
    (item.__EMPTY_5 && item.__EMPTY_5.includes(cursoSelecionado.nome) && 
    item.__EMPTY_4 === "CAMPO MOURÃO")
);

console.log(`Inscritos Apenas na Primeira Opção em ${cursoSelecionado.nome} no Campus Campo Mourão: ${incritoPrimeiraOpcao.length}`);

console.log(`Inscritos Apenas na Segunda Opção em ${cursoSelecionado.nome} no Campus Campo Mourão: ${segundaPrimeiraOpcao.length}`);
