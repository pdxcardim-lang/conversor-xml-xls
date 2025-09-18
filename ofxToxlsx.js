// ==========================================================
// SCRIPT: Ler arquivo OFX com ofx-js e salvar em Excel
// Requisitos:
//   npm install ofx-js xlsx inquirer
// ==========================================================

import fs from "fs";
import path from "path";
import { parse } from "ofx-js";
import * as XLSX from "xlsx";
import inquirer from "inquirer";

// ==========================================================
// Função: desfazer execução existente
// ==========================================================
async function desfazerExecucao(caminhoExcel) {
  const workbook = XLSX.readFile(caminhoExcel);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  const { execNum } = await inquirer.prompt([
    {
      type: "number",
      name: "execNum",
      message: "Digite o número da execução para remover:"
    }
  ]);

  const novaMatriz = [data[0]]; // cabeçalho
  let removidas = 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][data[i].length - 1] !== execNum) {
      novaMatriz.push(data[i]);
    } else {
      removidas++;
    }
  }

  if (removidas === 0) {
    console.log(`⚠️ Nenhuma linha encontrada para a execução ${execNum}.`);
  } else {
    const newSheet = XLSX.utils.aoa_to_sheet(novaMatriz);
    const newWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWb, newSheet, "Extrato OFX");
    XLSX.writeFile(newWb, caminhoExcel);
    console.log(`✅ Foram removidas ${removidas} linhas da execução ${execNum}.`);
  }
}

// ==========================================================
// 1. Selecionar arquivo OFX
// ==========================================================
const { caminhoArquivo } = await inquirer.prompt([
  {
    type: "input",
    name: "caminhoArquivo",
    message: "Digite o caminho do arquivo OFX:"
  }
]);

if (!fs.existsSync(caminhoArquivo)) {
  console.error("❌ Nenhum arquivo foi encontrado.");
  process.exit(1);
}

const saida = path.join(path.dirname(caminhoArquivo), "extrato_ofx.xlsx");

// Perguntar se deseja desfazer execução
if (fs.existsSync(saida)) {
  const { opcao } = await inquirer.prompt([
    {
      type: "confirm",
      name: "opcao",
      message: "Deseja desfazer uma execução existente?"
    }
  ]);

  if (opcao) {
    await desfazerExecucao(saida);
    process.exit(0);
  }
}

// ==========================================================
// 2. Ler OFX
// ==========================================================
const ofxData = fs.readFileSync(caminhoArquivo, "utf8");
const json = await parse(ofxData);

// Achar transações
let transacoes = [];
let bancoId = "N/A", contaId = "N/A", contaTipo = "N/A";

if (json.OFX.BANKMSGSRSV1?.STMTTRNRS) {
  const stmt = json.OFX.BANKMSGSRSV1.STMTTRNRS.STMTRS;
  transacoes = stmt.BANKTRANLIST.STMTTRN;
  bancoId = stmt.BANKACCTFROM?.BANKID || "N/A";
  contaId = stmt.BANKACCTFROM?.ACCTID || "N/A";
  contaTipo = stmt.BANKACCTFROM?.ACCTTYPE || "N/A";
} else if (json.OFX.CREDITCARDMSGSRSV1?.CCSTMTTRNRS) {
  const stmt = json.OFX.CREDITCARDMSGSRSV1.CCSTMTTRNRS.CCSTMTRS;
  transacoes = stmt.BANKTRANLIST.STMTTRN;
}

// Se não achar nada, aborta
if (!transacoes || transacoes.length === 0) {
  console.error("❌ Não foi possível localizar lançamentos neste OFX.");
  process.exit(1);
}

// ==========================================================
// 3. Abrir ou criar planilha
// ==========================================================
let data = [];
let execucaoAtual = 1;

if (fs.existsSync(saida)) {
  const workbook = XLSX.readFile(saida);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  const execucoes = data.slice(1).map(row => row[row.length - 1]).filter(Boolean);
  execucaoAtual = execucoes.length ? Math.max(...execucoes) + 1 : 1;
} else {
  data.push([
    "Data", "Tipo", "Descrição", "Valor (R$)", "Saldo acumulado (R$)",
    "Banco ID", "Processado em", "Execução"
  ]);
}

// ==========================================================
// 4. Processar transações
// ==========================================================
let saldo = 0;
let processados = 0;
const dataProcesso = new Date().toLocaleString("pt-BR");

for (const trn of transacoes) {
  const valor = parseFloat(trn.TRNAMT);
  saldo += valor;
  const tipo = valor > 0 ? "Entrada" : "Saída";
  const descricao = trn.NAME || "Sem descrição";
  const dataMov = new Date(trn.DTPOSTED.substring(0,8)).toLocaleDateString("pt-BR");

  processados++;
  data.push([
    dataMov, tipo, descricao, valor, saldo,
    bancoId, dataProcesso, execucaoAtual
  ]);
}

// ==========================================================
// 5. Totais
// ==========================================================
const totalEntradas = transacoes.filter(t => parseFloat(t.TRNAMT) > 0)
                                .reduce((s, t) => s + parseFloat(t.TRNAMT), 0);
const totalSaidas = transacoes.filter(t => parseFloat(t.TRNAMT) < 0)
                              .reduce((s, t) => s + parseFloat(t.TRNAMT), 0);

// ==========================================================
// 6. Salvar planilha
// ==========================================================
const newSheet = XLSX.utils.aoa_to_sheet(data);
const newWb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(newWb, newSheet, "Extrato OFX");
XLSX.writeFile(newWb, saida);

console.log(`
✅ Planilha atualizada com sucesso: ${saida}

Banco: ${bancoId} | Conta: ${contaId} (${contaTipo})
Transações adicionadas: ${processados}
Execução nº ${execucaoAtual}

Resumo:
 - Saldo final: R$ ${saldo.toFixed(2)}
 - Entradas: R$ ${totalEntradas.toFixed(2)}
 - Saídas: R$ ${totalSaidas.toFixed(2)}
`);
