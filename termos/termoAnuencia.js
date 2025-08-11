// termos/termoAnuencia.js
const { Document, Packer, Paragraph, TextRun, AlignmentType } = require('docx');
const extenso = require('numero-por-extenso');

function dataPorExtenso(date = new Date(), cidade = "Teresina") {
  const meses = ['janeiro','fevereiro','março','abril','maio','junho',
    'julho','agosto','setembro','outubro','novembro','dezembro'];
  return `${cidade},  ${date.getDate()} de ${meses[date.getMonth()].charAt(0).toUpperCase() + meses[date.getMonth()].slice(1)} de ${date.getFullYear()}`;
}

function sanitizeFileName(str = "") {
  return String(str)
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^\w.-]+/g, "_");
}

async function gerarTermoAnuencia(payload) {
  const { processo, nome, valor, numPrestacoes } = payload;

  const valorNum = parseFloat(valor);
  const prestacoes = parseInt(numPrestacoes, 10);

  if (!processo || !nome || isNaN(valorNum) || isNaN(prestacoes) || prestacoes <= 0) {
    const faltantes = [];
    if (!processo) faltantes.push("processo");
    if (!nome) faltantes.push("nome");
    if (isNaN(valorNum)) faltantes.push("valor (número)");
    if (isNaN(prestacoes) || prestacoes <= 0) faltantes.push("numPrestacoes (inteiro > 0)");
    const msg = `Campos obrigatórios inválidos/ausentes: ${faltantes.join(", ")}`;
    const err = new Error(msg);
    err.status = 400;
    throw err;
  }

  const valorPrestacao = valorNum / prestacoes;

  const valorExtenso = extenso.porExtenso(valorNum, 'monetario');
  const valorParcelaExtenso = extenso.porExtenso(valorPrestacao, 'monetario');

  const doc = new Document({
    sections: [{
      properties: {
        page: { margin: { top: 1440, bottom: 1080, left: 1440, right: 1440 } }
      },
      children: [
        new Paragraph({ spacing: { after: 400 } }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 700 },
          children: [
            new TextRun({ text: "TERMO DE ANUÊNCIA", bold: true, allCaps: true, size: 32, font: "Calibri" })
          ]
        }),
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          spacing: { after: 300 },
          children: [ new TextRun({ text: `Processo SEI ${processo}`, font: "Calibri", size: 28 }) ]
        }),
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          spacing: { after: 300 },
          children: [
            new TextRun({ text: "Eu, ", font: "Calibri", size: 28 }),
            new TextRun({ text: nome, bold: true, font: "Calibri", size: 28 }),
            new TextRun({ text: ", venho por meio deste, AUTORIZAR a Secretaria da Administração para que proceda o pagamento do crédito no valor de ", font: "Calibri", size: 28 }),
            new TextRun({ text: `R$ ${valorNum.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}`, bold: true, font: "Calibri", size: 28 }),
            new TextRun({ text: ` (${valorExtenso}) em ${prestacoes} prestações fixas de `, font: "Calibri", size: 28 }),
            new TextRun({ text: `R$ ${valorPrestacao.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}`, bold: true, font: "Calibri", size: 28 }),
            new TextRun({ text: ` (${valorParcelaExtenso}). e referente ao requerimento solicitado no processo SEI ${processo}.`, font: "Calibri", size: 28 })
          ]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
          children: [ new TextRun({ text: dataPorExtenso(new Date(), "Teresina"), font: "Calibri", size: 28 }) ]
        }),
        new Paragraph({ spacing: { after: 450 } }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [ new TextRun({ text: nome, bold: true, font: "Calibri", size: 28 }) ]
        }),
      ]
    }]
  });

  const buffer = await Packer.toBuffer(doc);
  const fileName = `Termo_Anuencia_${sanitizeFileName(nome)}.docx`;

  return { buffer, fileName };
}

module.exports = { gerarTermoAnuencia };
