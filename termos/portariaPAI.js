// termos/portariaPAI.js
const { Document, Packer, Paragraph, TextRun, AlignmentType } = require('docx');

function p(txt, opts = {}) {
  return new Paragraph({
    alignment: opts.align || AlignmentType.JUSTIFIED,
    spacing: { after: opts.after ?? 300 },
    children: [
      new TextRun({
        text: txt,
        font: "Calibri",
        size: opts.size ?? 28,
        bold: !!opts.bold,
        allCaps: !!opts.allCaps
      })
    ]
  });
}

function toPtBR(d) {
  // aceita Date ou string YYYY-MM-DD
  const dt = (d instanceof Date) ? d : new Date(String(d) + "T00:00:00");
  if (isNaN(dt.getTime())) return null;
  return dt.toLocaleDateString('pt-BR');
}

function sanitizeFileName(str = "") {
  return String(str)
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^\w.-]+/g, "_");
}

async function gerarPortariaPAI(dados) {
  const {
    numeroRes,
    dataRes,         // YYYY-MM-DD
    cargoFunc,
    nomeFunc,
    matriculaFunc,
    numSei,
    lotacao,
    dataVigorRes     // YYYY-MM-DD
  } = dados;

  // Validações básicas
  const obrig = ["numeroRes","dataRes","cargoFunc","nomeFunc","matriculaFunc","numSei","lotacao","dataVigorRes"];
  const falt = obrig.filter(k => !dados[k] || String(dados[k]).trim() === "");
  if (falt.length) {
    const err = new Error(`Campos obrigatórios ausentes: ${falt.join(", ")}`);
    err.status = 400;
    throw err;
  }

  const dataResBR = toPtBR(dataRes);
  const dataVigorBR = toPtBR(dataVigorRes);
  if (!dataResBR) {
    const err = new Error("dataRes inválida. Use formato YYYY-MM-DD.");
    err.status = 400; throw err;
  }
  if (!dataVigorBR) {
    const err = new Error("dataVigorRes inválida. Use formato YYYY-MM-DD.");
    err.status = 400; throw err;
  }

  const doc = new Document({
    sections: [{
      properties: { page: { margin: { top: 1440, bottom: 1080, left: 1440, right: 1440 } } },
      children: [
        new Paragraph({ spacing: { after: 400 } }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 300 },
          children: [
            new TextRun({ text: "PORTARIA", bold: true, allCaps: true, size: 32, font: "Calibri" })
          ]
        }),

        p(`A ÁGUAS E ESGOTOS DO PIAUÍ S/A – AGESPISA, por meio de seu Diretor-Presidente, no uso das atribuições que lhe confere o Estatuto Social e Jurídico da Empresa,`),

        p(`CONSIDERANDO o disposto na Resolução SEI nº ${numeroRes}, de ${dataResBR}, que institui o novo Programa de Afastamento Incentivado (PAI) para empregados aposentados ou não, integrantes do quadro de provimento efetivo da AGESPISA;`),

        p(`CONSIDERANDO que o empregado desta empresa, ${cargoFunc}, ${nomeFunc} - MAT. ${matriculaFunc}, através do Termo de Adesão ao PAI/2025, aderiu ao novo Programa de Afastamento Incentivado – PAI, Processo SEI nº ${numSei};`),

        p(`CONSIDERANDO que o empregado preenche todos os requisitos elencados na citada Resolução, estando apto para adesão ao novo Programa de Afastamento Incentivado – PAI;`),

        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          spacing: { after: 100 },
          children: [ new TextRun({ text: "RESOLVE:", bold: true, font: "Calibri", size: 28 }) ]
        }),

        p(`1º) Rescindir, a pedido, o contrato de trabalho do empregado desta Empresa, ${cargoFunc}, ${nomeFunc} - MAT. ${matriculaFunc}, lotado no ELO ${lotacao}, nos termos do novo Programa de Afastamento Incentivado – PAI, com percepção de todos os direitos e vantagens indenizatórias, de caráter trabalhista, previstos na Resolução nº ${numeroRes}, de ${dataResBR}.`),

        p(`2º) Determinar que a Diretoria Administrativa e de Gestão Corporativa - DIAGC, através da Superintendência de Gestão de Pessoas - SUGEP, adote as providências necessárias ao cumprimento da presente Portaria.`),

        p(`3º) Revogadas as disposições em contrário, os efeitos da presente portaria entram em vigor na data de ${dataVigorBR}.`),

        new Paragraph({ spacing: { after: 450 } }),

        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [ new TextRun({ text: "GARCIAS GUEDES RODRIGUES JÚNIOR", bold: true, font: "Calibri", size: 28 }) ]
        }),
        p("Diretor-Presidente – AGESPISA", { align: AlignmentType.CENTER, size: 24 }),
      ]
    }]
  });

  const buffer = await Packer.toBuffer(doc);
  const fileName = `Portaria_PAI_${sanitizeFileName(nomeFunc)}_${sanitizeFileName(matriculaFunc)}.docx`;
  return { buffer, fileName };
}

module.exports = { gerarPortariaPAI };
