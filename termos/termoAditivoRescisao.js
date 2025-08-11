// termos/termoAditivoRescisao.js
const { Document, Packer, Paragraph, TextRun, AlignmentType } = require('docx');

function dataPorExtenso(date = new Date(), cidade = "Teresina") {
  const meses = ['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'];
  return `${cidade},  ${date.getDate()} de ${meses[date.getMonth()].charAt(0).toUpperCase() + meses[date.getMonth()].slice(1)} de ${date.getFullYear()}`;
}

function p(txt, opts = {}) {
  return new Paragraph({
    alignment: opts.align || AlignmentType.JUSTIFIED,
    spacing: { after: opts.after ?? 300 },
    children: [ new TextRun({ text: txt, font: "Calibri", size: opts.size ?? 28, bold: !!opts.bold, allCaps: !!opts.allCaps }) ]
  });
}

function sanitizeFileName(str = "") {
  return String(str)
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^\w.-]+/g, "_");
}

async function gerarTermoAditivoRescisao(dados) {
  const {
    tipoContrato,               // "PRESTAÇÃO DE SERVIÇOS" ou "FORNECIMENTO DE BENS"
    razaoSocialContratada,
    cnpjContratada,
    enderecoContratada,
    numeroContrato,             // "X/AAAA"
    dataAssinaturaContrato,     // "YYYY-MM-DD"
    objetoOriginal,             // descrição
    cidadeData = "Teresina"
  } = dados;

  // Validações básicas
  const obrig = ["tipoContrato","razaoSocialContratada","cnpjContratada","enderecoContratada","numeroContrato","dataAssinaturaContrato","objetoOriginal"];
  const falt = obrig.filter(k => !dados[k] || String(dados[k]).trim() === "");
  if (falt.length) {
    const err = new Error(`Campos obrigatórios ausentes: ${falt.join(", ")}`);
    err.status = 400;
    throw err;
  }
  const tipoUp = String(tipoContrato).trim().toUpperCase();
  if (tipoUp !== "PRESTAÇÃO DE SERVIÇOS" && tipoUp !== "FORNECIMENTO DE BENS") {
    const err = new Error('tipoContrato deve ser "PRESTAÇÃO DE SERVIÇOS" ou "FORNECIMENTO DE BENS"');
    err.status = 400;
    throw err;
  }
  if (isNaN(Date.parse(dataAssinaturaContrato))) {
    const err = new Error("dataAssinaturaContrato inválida. Use formato YYYY-MM-DD.");
    err.status = 400;
    throw err;
  }

  const dataAssinStr = new Date(dataAssinaturaContrato).toLocaleDateString('pt-BR');

  const doc = new Document({
    sections: [{
      properties: { page: { margin: { top: 1440, bottom: 1080, left: 1440, right: 1440 } } },
      children: [
        new Paragraph({ spacing: { after: 400 } }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
          children: [ new TextRun({ text: "TERMO ADITIVO AO CONTRATO Nº", bold: true, allCaps: true, size: 32, font: "Calibri" }) ]
        }),
        p(`${numeroContrato} DE ${tipoUp},`, { align: AlignmentType.CENTER, after: 50 }),
        p("VISANDO A RESCISÃO ADMINISTRATIVA UNILATERAL.", { align: AlignmentType.CENTER, after: 700 }),

        p(
          `Pelo presente instrumento, a ÁGUAS E ESGOTOS DO PIAUÍ S.A. – AGESPISA, sociedade de economia mista estadual, inscrita no CNPJ sob o nº 06.845.747/0001-27, com sede na Av. Mal. Castelo Branco, Nº 101/N, Bairro Cabral, na cidade de Teresina-PI, neste ato representada por seu Diretor-Presidente, GARCIAS GUEDES RODRIGUES JÚNIOR, doravante denominada simplesmente CONTRATANTE; frente a empresa ${razaoSocialContratada}, inscrita no CNPJ sob o nº ${cnpjContratada}, com sede na ${enderecoContratada}, doravante denominada CONTRATADA, unilateralmente resolve celebrar o presente Termo Aditivo de Rescisão Unilateral, com fundamento no artigo 137, inciso VIII, da Lei nº 14.133, de 1º de abril de 2021, mediante as cláusulas e condições a seguir pactuadas:`
        ),

        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          spacing: { after: 100 },
          children: [ new TextRun({ text: "CLÁUSULA PRIMEIRA – DO OBJETO", bold: true, font: "Calibri", size: 28 }) ]
        }),
        p(
          `1.1 O presente Termo Aditivo tem por objeto a rescisão unilateral do Contrato nº ${numeroContrato}, celebrado em ${dataAssinStr}, cujo objeto consiste em ${objetoOriginal}, em razão da superveniência de razões de interesse público, consistentes na extinção da função operacional da AGESPISA e na preparação para sua liquidação societária, conforme autorizado pela Lei Complementar nº 319, de 16 de julho de 2025.`
        ),

        p(dataPorExtenso(new Date(), cidadeData), { align: AlignmentType.CENTER, after: 400 }),

        new Paragraph({ spacing: { after: 450 } }),

        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [ new TextRun({ text: "GARCIAS GUEDES RODRIGUES JÚNIOR", bold: true, font: "Calibri", size: 28 }) ]
        }),
        p("Diretor-Presidente – AGESPISA", { align: AlignmentType.CENTER, size: 24 }),

        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 300 },
          children: [ new TextRun({ text: razaoSocialContratada, bold: true, font: "Calibri", size: 28 }) ]
        }),
        p("Representante Legal da Contratada", { align: AlignmentType.CENTER, size: 24, after: 0 }),
      ]
    }]
  });

  const buffer = await Packer.toBuffer(doc);
  const fileName = `Termo_Aditivo_Rescisao_${sanitizeFileName(numeroContrato)}.docx`;
  return { buffer, fileName };
}

module.exports = { gerarTermoAditivoRescisao };
