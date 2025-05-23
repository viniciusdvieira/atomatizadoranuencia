const express = require('express');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, AlignmentType } = require('docx');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public'))); // Serve o frontend

app.post('/gerar-termo', async (req, res) => {
    const { processo, nome, valor, numPrestacoes } = req.body;

    const valorNum = parseFloat(valor);
    const prestacoes = parseInt(numPrestacoes);
    const valorPrestacao = valorNum / prestacoes;

    // Data formatada
    const meses = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
        'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
    const data = new Date();
    const dataFormatada = `Teresina,  ${data.getDate()} de ${meses[data.getMonth()].charAt(0).toUpperCase() + meses[data.getMonth()].slice(1)} de ${data.getFullYear()}`;

    // Valor por extenso
    const extenso = require('numero-por-extenso');
    const valorExtenso = extenso.porExtenso(valorNum, 'monetario');
    const valorParcelaExtenso = extenso.porExtenso(valorPrestacao, 'monetario');

    const doc = new Document({
        sections: [{
            properties: {
                page: {
                    margin: { top: 1440, bottom: 1080, left: 1440, right: 1440 }
                }
            },
            children: [
                new Paragraph({
                    spacing: { after: 400 } 
                }),

                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                        new TextRun({
                            text: "TERMO DE ANUÊNCIA",
                            bold: true,
                            allCaps: true,
                            size: 32,
                            font: "Calibri"
                        })
                    ],
                    spacing: { after: 700 }
                }),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    children: [
                        new TextRun({
                            text: `Processo SEI ${processo}`,
                            font: "Calibri",
                            size: 28
                        })
                    ],
                    spacing: { after: 300 }
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
                    children: [
                        new TextRun({
                            text: dataFormatada,
                            font: "Calibri",
                            size: 28
                        })
                    ],
                    spacing: { after: 400 }
                }),

                new Paragraph({
                    spacing: { after: 450 } 
                }),

                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                        new TextRun({
                            text: nome,
                            bold: true,
                            font: "Calibri",
                            size: 28
                        })
                    ]
                }),
            ]
        }]
    });

    const buffer = await Packer.toBuffer(doc);

    const sanitizeFileName = nome.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, "_");
    const fileName = `Termo_Anuencia_${sanitizeFileName}.docx`;
    
    res.set({
        'Content-Disposition': `attachment; filename="${fileName}"`,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });

    res.send(buffer);
});


app.listen(3000, () => {
    console.log("Servidor rodando em http://localhost:3000");
});
