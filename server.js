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
app.post('/gerar-contrato-estagio', async (req, res) => {
  const {
    nomeEstagiario, cpfEstagiario, ruaEstagiario, bairroEstagiario, cepEstagiario, cidadeEstagiario, curso, periodo,
    nomeInstituicao, cnpjInstituicao, ruaInstituicao, bairroInstituicao, cepInstituicao, cidadeInstituicao,
    nomeOrgao, cnpjOrgao, ruaOrgao, bairroOrgao, cepOrgao, cidadeOrgao, representanteOrgao
  } = req.body;

  const enderecoEstagiario = `${ruaEstagiario}, Bairro: ${bairroEstagiario}, CEP: ${cepEstagiario}, ${cidadeEstagiario}`;
  const enderecoInstituicao = `${ruaInstituicao}, Bairro: ${bairroInstituicao}, CEP: ${cepInstituicao}, ${cidadeInstituicao}`;
  const enderecoOrgao = `${ruaOrgao}, Bairro: ${bairroOrgao}, CEP: ${cepOrgao}, ${cidadeOrgao}`;

  function p(text, opts = {}) {
    return new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing: { after: 300 },
      children: [new TextRun({ text, font: 'Calibri', size: 28, ...opts })]
    });
  }

  const clausulas = [
    p("CLÁUSULA PRIMEIRA - O estágio oferecido para discentes do curso de " + curso + ", no qual o estudante está cursando o " + periodo + "º período, é regido por este Termo de Compromisso, visando propiciar ao estudante uma experiência acadêmico-profissional."),
    p("CLÁUSULA SEGUNDA - O estágio será desenvolvido com base em programação elaborada em comum acordo entre as partes e terá acompanhamento efetivo do orientador da instituição de ensino e do supervisor da concedente."),
    p("CLÁUSULA TERCEIRA - O(A) estagiário(a) cumprirá carga horária de até 30 horas semanais, não ultrapassando 6 horas diárias, compatíveis com o horário escolar."),
    p("CLÁUSULA QUARTA - A duração do estágio não poderá exceder 2 (dois) anos, exceto quando se tratar de estagiário(a) portador(a) de deficiência."),
    p("CLÁUSULA QUINTA - O(A) estagiário(a) receberá bolsa-auxílio e auxílio-transporte conforme legislação vigente e disponibilidade orçamentária da concedente."),
    p("CLÁUSULA SEXTA - A instituição de ensino se compromete a indicar professor orientador responsável pelo acompanhamento e avaliação das atividades desenvolvidas pelo(a) estagiário(a)."),
    p("CLÁUSULA SÉTIMA - Compete à concedente indicar servidor como supervisor do estágio, responsável por orientar e acompanhar as atividades do(a) estagiário(a)."),
    p("CLÁUSULA OITAVA - O estágio poderá ser rescindido a qualquer tempo por qualquer das partes, mediante comunicação por escrito, com antecedência mínima de 5 (cinco) dias úteis."),
    p("CLÁUSULA NONA - O(A) estagiário(a) deverá apresentar à concedente, ao final do estágio, relatório das atividades desenvolvidas, com avaliação do orientador da instituição de ensino."),
    p("CLÁUSULA DÉCIMA - Este Termo de Compromisso entra em vigor na data de sua assinatura pelas partes e terá validade conforme previsto nas cláusulas anteriores.")
  ];

  const doc = new Document({
    sections: [{
      properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
      children: [
        p(`Pelo presente instrumento as partes abaixo discriminadas: Firmam entre si TERMO DE COMPROMISSO PARA A REALIZAÇÃO DE ESTÁGIO, regido pela Lei nº. 11.788, de 25 de setembro de 2008, e no que couber pelo Decreto estadual nº. 13.840, de 21 de setembro de 2009, segundo as seguintes cláusulas:`),
        p(`CONCEDENTE: ${nomeOrgao}, CNPJ: ${cnpjOrgao}, com endereço na ${enderecoOrgao}, representada por ${representanteOrgao}.`, { bold: true }),
        p(`INTERVENIENTE: SECRETARIA DA ADMINISTRAÇÃO DO ESTADO DO PIAUÍ - SEAD, CNPJ: 06.553.481/0003-00, com endereço na Av. Pedro Freitas, S/N. BL. 01 – CENTRO ADMINISTRATIVO, Bairro São Pedro, em Teresina - PI, representada por SAMUEL PONTES DO NASCIMENTO.`, { bold: true }),
        p(`ESTAGIÁRIO(A): ${nomeEstagiario}, brasileiro(a), CPF: ${cpfEstagiario}, residente e domiciliado na ${enderecoEstagiario}.`, { bold: true }),
        p(`INSTITUIÇÃO DE ENSINO: ${nomeInstituicao}, CNPJ: ${cnpjInstituicao}, com sede na ${enderecoInstituicao}, que subscreve esse ato através do seu representante legal (coordenador do curso, secretário acadêmico ou preposto).`, { bold: true }),
        ...clausulas,
        p("E, por estarem de inteiro e comum acordo com as condições e dizeres deste instrumento, as partes assinam-no em meio digital via SEI-PI."),
        new Paragraph({ text: "\n\n___________________________________________", alignment: AlignmentType.CENTER }),
        new Paragraph({ text: `SECRETARIA DE GOVERNO – SEGOV\n${representanteOrgao}\nCONCEDENTE`, alignment: AlignmentType.CENTER }),
        new Paragraph({ text: "\n___________________________________________", alignment: AlignmentType.CENTER }),
        new Paragraph({ text: "SECRETARIA DA ADMINISTRAÇÃO DO ESTADO DO PIAUÍ\nSAMUEL PONTES DO NASCIMENTO\nINTERVENIENTE", alignment: AlignmentType.CENTER }),
        new Paragraph({ text: "\n___________________________________________", alignment: AlignmentType.CENTER }),
        new Paragraph({ text: `${nomeEstagiario}\nESTAGIÁRIO(A)`, alignment: AlignmentType.CENTER }),
        new Paragraph({ text: "\n___________________________________________", alignment: AlignmentType.CENTER }),
        new Paragraph({ text: `${nomeInstituicao}\nREPRESENTANTE LEGAL DA INSTITUIÇÃO`, alignment: AlignmentType.CENTER })
      ]
    }]
  });

  const buffer = await Packer.toBuffer(doc);
  const nomeArquivo = `Contrato_Estagio_${nomeEstagiario.normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/\s+/g, "_")}.docx`;

  res.set({
    'Content-Disposition': `attachment; filename="${nomeArquivo}"`,
    'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  });

  res.send(buffer);
});


app.listen(3000, () => {
    console.log("Servidor rodando em http://localhost:3000");
});
