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

  const enderecoEstagiario = `${ruaEstagiario}, Bairro: ${bairroEstagiario}, CEP: ${cepEstagiario}, ${cidadeEstagiario}-PI.`;
  const enderecoInstituicao = `${ruaInstituicao}, N°: XXXX, Bairro: ${bairroInstituicao}, ${cidadeInstituicao}-PI`;
  const enderecoOrgao = `${ruaOrgao}, N°: XXXX, Bairro: ${bairroOrgao}, CEP:${cepOrgao}, em ${cidadeOrgao}`;

  const doc = new Document({
    sections: [{
      properties: {
        page: { margin: { top: 1440, bottom: 1080, left: 1440, right: 1440 } }
      },
      children: [
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        children: [
            new TextRun({ text: "Pelo presente instrumento as partes abaixo discriminadas:", font: "Calibri", size: 24 })
        ]
        }),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        children: [
            new TextRun({ text: "Firmam entre si ", font: "Calibri", size: 24 }),
            new TextRun({ text: "TERMO DE COMPROMISSO PARA A REALIZAÇÃO DE ESTÁGIO", bold: true, font: "Calibri", size: 24 }),
            new TextRun({ text: ", regido pela Lei nº. 11.788, de 25 de setembro de 2008, e no que couber pelo Decreto estadual nº. 13.840, de 21 de setembro de 2009, segundo as seguintes cláusulas:", font: "Calibri", size: 24 })
        ]
        }),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        children: [
            new TextRun({ text: "CONCEDENTE: ", bold: true, font: "Calibri", size: 24 }),
            new TextRun({ text: "ESTADO DO PIAUÍ, através da ", font: "Calibri", size: 24 }),
            new TextRun({ text: nomeOrgao.toUpperCase(), bold: true, font: "Calibri", size: 24 }),
            new TextRun({ text: ", CNPJ: ", font: "Calibri", size: 24 }),
            new TextRun({ text: cnpjOrgao, font: "Calibri", size: 24 }),
            new TextRun({ text: ", com endereço na ", font: "Calibri", size: 24 }),
            new TextRun({ text: enderecoOrgao, font: "Calibri", size: 24 }),
            new TextRun({ text: ", representada por ", font: "Calibri", size: 24 }),
            new TextRun({ text: representanteOrgao.toUpperCase() + ".", bold: true, font: "Calibri", size: 24 })
        ]
        }),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        children: [
            new TextRun({ text: "INTERVENIENTE: SECRETARIA DA ADMINISTRAÇÃO DO ESTADO DO PIAUÍ - SEAD,", bold: true, font: "Calibri", size: 24 }),
            new TextRun({
            text: " CNPJ: 06.553.481/0003-00, com endereço na Av. Pedro Freitas, S/N. BL. 01 - CENTRO ADMINISTRATIVO, Bairro São Pedro, em Teresina - PI, representada por ",
            font: "Calibri", size: 24
            }),
            new TextRun({ text: "SAMUEL PONTES DO NASCIMENTO.", bold: true, font: "Calibri", size: 28 })
        ]
        }),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        children: [
            new TextRun({ text: "ESTAGIÁRIO(A): ",  bold: true ,font: "Calibri", size: 24 }),
            new TextRun({ text: nomeEstagiario.toUpperCase(), bold: true, font: "Calibri", size: 24 }),
            new TextRun({ text: `, brasileiro, CPF: ${cpfEstagiario}, residente e domiciliado na ${ruaEstagiario}, Bairro: ${bairroEstagiario}, CEP: ${cepEstagiario} em ${cidadeEstagiario}-PI.`, font: "Calibri", size: 24 })
        ]
        }),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 600 },
        children: [
            new TextRun({ text: "INSTITUIÇÃO DE ENSINO: ", bold: true ,font: "Calibri", size: 24 }),
            new TextRun({ text: nomeInstituicao.toUpperCase(), bold: true, font: "Calibri", size: 24 }),
            new TextRun({ text: `, CNPJ: ${cnpjInstituicao}, com sede na ${ruaInstituicao}, Bairro: ${bairroInstituicao}, CEP: ${cepInstituicao}, ${cidadeInstituicao}-PI, que subscreve esse ato através do seu representante legal (coordenador do curso, secretário acadêmico ou preposto).`, font: "Calibri", size: 24 })
        ]
        }),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        children: [
            new TextRun({ text: "CLÁUSULA PRIMEIRA", bold: true, font: "Calibri", size: 26 })
        ]
        }),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        children: [
            new TextRun({ text: "O estágio oferecido para discentes do curso de ", font: "Calibri", size: 24 }),
            new TextRun({ text: curso, bold: true, font: "Calibri", size: 24 }),
            new TextRun({ text: ", no qual o estudante está cursando o ", font: "Calibri", size: 24 }),
            new TextRun({ text: `${cnpjInstituicao} período `, bold: true, font: "Calibri", size: 24 }),
            new TextRun({ text: "em andamento, é regido por este Termo de Compromisso, visando propiciar ao estudante uma experiência acadêmico - profissional em um campo de trabalho determinado, visando:", font: "Calibri", size: 24 })
        ]
        }),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        children: [
            new TextRun({ text: "PARÁGRAFO PRIMEIRO: ",font: "Calibri", size: 24 }),
            new TextRun({ text: "O estágio não cria vínculo empregatício de qualquer natureza, devendo observar os seguintes requisitos:", font: "Calibri", size: 24 })
        ]
        }),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 100 },
        children: [
            new TextRun({ text: "I - matrícula e frequência regular do educando em curso de educação superior, de educação profissional ou de ensino médio e atestados pela instituição de ensino;", font: "Calibri", size: 24 })
        ]
        }),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        children: [
            new TextRun({ text: "II - compatibilidade entre as atividades desenvolvidas no estágio e aquelas previstas neste termo de compromisso.", font: "Calibri", size: 24 })
        ]
        }),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        children: [
            new TextRun({ text: "PARÁGRAFO SEGUNDO: ", font: "Calibri", size: 24 }),
            new TextRun({ text: "As atividades a serem desenvolvidas durante o ESTÁGIO, objeto do presente TERMO DE COMPROMISSO, constarão no Plano de Atividades construído pelo ESTAGIÁRIO em conjunto com a CONCEDENTE e orientado por professor da INSTITUIÇÃO DE ENSINO.", font: "Calibri", size: 24 })
        ]
        }),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 600 },
        children: [
            new TextRun({ text: "PARÁGRAFO TERCEIRO: ",font: "Calibri", size: 24 }),
            new TextRun({ text: "O plano de atividades do estagiário deverá ser incorporado ao termo de compromisso por meio de aditivos à medida que for avaliado, progressivamente, o desempenho do estudante. (Art. 7º, parágrafo único, da lei 11.788/2008).", font: "Calibri", size: 24 })
        ]
        }),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        children: [
            new TextRun({ text: "CLÁUSULA SEGUNDA", bold: true, font: "Calibri", size: 24 })
        ]
        }),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        children: [
            new TextRun({ text: "O estágio será desenvolvido no ", font: "Calibri", size: 24 }),
            new TextRun({ text: "período de 12 (doze) meses", bold: true, font: "Calibri", size: 24 }),
            new TextRun({ text: ", contados a partir da data da última assinatura deste termo, no horário determinado pela escala do concedente, em acordo com a Coordenação do curso, não podendo ultrapassar 20 (vinte) horas semanais, durante o período letivo.", font: "Calibri", size: 24 })
        ]
        }),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        children: [
            new TextRun({ text: "PARÁGRAFO PRIMEIRO: ",  font: "Calibri", size: 24 }),
            new TextRun({ text: "O contrato poderá ser prorrogado, através de emissão de Termo Aditivo, respeitado o limite máximo de 2 (dois) anos, na forma prevista no art. 11 da Lei nº. 11.788/2008.", font: "Calibri", size: 24 })
        ]
        }),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        children: [
            new TextRun({ text: "PARÁGRAFO SEGUNDO: ",  font: "Calibri", size: 24 }),
            new TextRun({ text: "Em caso do presente estágio ser prorrogado, o preenchimento e a assinatura do Termo Aditivo deverá ser providenciado antes da data de encerramento, contida na Cláusula Terceira neste Termo de Compromisso;", font: "Calibri", size: 24 })
        ]
        }),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 600 },
        children: [
            new TextRun({ text: "PARÁGRAFO SEGUNDO: ",  font: "Calibri", size: 24 }),
            new TextRun({ text: "Em período de recesso ou férias escolares, o estágio poderá ser realizado com carga horária de até 30 (trinta) horas semanais.", font: "Calibri", size: 24 })
        ]
        }),

        // CLÁUSULA TERCEIRA
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        children: [
            new TextRun({ text: "CLÁUSULA TERCEIRA", bold: true, font: "Calibri", size: 24 })
        ]
        }),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        children: [
            new TextRun({ text: "Fica compromissado entre as partes que:", font: "Calibri", size: 24 })
        ]
        }),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 100 },
        children: [
            new TextRun({
            text: "I - as atividades do estágio devem ser cumpridas em horário compatível com horário escolar do(a) ESTAGIÁRIO(A) e com o horário de funcionamento do CONCEDENTE, atendendo ao disposto no art. 10 da Lei nº. 11.788/2008;",
            font: "Calibri",
            size: 24
            })
        ]
        }),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 600 },
        children: [
            new TextRun({
            text: "II - nos períodos de férias escolares, a jornada de estágio será estabelecida de comum acordo entre o(a) ESTAGIÁRIO(A) e o(a) CONCEDENTE;",
            font: "Calibri",
            size: 24
            })
        ]
        }),

        // CLÁUSULA QUARTA
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        children: [
            new TextRun({ text: "CLÁUSULA QUARTA", bold: true, font: "Calibri", size: 24 })
        ]
        }),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        children: [
            new TextRun({ text: "Caberá ao CONCEDENTE:", font: "Calibri", size: 24 })
        ]
        }),
        ...[
        "I - apresentar um Plano de Estágio à INSTITUIÇÃO DE ENSINO;",
        "II - contratar em favor do estagiário seguro contra acidentes pessoais, cuja apólice seja compatível com valores de mercado;",
        "III - proporcionar ao(à) ESTAGIÁRIO(A) atividades de aprendizagem social, profissional e cultural compatíveis com sua formação profissional;",
        "IV - proporcionar ao(à) ESTAGIÁRIO(A) condições de treinamento prático e de relacionamento humano;",
        "V - designar funcionário, com formação ou experiência profissional na área de conhecimento desenvolvida no curso do estagiário para orientar as tarefas do(a) ESTAGIÁRIO(A);",
        "VI - enviar à INSTITUIÇÃO DE ENSINO, com periodicidade mínima de 6 (seis) meses, relatório de atividades, com vista obrigatória ao ESTAGIÁRIO(A);",
        "VII - fornecer relatório à INSTITUIÇÃO DE ENSINO, ao final do estágio, com as atividades desenvolvidas pelo(a) ESTAGIÁRIO(a) e a avaliação de desempenho;",
        "VIII - proporcionar ao ESTAGIÁRIO(A), sempre que o estágio tenha duração igual ou superior a 1 (um) ano, período de recesso de 30 (trinta) dias, a ser gozado preferencialmente em suas férias escolares, conforme o art. 13 da Lei nº. 11.788/2008;",
        "IX - velar pela assiduidade e pontualidade do ESTAGIÁRIO(A), fazendo o desconto proporcional das faltas não justificadas."
        ].map((texto) =>
        new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
            spacing: { after: 100 },
            children: [new TextRun({ text: texto, font: "Calibri", size: 24 })]
        })
        ),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        children: [
            new TextRun({ text: "PARÁGRAFO PRIMEIRO: ",  font: "Calibri", size: 24 }),
            new TextRun({
            text: "No caso de estágio obrigatório, a responsabilidade pela contratação do seguro de que trata o inciso II poderá, alternativamente, ser assumida pela instituição de ensino.",
            font: "Calibri",
            size: 24
            })
        ]
        }),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 600 },
        children: [
            new TextRun({ text: "PARÁGRAFO SEGUNDO: ",  font: "Calibri", size: 24 }),
            new TextRun({
            text: "O recesso de que trata o inciso VIII deverá ser remunerado quando o estágio receber bolsa ou outra forma de contraprestação, e os dias de recesso previstos serão concedidos de maneira proporcional, nos casos de o estágio ter duração inferior a 1 (um) ano.",
            font: "Calibri",
            size: 24
            })
        ]
        }),

        // CLÁUSULA QUINTA
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        children: [
            new TextRun({ text: "CLÁUSULA QUINTA", bold: true, font: "Calibri", size: 24 })
        ]
        }),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        children: [new TextRun({ text: "O(a) ESTAGIÁRIO(A) é obrigado a:", font: "Calibri", size: 24 })]
        }),
        ...[
        "a) estar regularmente matriculado(a) na INSTITUIÇÃO DE ENSINO, em semestre compatível com a prática exigida no estágio;",
        "b) observar as diretrizes e/ou normas internas do(a) CONCEDENTE e os dispositivos legais aplicáveis ao estágio, bem como as orientações do seu orientador e do seu supervisor;",
        "c) cumprir com seriedade e responsabilidade a programação estabelecida entre a CONCEDENTE, o(a) ESTAGIÁRIO(A) e a INSTITUIÇÃO DE ENSINO;",
        "d) ser assíduo e pontual no estágio;",
        "e) guardar segredo sobre decisões ou assuntos que tomar conhecimento em razão do estágio;",
        "f) comparecer às reuniões de discussão de estágio na INSTITUIÇÃO DE ENSINO;",
        "g) elaborar e entregar à INSTITUIÇÃO DE ENSINO relatórios periódicos e final sobre seu estágio, na forma por ela estabelecida;",
        "h) responder pelas perdas e danos consequentes da inobservância das cláusulas constantes do presente termo;"
        ].map(texto =>
        new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
            spacing: { after: 200 },
            children: [new TextRun({ text: texto, font: "Calibri", size: 24 })]
        })
        ),

        // CLÁUSULA SEXTA
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        spacing: { before: 600 },
        children: [new TextRun({ text: "CLÁUSULA SEXTA", bold: true, font: "Calibri", size: 24 })]
        }),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        children: [new TextRun({ text: "Cabe à INSTITUIÇÃO DE ENSINO:", font: "Calibri", size: 24 })]
        }),
        ...[
        "a) determinar um professor orientador, da área a ser desenvolvida no estágio, como responsável pelo acompanhamento e avaliação das atividades do estágio;",
        "b) planejar o estágio e orientar, supervisionar e avaliar, através do professor orientador, o(a) ESTAGIÁRIO(A);",
        "c) avaliar as instalações da parte concedente do estágio e sua adequação à formação cultural e profissional do educando;",
        "d) exigir do estagiário a apresentação periódica, em prazo não superior a 6 (seis) meses, de relatório das atividades;",
        "e) comunicar à parte concedente do estágio, no início do período letivo, as datas de realização de avaliações escolares ou acadêmicas;"
        ].map(texto =>
        new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
            spacing: { after: 200 },
            children: [new TextRun({ text: texto, font: "Calibri", size: 24 })]
        })
        ),

        // CLÁUSULA SÉTIMA
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        children: [new TextRun({ text: "CLÁUSULA SÉTIMA", bold: true, font: "Calibri", size: 24 })]
        }),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 600 },
        children: [
            new TextRun({
            text: "A realização de estágio deverá ser precedida da cobertura de seguro de acidentes pessoais em favor do estagiário, nos termos do Inciso IV e do parágrafo único do art. 9º da Lei nº. 11.788/2008.",
            font: "Calibri",
            size: 24
            })
        ]
        }),

        // CLÁUSULA OITAVA
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        children: [new TextRun({ text: "CLÁUSULA OITAVA", bold: true, font: "Calibri", size: 24 })]
        }),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        children: [
            new TextRun({
            text: "O estágio dar-se-á com remuneração (bolsa estágio), no valor de R$ 850,00 (OITOCENTOS E CINQUENTA REAIS), bem como auxílio - transporte.",
            font: "Calibri",
            size: 24
            })
        ]
        }),
        ...[
        ["PARÁGRAFO PRIMEIRO: ", "A concessão de auxílio - transporte ocorre mediante participação do estagiário no seu custeio."],
        ["PARÁGRAFO SEGUNDO: ", "Haverá o desconto proporcional na bolsa de estágio dos dias em que o ESTAGIÁRIO faltar, sem justificativa aceita pela dirigente do CONCEDENTE."],
        ["PARÁGRAFO TERCEIRO: ", "É vedado o pagamento de hora extra ou de qualquer tipo de gratificação ao estagiário."]
        ].map(([titulo, texto]) =>
        new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
            spacing: { after: 200 },
            children: [
            new TextRun({ text: titulo, font: "Calibri", size: 24 }),
            new TextRun({ text: texto, font: "Calibri", size: 24 })
            ]
        })
        ),

        // CLÁUSULA NONA
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        spacing: { before: 600 },
        children: [
            new TextRun({ text: "CLÁUSULA NONA", bold: true, font: "Calibri", size: 24 })
        ]
        }),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        children: [
            new TextRun({ text: "O TERMO DE COMPROMISSO será extinto:", font: "Calibri", size: 24 })
        ]
        }),
        ...[
        "a) automaticamente, ao final da sua vigência e no caso de conclusão, abandono ou a mudança de curso ou o trancamento de matrícula pelo ESTAGIÁRIO(A);",
        "b) no caso de não cumprimento do convencionado neste TERMO DE COMPROMISSO, bem como no Convênio do qual decorre;",
        "c) por determinação do CONCEDENTE ou solicitação do ESTAGIÁRIO(A);",
        "d) no caso de não comparecimento ao estágio, sem motivo justificado, por mais de 5 (cinco) dias no período de um mês ou por 30 (trinta) dias durante o estágio;",
        "e) acumular estágio em qualquer órgão ou entidade, pública ou particular."
        ].map((texto) =>
        new Paragraph({
            alignment: AlignmentType.JUSTIFIED,
            spacing: { after: 100 },
            children: [new TextRun({ text: texto, font: "Calibri", size: 24 })]
        })
        ),

        // CLÁUSULA DÉCIMA
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        spacing: { before: 600 },
        children: [
            new TextRun({ text: "CLÁUSULA DÉCIMA", bold: true, font: "Calibri", size: 24 })
        ]
        }),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 200 },
        children: [
            new TextRun({
            text: "Assim materializado e caracterizado, o presente estágio segundo a legislação, não acarretará vínculo empregatício de qualquer natureza, entre o(a) ESTAGIÁRIO(A) e o(a) CONCEDENTE, nos termos do que dispõe o art. 3º da Lei nº. 11.788/2008.",
            font: "Calibri",
            size: 24
            })
        ]
        }),
        new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        spacing: { after: 400 },
        children: [
            new TextRun({
            text: "E, por estarem de inteiro e comum acordo com as condições e dizeres deste instrumento, as partes assinam-no em meio digital via SEI-PI.",
            font: "Calibri",
            size: 24
            })
        ]
        }),

        // Espaçamento antes da assinatura
        new Paragraph({ spacing: { after: 500 } }),

        // 1. CONCEDENTE
        new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "______________________________________________", font: "Calibri", size: 24 })],
        spacing: { after: 100 }
        }),
        new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: nomeOrgao.toUpperCase(), font: "Calibri", size: 24 })],
        spacing: { after: 50 }
        }),
        new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: representanteOrgao.toUpperCase(), font: "Calibri", size: 24 })],
        spacing: { after: 50 }
        }),
        new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "CONCEDENTE", font: "Calibri", size: 24 })],
        spacing: { after: 200 }
        }),

        new Paragraph({ spacing: { after: 200 } }),

        // 2. INTERVENIENTE
        new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "______________________________________________", font: "Calibri", size: 24 })],
        spacing: { after: 100 }
        }),
        new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "SECRETARIA DA ADMINISTRAÇÃO DO ESTADO DO PIAUÍ", font: "Calibri", size: 24 })],
        spacing: { after: 50 }
        }),
        new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "SAMUEL PONTES DO NASCIMENTO", font: "Calibri", size: 24 })],
        spacing: { after: 50 }
        }),
        new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "INTERVENIENTE", font: "Calibri", size: 24 })],
        spacing: { after: 200 }
        }),

        new Paragraph({ spacing: { after: 200 } }),

        // 3. ESTAGIÁRIO
        new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "______________________________________________", font: "Calibri", size: 24 })],
        spacing: { after: 100 }
        }),
        new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: nomeEstagiario.toUpperCase(), font: "Calibri", size: 24 })],
        spacing: { after: 50 }
        }),
        new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "ESTAGIÁRIO(A)", font: "Calibri", size: 24 })],
        spacing: { after: 200 }
        }),

        new Paragraph({ spacing: { after: 200 } }),

        // 4. INSTITUIÇÃO
        new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "______________________________________________", font: "Calibri", size: 24 })],
        spacing: { after: 100 }
        }),
        new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: nomeInstituicao.toUpperCase(), font: "Calibri", size: 24 })],
        spacing: { after: 50 }
        }),
        new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "REPRESENTANTE LEGAL DA INSTITUIÇÃO", font: "Calibri", size: 24 })]
        }),



      ]
    }]
  });

  const buffer = await Packer.toBuffer(doc);
  const fileName = `Contrato_Estagio_${nomeEstagiario.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, "_")}.docx`;

  res.set({
    'Content-Disposition': `attachment; filename="${fileName}"`,
    'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });

  res.send(buffer);
});


app.listen(3000, () => {
    console.log("Servidor rodando em http://localhost:3000");
});
