const fs = require("fs");
const { Document, Packer, Paragraph, TextRun } = require("docx");

const doc = new Document({
    sections: [{
        children: [
            new Paragraph({
                children: [
                    new TextRun("Documento de Teste"),
                ],
            }),
        ],
    }],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("teste.docx", buffer);
    console.log("Documento criado com sucesso.");
});
