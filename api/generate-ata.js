const { Document, Packer, Paragraph, TextRun, Header, Footer, PageNumber, AlignmentType, BorderStyle } = require('docx');

module.exports = async function (req, res) {
  // Permite CORS (importante para o Power Automate chamar depois)
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Método não permitido. Use POST.' });
  }

  try {
    // Por enquanto vamos gerar uma ata de teste
    const doc = new Document({
      styles: {
        default: {
          document: {
            run: { font: "Arial", size: 20 }
          }
        }
      },
      sections: [{
        properties: {
          page: {
            size: { width: 11906, height: 16838 }, // A4
            margin: { top: 850, right: 850, bottom: 850, left: 850 }
          }
        },
        headers: {
          default: new Header({
            children: [
              new Paragraph({
                border: {
                  bottom: { style: BorderStyle.SINGLE, size: 12, color: "2E75B6", space: 8 }
                },
                children: [
                  new TextRun({ text: "ZURI SMART FLOWS", bold: true, color: "2E75B6", size: 18 }),
                  new TextRun({ text: "  |  Ata de Reunião – Kick-off", color: "595959", size: 18 })
                ]
              })
            ]
          })
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: "Projeto Teste | Confidencial  |  Página ", size: 16, color: "595959" }),
                  new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "595959" })
                ]
              })
            ]
          })
        },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 },
            children: [
              new TextRun({
                text: "ATA DE REUNIÃO DE KICK-OFF",
                bold: true,
                size: 32,
                color: "1F4E79"
              })
            ]
          }),
          new Paragraph({
            spacing: { after: 120 },
            children: [
              new TextRun({ text: "Este é um documento de teste gerado pela função Node.js.", size: 22 })
            ]
          }),
          new Paragraph({
            children: [
              new TextRun({ text: "Se você está vendo este arquivo, o endpoint está funcionando corretamente.", size: 20 })
            ]
          })
        ]
      }]
    });

    const buffer = await Packer.toBuffer(doc);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename=Ata_Teste_Zuri.docx');
    res.send(buffer);

  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Erro ao gerar o documento', details: error.message });
  }
};
