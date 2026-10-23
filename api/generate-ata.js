const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        Header, Footer, AlignmentType, LevelFormat, BorderStyle, WidthType, 
        ShadingType, PageNumber, VerticalAlign } = require('docx');
const AZUL_ZURI = "2E75B6";
const AZUL_ESCURO = "1F4E79";
const CINZA = "595959";
const CINZA_CLARO = "F2F2F2";
const border = { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const headerBorder = { style: BorderStyle.SINGLE, size: 4, color: AZUL_ZURI };
const headerBorders = { top: headerBorder, bottom: headerBorder, left: headerBorder, right: headerBorder };
function createHeaderCell(text, width) {
  return new TableCell({
    borders: headerBorders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: AZUL_ZURI, type: ShadingType.CLEAR },
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      children: [new TextRun({ text: text || "", bold: true, color: "FFFFFF", font: "Arial", size: 18 })]
    })]
  });
}
function createCell(text, width, options = {}) {
  const { bold = false, fill = null, align = AlignmentType.LEFT } = options;
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: fill ? { fill, type: ShadingType.CLEAR } : undefined,
    margins: { top: 50, bottom: 50, left: 100, right: 100 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text: text || "", bold, font: "Arial", size: 18, color: CINZA })]
    })]
  });
}
function createParagraph(text, options = {}) {
  const { bold = false, size = 20, color = "000000", spacingAfter = 80, alignment = AlignmentType.LEFT } = options;
  return new Paragraph({
    alignment,
    spacing: { after: spacingAfter },
    children: [new TextRun({ text: text || "", bold, font: "Arial", size, color })]
  });
}
module.exports = async function (req, res) {
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
    const data = typeof req.body === 'string' ? JSON.parse(req.body) : (req.body || {});
    const projeto = data.projeto || {};
    const reuniao = data.reuniao || {};
    const participantes = data.participantes || { zuri: [], cliente: [] };
    const secoes = data.secoes || [];
    const decisoes = data.decisoes || [];
    const acoes = data.acoes || [];
    const elaboradoPor = data.elaborado_por || "Roberto Augusto Rosa – Gerente de Projetos (Zuri)";
    const observacoes = data.observacoes_finais || "Esta ata reflete o conteúdo apresentado e discutido na reunião.";
    const tipoReuniao = (reuniao.tipo || "Reunião").toUpperCase();
    const children = [];
    // Título dinâmico
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 60 },
      children: [new TextRun({ text: `ATA DE REUNIÃO DE ${tipoReuniao}`, bold: true, font: "Arial", size: 32, color: AZUL_ESCURO })]
    }));
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [new TextRun({ 
        text: `Projeto ${projeto.codigo || ""} – ${projeto.nome || ""}`, 
        bold: true, font: "Arial", size: 22, color: AZUL_ZURI 
      })]
    }));
    // Tabela de identificação
    children.push(new Table({
      width: { size: 10206, type: WidthType.DXA },
      columnWidths: [2550, 7656],
      rows: [
        new TableRow({ children: [
          createCell("Data / Hora", 2550, { bold: true, fill: CINZA_CLARO }),
          createCell(`${reuniao.data || ""} | ${reuniao.horario || ""} | Duração: ${reuniao.duracao || ""}`, 7656)
        ]}),
        new TableRow({ children: [
          createCell("Meio", 2550, { bold: true, fill: CINZA_CLARO }),
          createCell(reuniao.meio || "Microsoft Teams (remoto)", 7656)
        ]}),
        new TableRow({ children: [
          createCell("Projeto", 2550, { bold: true, fill: CINZA_CLARO }),
          createCell(`${projeto.codigo || ""} – ${projeto.nome || ""}`, 7656)
        ]}),
        new TableRow({ children: [
          createCell("Tipo de reunião", 2550, { bold: true, fill: CINZA_CLARO }),
          createCell(reuniao.tipo || "Kick-off", 7656)
        ]}),
        new TableRow({ children: [
          createCell("Elaborado por", 2550, { bold: true, fill: CINZA_CLARO }),
          createCell(elaboradoPor, 7656)
        ]}),
      ]
    }));
    // Participantes
    children.push(new Paragraph({
      spacing: { before: 280, after: 120 },
      children: [new TextRun({ text: "1. Participantes", bold: true, font: "Arial", size: 28, color: AZUL_ESCURO })]
    }));
    if (participantes.zuri && participantes.zuri.length > 0) {
      children.push(createParagraph("Zuri Smart Flows", { bold: true, size: 20, color: AZUL_ZURI, spacingAfter: 60 }));
      participantes.zuri.forEach(p => {
        children.push(new Paragraph({
          spacing: { after: 40 },
          indent: { left: 360 },
          children: [new TextRun({ text: `• ${p.nome}${p.papel ? " – " + p.papel : ""}`, font: "Arial", size: 20 })]
        }));
      });
    }
    if (participantes.cliente && participantes.cliente.length > 0) {
      children.push(createParagraph(projeto.cliente || "Cliente", { bold: true, size: 20, color: AZUL_ZURI, spacingAfter: 60 }));
      participantes.cliente.forEach(p => {
        children.push(new Paragraph({
          spacing: { after: 40 },
          indent: { left: 360 },
          children: [new TextRun({ text: `• ${p.nome}${p.papel ? " – " + p.papel : ""}`, font: "Arial", size: 20 })]
        }));
      });
    }
    // Seções dinâmicas
    let secaoNum = 2;
    secoes.forEach(secao => {
      children.push(new Paragraph({
        spacing: { before: 280, after: 120 },
        children: [new TextRun({ text: `${secaoNum}. ${secao.titulo || "Seção"}`, bold: true, font: "Arial", size: 28, color: AZUL_ESCURO })]
      }));
      if (secao.tipo === "tabela" && secao.colunas && secao.linhas) {
        const colCount = secao.colunas.length;
        const colWidth = Math.floor(10206 / colCount);
        const columnWidths = Array(colCount).fill(colWidth);
        const headerRow = new TableRow({
          children: secao.colunas.map((col, i) => createHeaderCell(col, columnWidths[i]))
        });
        const dataRows = secao.linhas.map(linha => {
          return new TableRow({
            children: (Array.isArray(linha) ? linha : Object.values(linha)).map((cel, i) => 
              createCell(String(cel || ""), columnWidths[i] || colWidth)
            )
          });
        });
        children.push(new Table({
          width: { size: 10206, type: WidthType.DXA },
          columnWidths,
          rows: [headerRow, ...dataRows]
        }));
      } else if (secao.conteudo) {
        if (Array.isArray(secao.conteudo)) {
          secao.conteudo.forEach(item => {
            children.push(new Paragraph({
              spacing: { after: 60 },
              indent: { left: 360 },
              children: [new TextRun({ text: `• ${item}`, font: "Arial", size: 20 })]
            }));
          });
        } else {
          children.push(createParagraph(secao.conteudo, { size: 20, spacingAfter: 80 }));
        }
      }
      secaoNum++;
    });
    // Decisões
    if (decisoes.length > 0) {
      children.push(new Paragraph({
        spacing: { before: 280, after: 120 },
        children: [new TextRun({ text: `${secaoNum}. Decisões`, bold: true, font: "Arial", size: 28, color: AZUL_ESCURO })]
      }));
      decisoes.forEach(d => {
        children.push(new Paragraph({
          spacing: { after: 60 },
          indent: { left: 360 },
          children: [new TextRun({ text: `• ${d}`, font: "Arial", size: 20 })]
        }));
      });
      secaoNum++;
    }
    // Ações
    if (acoes.length > 0) {
      children.push(new Paragraph({
        spacing: { before: 280, after: 120 },
        children: [new TextRun({ text: `${secaoNum}. Ações / Próximos passos`, bold: true, font: "Arial", size: 28, color: AZUL_ESCURO })]
      }));
      const acoesHeader = new TableRow({
        children: [
          createHeaderCell("#", 600),
          createHeaderCell("Ação", 5000),
          createHeaderCell("Responsável", 2400),
          createHeaderCell("Prazo", 2206)
        ]
      });
      const acoesRows = acoes.map((a, idx) => new TableRow({
        children: [
          createCell(String(a.numero || (idx + 1)), 600, { align: AlignmentType.CENTER }),
          createCell(a.acao || "", 5000),
          createCell(a.responsavel || "", 2400),
          createCell(a.prazo || "", 2206)
        ]
      }));
      children.push(new Table({
        width: { size: 10206, type: WidthType.DXA },
        columnWidths: [600, 5000, 2400, 2206],
        rows: [acoesHeader, ...acoesRows]
      }));
    }
    // Observações finais
    children.push(new Paragraph({
      spacing: { before: 300 },
      children: [new TextRun({ text: observacoes, font: "Arial", size: 18, italics: true, color: CINZA })]
    }));
    children.push(new Paragraph({
      spacing: { before: 200 },
      children: [new TextRun({ text: `Elaborada por: ${elaboradoPor}`, font: "Arial", size: 20 })]
    }));
    // Documento final
    const doc = new Document({
      styles: {
        default: { document: { run: { font: "Arial", size: 20 } } }
      },
      sections: [{
        properties: {
          page: {
            size: { width: 11906, height: 16838 },
            margin: { top: 850, right: 850, bottom: 850, left: 850 }
          }
        },
        headers: {
          default: new Header({
            children: [
              new Paragraph({
                border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: AZUL_ZURI, space: 8 } },
                spacing: { after: 120 },
                children: [
                  new TextRun({ text: "ZURI SMART FLOWS", bold: true, font: "Arial", size: 18, color: AZUL_ZURI }),
                  new TextRun({ text: `  |  Ata de Reunião – ${reuniao.tipo || "Reunião"}`, font: "Arial", size: 18, color: CINZA })
                ]
              })
            ]
          })
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                border: { top: { style: BorderStyle.SINGLE, size: 6, color: "CCCCCC", space: 6 } },
                spacing: { before: 80 },
                children: [
                  new TextRun({ text: `Projeto ${projeto.codigo || ""} | Confidencial  |  Página `, font: "Arial", size: 16, color: CINZA }),
                  new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 16, color: CINZA }),
                  new TextRun({ text: " de ", font: "Arial", size: 16, color: CINZA }),
                  new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Arial", size: 16, color: CINZA })
                ]
              })
            ]
          })
        },
        children
      }]
    });
    const buffer = await Packer.toBuffer(doc);
    const fileName = `Ata_${projeto.codigo || "Projeto"}_${(reuniao.tipo || "Reuniao").replace(/\s+/g, "_")}.docx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
    res.send(buffer);
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: 'Erro ao gerar o documento', details: error.message });
  }
};
