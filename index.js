const fs = require('fs');
const XLSX = require('xlsx');
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');

// Função para ler a planilha
function readExcel(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(worksheet);
}

// Função para preencher e salvar o PDF
async function fillPdfTemplate(participant, pdfTemplatePath, outputPath) {
    const existingPdfBytes = fs.readFileSync(pdfTemplatePath);
    const pdfDoc = await PDFDocument.load(existingPdfBytes);
    const form = pdfDoc.getForm();

    // Função auxiliar para converter valores para string
    const toString = (value) => value != null ? String(value) : '';

    // Preencher os campos do PDF
    form.getTextField('Text1').setText(toString(participant.nome));
    form.getTextField('Text2').setText(toString(participant.tipoInscricao));
    form.getTextField('Text3').setText(toString(participant.dataNascimento));
    form.getTextField('Text4').setText(toString(participant.rg));
    form.getTextField('Text5').setText(toString(participant.cpf));
    form.getTextField('Text6').setText(toString(participant.logradouro));
    form.getTextField('Text7').setText(toString(participant.numero));
    form.getTextField('Text8').setText(toString(participant.bairro));
    form.getTextField('Text9').setText(toString(participant.cidade));
    form.getTextField('Text10').setText(toString(participant.cep));
    form.getTextField('Text11').setText(toString(participant.telefone));

    const pdfBytes = await pdfDoc.save();
    fs.writeFileSync(outputPath, pdfBytes);
}

// Caminhos dos arquivos
const pdfTemplatePath = 'ficha.pdf';
const outputDir = 'pdfs/';
if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
}

var excelFilePath = 'promocional.xlsx';
var participants = readExcel(excelFilePath);
participants.forEach(async (participant) => {
    const outputPath = `${outputDir}/${participant.nome}.pdf`;
    await fillPdfTemplate(participant, pdfTemplatePath, outputPath);
    console.log(`PDF gerado para: ${participant.nome}`);
});

var excelFilePath = 'parcelado.xlsx';
var participants = readExcel(excelFilePath);
participants.forEach(async (participant) => {
    const outputPath = `${outputDir}/${participant.nome}.pdf`;
    await fillPdfTemplate(participant, pdfTemplatePath, outputPath);
    console.log(`PDF gerado para: ${participant.nome}`);
});

var excelFilePath = 'vista.xlsx';
var participants = readExcel(excelFilePath);
participants.forEach(async (participant) => {
    const outputPath = `${outputDir}/${participant.nome}.pdf`;
    await fillPdfTemplate(participant, pdfTemplatePath, outputPath);
    console.log(`PDF gerado para: ${participant.nome}`);
});