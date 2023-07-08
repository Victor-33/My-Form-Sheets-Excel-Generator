const { MongoClient, ObjectId } = require('mongodb');
const ExcelJS = require('exceljs');
const moment = require('moment');

async function generateExcel(formId) {
  try {
    const uri =
      'mongodb+srv://my-form-sheets:(KEY_CODE)@my-form-sheets.3herknk.mongodb.net/MyFormSheets?retryWrites=true&w=majority';
    const client = new MongoClient(uri);
    await client.connect();

    const collection = client.db('MyFormSheets').collection('Form');
    const form = await collection.findOne({ _id: new ObjectId(formId) });

    if (!form) {
      console.log('Formulário não encontrado');
      return;
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Respostas');

    const headerStyle = {
      font: { bold: false, size: 12, color: { argb: 'FFFFFFFF' } },
      alignment: { horizontal: 'left' },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF005536' } },
      border: {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } },
      },
    };

    const dataStyle = {
      alignment: { horizontal: 'center' },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } },
      border: {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } },
      },
    };

    const headers = Object.keys(form.answers[0]);
    const rowStartIndex = 7;

    headers.forEach((header, index) => {
      const questionCell = worksheet.getCell(`B${rowStartIndex + index}`);
      questionCell.value = header;
      questionCell.alignment = headerStyle.alignment;
      questionCell.fill = headerStyle.fill;
      questionCell.border = headerStyle.border;

      const answerCell = worksheet.getCell(`C${rowStartIndex + index}`);
      answerCell.font = dataStyle.font;
      answerCell.alignment = dataStyle.alignment;
      answerCell.fill = dataStyle.fill;
      answerCell.border = dataStyle.border;
    });

    let respostaCount = 0;

    for (let i = 0; i < form.answers.length; i++) {
      const answer = form.answers[i];
      const rowStartIndex = 7 + i * (headers.length + 1);

      headers.forEach((header, index) => {
        const questionCell = worksheet.getCell(`B${rowStartIndex + index}`);
        const answerCell = worksheet.getCell(`C${rowStartIndex + index}`);

        questionCell.value = header;
        answerCell.value = answer[header];

        questionCell.alignment = dataStyle.alignment;
        questionCell.fill = dataStyle.fill;
        questionCell.border = dataStyle.border;

        answerCell.alignment = dataStyle.alignment;
        answerCell.fill = dataStyle.fill;
        answerCell.border = dataStyle.border;
      });

      worksheet.getCell('B2').value = (`${respostaCount + 1} Responses`)
      worksheet.getCell('B2').border = dataStyle.border;
      worksheet.getCell('B2').alignment = dataStyle.alignment;
  
      const name = form.name;
      const date = form.date;
      const time = form.time;
      const currentDate = new Date(date);
      const creationDate = new ObjectId(form._id).getTimestamp();
      
      currentDate.setDate(currentDate.getDate());
      const formattedDate = currentDate.toISOString().split('T')[0];
      
      const formattedCreationDate = `${creationDate.getFullYear()}-${(creationDate.getMonth() + 1).toString().padStart(2, '0')}-${creationDate.getDate().toString().padStart(2, '0')} ${creationDate.getHours().toString().padStart(2, '0')}:${creationDate.getMinutes().toString().padStart(2, '0')}`;
      
      worksheet.getCell('C2').value = `Creation Date: ${formattedCreationDate} (GMT-3)`;
      worksheet.getCell('C2').border = dataStyle.border;
      worksheet.getCell('C2').alignment = dataStyle.alignment;
      
      worksheet.getCell('D2').value = `End Date: ${formattedDate} ${time} (GMT-3)`;
      worksheet.getCell('D2').border = dataStyle.border;
      worksheet.getCell('D2').alignment = dataStyle.alignment;
  
      worksheet.getCell('E2').value = ' ';
  
      worksheet.getCell(`C4`).value = 'My Forms Sheets: Response Report - ' + name;
      worksheet.getCell(`C4`).fill = headerStyle.fill;
      worksheet.getCell(`C4`).font = headerStyle.font;
      worksheet.getCell(`C4`).alignment = headerStyle.alignment;
      worksheet.getCell('C4').font = {bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
      

      worksheet.getCell('B6').value = 'Questions';
      worksheet.getCell('B6').font = headerStyle.font;
      worksheet.getCell('B6').alignment = headerStyle.alignment;
      worksheet.getCell('B6').fill = headerStyle.fill;
      
  
      worksheet.getCell('C6').value = 'Answers';
      worksheet.getCell('C6').font = headerStyle.font;
      worksheet.getCell('C6').alignment = headerStyle.alignment;
      worksheet.getCell('C6').fill = headerStyle.fill;
      worksheet.getCell('C6').border = dataStyle.border;

      worksheet.getCell(`B${rowStartIndex + headers.length}`).value = worksheet.getCell(`C${rowStartIndex + headers.length}`).value;
      worksheet.getCell(`B${rowStartIndex + headers.length}`).font = dataStyle.font;
      worksheet.getCell(`B${rowStartIndex + headers.length}`).alignment = dataStyle.alignment;
      worksheet.getCell(`B${rowStartIndex + headers.length}`).fill = dataStyle.fill;
      worksheet.getCell(`B${rowStartIndex + headers.length}`).border = dataStyle.border;

      worksheet.getCell(`C${rowStartIndex + headers.length}`).value = null;
      worksheet.getCell(`C${rowStartIndex + headers.length}`).font = null;
      worksheet.getCell(`C${rowStartIndex + headers.length}`).alignment = null;
      worksheet.getCell(`C${rowStartIndex + headers.length}`).fill = null;
      worksheet.getCell(`C${rowStartIndex + headers.length}`).border = null;

      respostaCount++;
    }

    worksheet.getColumn('B').width = 15;
    worksheet.getColumn('B').alignment = headerStyle.alignment;
    worksheet.getColumn('C').width = 75;
    worksheet.getColumn('C').alignment = dataStyle.alignment;
    worksheet.getColumn('D').width = 45;
    worksheet.getColumn('D').alignment = dataStyle.alignment;

    const filePath = 'respostas.xlsx';
    await workbook.xlsx.writeFile(filePath);

    // ** here you can define to send the file to email or save it to a cloud storage ** //

    console.log(`Foram salvas ${respostaCount} respostas.`);

    // ** //

    await client.close();
  } catch (error) {
    console.error('Erro ao gerar o arquivo Excel:', error);
  }
}

generateExcel();

module.exports = generateExcel;