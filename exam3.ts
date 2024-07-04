const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
import { TextRun } from 'docxtemplater'
// Đọc file template
const content = fs.readFileSync('Quyết định.docx', 'binary');

// Load content vào PizZip
const zip = new PizZip(content);
// Tạo instance của Docxtemplater với nội dung đã load
const doc = new Docxtemplater(zip, {
  paragraphLoop: true,
  linebreaks: true,
});

// Data để vá vào template
const data = {
  p_x: "Đầu máy",
};

// Set data vào tài liệu
doc.setData(data);

try {
  // Render tài liệu
  doc.render();
} catch (error) {
  console.error(error);
  throw error;
}

// Xuất tài liệu mới
const buf = doc.getZip().generate({ type: 'nodebuffer' });

// Ghi tài liệu mới vào file
fs.writeFileSync('output.docx', buf);
