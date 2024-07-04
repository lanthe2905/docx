const { DOMParser } = require("xmldom");
const xpath = require("xpath");
const JsZip = require("jszip");
const fs = require("fs");

//need to add declare:
let docxInputPath = "Quyết định.docx";
let strOutputPath = "Quyết định thay đổi.docx";

// Read the docx internal xdocument
let wSelect = xpath.useNamespaces({
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
});

let docxFile = fs.readFileSync(docxInputPath);
JsZip.loadAsync(docxFile).then(async (zip) => {
  await zip
    .file("word/document.xml")
    .async("string")
    .then((docx_str) => {
      let docx = new DOMParser().parseFromString(docx_str);
      let outputString = "";
      let paragraphElements = this.wSelect("//w:p", docx);
      paragraphElements.forEach((paragraphElement) => {
        let textElements = this.wSelect(".//w:t", paragraphElement);
        textElements.forEach(
          (textElement) => (outputString += textElement.textContent)
        );
        if (textElements.length > 0) outputString += "\n";
      });
      fs.writeFileSync(strOutputPath, outputString);
    });
});
