const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const fs = require("fs");
const path = require("path");

const content = fs.readFileSync(
    path.resolve(__dirname, "template.docx"),
    "binary"
);

const zip = new PizZip(content);

const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
});

doc.render({
    shareholder_name_lc: "Daniel Adams",
    shareholder_name_uc: "DANIEL ADAMS",
    shareholder_address_uc: "103 LISNAMURRIKIN ROAD, BROUGHSHANE, BALLYMENA, NORTHERN IRELAND, BT42 4PP",
    date: new Date().toLocaleString('en-GB', { dateStyle: "short" }),
    no: String(-1),
    quantity: String(-1),
    class: "ORDINARY",
});

const buf = doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
});

fs.writeFileSync(path.resolve(__dirname, "output.docx"), buf);