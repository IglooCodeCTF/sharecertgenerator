const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const inquirer = require('inquirer');

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

console.log("This utility will walk you through creating a share certificate for IglooCode Ltd.");
console.log("The generated share certificate must be printed and signed by two directors and a witness.")
console.log("")
console.log("Press ^C at any time to quit.")

inquirer.prompt([
    {
        type: 'input',
        name: 'name',
        message: 'Shareholder name:',
        default: require("os").userInfo().username
    },
    {
        type: 'input',
        name: 'address',
        message: 'Shareholder address:',
    },
    {
        type: 'input',
        name: 'quantity',
        message: 'Number of shares:',
        validate(value) {
          const valid = !isNaN(parseFloat(value));
          return valid || 'Please enter a number';
        },
        filter: Number,
      },
      {
        type: 'list',
        name: 'class',
        message: 'Share class:',
        choices: [
            {
                key: 'o',
                name: 'Ordinary',
                value: 'ORDINARY'
            },
            {
                key: 'p',
                name: 'Preference',
                value: 'PREFERENCE'
            },
            {
                key: 'd',
                name: 'Deferred',
                value: 'DEFERRED'
            },
            {
                key: 'r',
                name: 'Redeemable',
                value: 'REDEEMABLE'
            }
        ],
        default: 'ORDINARY'
      },
      {
        type: 'input',
        name: 'no',
        message: 'Certificate no:',
        validate(value) {
          const valid = !isNaN(parseFloat(value));
          return valid || 'Please enter a number';
        },
        filter: Number,
      },
]).then(a => {

    doc.render({
        shareholder_name_lc: a.name,
        shareholder_name_uc: a.name.toUpperCase(),
        shareholder_address_uc: a.address.toUpperCase(),
        date: new Date().toLocaleString('en-GB', { dateStyle: "short" }),
        no: String(a.no),
        quantity: String(a.quantity),
        class: a.class,
    });

    const buf = doc.getZip().generate({
        type: "nodebuffer",
        compression: "DEFLATE",
    });

    const rnd = (Math.random() + 1).toString(36).substring(7);
    
    fs.writeFileSync(path.resolve(require("os").tmpdir(), `${rnd}.docx`), buf);

    var exec = require('child_process').exec;
    exec(getCommandLine() + ' ' + path.resolve(require("os").tmpdir(), `${rnd}.docx`));

})

function getCommandLine() {
    switch (process.platform) { 
       case 'darwin' : return 'open';
       case 'win32' : return 'start';
       case 'win64' : return 'start';
       default : return 'xdg-open';
    }
 }