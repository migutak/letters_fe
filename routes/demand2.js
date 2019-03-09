var express = require('express');
var router = express.Router();
const app = express();
const path = require('path');
const docx = require('docx');
const fs = require('fs');
var numeral = require('numeral');
const bodyParser = require("body-parser");
var dateFormat = require('dateformat');
const word2pdf = require('word2pdf-promises');

var data = require('./data.js');

const LETTERS_DIR = data.filePath;

const { Document, Paragraph, Packer, TextRun } = docx;

router.use(bodyParser.urlencoded({
    extended: true
}));

router.use(bodyParser.json());

router.post('/download', function (req, res) {
    const letter_data = req.body;
    const GURARANTORS = req.body.guarantors;
    const INCLUDELOGO = req.body.showlogo;
    const DATA = req.body.accounts;
    const DATE = dateFormat(new Date(), "isoDate");
    //
    //
    const document = new Document();
    if (INCLUDELOGO == 'Y') {
        const footer1 = new TextRun("Directors: John Murugu (Chairman), Dr. Gideon Muriuki (Group Managing Director & CEO), M. Malonza (Vice Chairman),")
            .size(16)
        const parafooter1 = new Paragraph()
        parafooter1.addRun(footer1).center();
        document.Footer.addParagraph(parafooter1);
        const footer2 = new TextRun("J. Sitienei, B. Simiyu, P. Githendu, W. Ongoro, R. Kimanthi, W. Mwambia, R. Simani (Mrs), L. Karissa, G. Mburia.")
            .size(16)
        const parafooter2 = new Paragraph()
        parafooter1.addRun(footer2).center();
        document.Footer.addParagraph(parafooter2);

        //logo start

        document.createImage(fs.readFileSync("./coop.jpg"), 350, 60, {
            floating: {
                horizontalPosition: {
                    offset: 1000000,
                },
                verticalPosition: {
                    offset: 1014400,
                },
                margins: {
                    top: 0,
                    bottom: 201440,
                },
            },
        });
    }
    // logo end

    document.createParagraph("The Co-operative Bank of Kenya Limited").right();
    document.createParagraph("Co-operative Bank House").right();
    document.createParagraph("Haile Selassie Avenue").right();
    document.createParagraph("P.O.Box 48231-00100 GPO, Nairobi").right();
    document.createParagraph("Tel: (020) 3276100").right();
    document.createParagraph("Fax: (020) 2227747/2219831").right();

    document.createParagraph(" ");

    const text = new TextRun(" ''Without Prejudice'' ")
    const paragraph = new Paragraph();
    text.bold();
    // document.createParagraph(" ''Without Prejudice'' ").title().heading3();
    paragraph.addRun(text).center();
    document.addParagraph(paragraph);

    document.createParagraph(" ");

    document.createParagraph("Our Ref: DEMAND2/"+ letter_data.branchcode + '/' + letter_data.arocode +'/'+ DATE);
    document.createParagraph(" ");
    const ddate = new TextRun(dateFormat(new Date(), 'fullDate'));
    const pddate = new Paragraph();
    ddate.size(20);
    pddate.addRun(ddate);
    document.addParagraph(pddate);

    document.createParagraph(" ");
    document.createParagraph(letter_data.custname);
    document.createParagraph(letter_data.address);
    document.createParagraph(letter_data.custname);
    document.createParagraph(" ");

    document.createParagraph("Dear sir/madam ");
    document.createParagraph(" ");

    const headertext = new TextRun("RE: OUTSTANDING LIABILITIES A/C NO. "+ letter_data.acc +" - "+ letter_data.custname +" ");
    const paragraphheadertext = new Paragraph();
    headertext.bold();
    headertext.underline();
    paragraphheadertext.addRun(headertext);
    document.addParagraph(paragraphheadertext);

    document.createParagraph(" ");
    document.createParagraph("Following our 1st notice (dated), we note with concern that your account/s is/are still in arrears/overdrawn ");
    document.createParagraph("Kindly note that your current balance/s as indicated below and it/they continue/s to accrue interest until payment is made in full.  ");
    document.createParagraph(" ");

    const table = document.createTable(DATA.length + 2, 7);
    /*float({
      horizontalAnchor: TableAnchorType.MARGIN,
      verticalAnchor: TableAnchorType.MARGIN,
      relativeHorizontalPosition: RelativeHorizontalPosition.RIGHT,
      relativeVerticalPosition: RelativeVerticalPosition.BOTTOM,
  });*/
    // table.setFixedWidthLayout();
    // table.setWidth('45', WidthType.DXA);
    table.getCell(0, 1).addContent(new Paragraph("Account no"));
    table.getCell(0, 2).addContent(new Paragraph("Principal Loan"));
    table.getCell(0, 3).addContent(new Paragraph("Outstanding Interest "));
    table.getCell(0, 4).addContent(new Paragraph("Principal Arrears"));
    table.getCell(0, 5).addContent(new Paragraph("Total Arrears"));
    table.getCell(0, 6).addContent(
        new Paragraph("Total Outstanding")
    );
    // table rows
    for (i = 0; i < DATA.length; i++) {
        row = i + 1
        table.getCell(row, 1).addContent(new Paragraph(DATA[i].accnumber));
        table.getCell(row, 2).addContent(new Paragraph(numeral(DATA[i].oustbalance).format('0,0.00')));
        table.getCell(row, 3).addContent(new Paragraph(numeral(DATA[i].princarrears).format('0,0.00')));
        table.getCell(row, 4).addContent(new Paragraph(numeral(DATA[i].intarrears).format('0,0.00')));
        table.getCell(row, 5).addContent(new Paragraph(numeral(DATA[i].totalarrears).format('0,0.00')));
        table.getCell(row, 6).addContent(new Paragraph(numeral(DATA[i].oustbalance + DATA[i].totalarrears).format('0,0.00')));
    }

    document.createParagraph(" ");

    document.createParagraph(" ");
    document.createParagraph("We DEMAND that you pay the amount in arrears, plus the accrued interest within fourteen days (14) from the date hereof. ");


    document.createParagraph(" ");
    const txt = new TextRun("Kindly also note that under the provisions of the Banking (Credit Reference Bureau) Regulations 2013, it is now a mandatory requirement in law that all financial institutions share positive and negative credit information while assessing customers credit worthiness, standing and capacity through duly licensed Credit Reference Bureaus (CRBs) for inclusion and maintenance in their database for purposes of sharing the said information..");
    const ptxt = new Paragraph();
    txt.size(20);
    ptxt.addRun(txt);
    ptxt.justified();
    document.addParagraph(ptxt);

    document.createParagraph(" ");
    const txt2 = new TextRun("We would therefore as a matter of courtesy like to notify you that unless you fully settle all your outstanding arrears with the Bank from the date stated above we shall proceed to adversely update your details and information with the CRBs relating to your credit worthiness and standing.  ");
    const ptxt2 = new Paragraph();
    txt2.size(20);
    ptxt2.addRun(txt2);
    ptxt2.justified();
    document.addParagraph(ptxt2);

    document.createParagraph(" ");
    const txt3 = new TextRun("In the event that you require any clarification or information, you may contact the undersigned on Telephone number 0203276000/ 0711049000/0732106000 ");
    const ptxt3 = new Paragraph();
    txt3.size(20);
    ptxt3.addRun(txt3);
    ptxt3.justified();
    document.addParagraph(ptxt3);

    document.createParagraph(" ");
    document.createParagraph("Yours Faithfully, ");

    document.createParagraph(" ");
    document.createParagraph( letter_data.manager );
    document.createParagraph("BRANCH MANAGER ");
    document.createParagraph( letter_data.branchname + " BRANCH");


    if (GURARANTORS) {
        document.createParagraph("cc: ");

        for (g = 0; g < GURARANTORS.length; g++) {
            document.createParagraph(" ");
            document.createParagraph( GURARANTORS[g].name );
            document.createParagraph( GURARANTORS[g].address );
        }
    }

    document.createParagraph(" ");
    document.createParagraph("This letter is valid without a signature ");

    const packer = new Packer();

    packer.toBuffer(document).then((buffer) => {
        fs.writeFileSync(LETTERS_DIR + letter_data.acc + DATE + "demand2.docx", buffer);
        //conver to pdf
        // if pdf format
        if(letter_data.format == 'pdf'){
          const convert = () => {
            word2pdf.word2pdf(LETTERS_DIR + letter_data.acc + DATE + "demand2.docx")
              .then(data => {
                fs.writeFileSync(LETTERS_DIR+ letter_data.acc + DATE + 'demand2.pdf', data);
                res.json({result: 'success', message: LETTERS_DIR + letter_data.acc + DATE + "demand2.pdf"})
              }, error  => {
                console.log('error ...', error)
                res.json({result: 'error', message: 'Exception occured'});
              })
          }
          convert();
        } else {
          // res.sendFile(path.join(LETTERS_DIR + letter_data.acc + DATE + 'prelistingcc.docx'));
          res.json({result: 'success', message: LETTERS_DIR + letter_data.acc + DATE + "demand2.docx"})
        }
      }).catch((err) => {
        console.log(err);
        res.json({result: 'error', message: 'Exception occured'});
      });
});

module.exports = router;