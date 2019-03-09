var express = require('express');
var router = express.Router();
const app = express();
const path = require('path');
const docx = require('docx');
const fs = require('fs');
var numeral = require('numeral');
var dateFormat = require('dateformat');
const bodyParser = require("body-parser");
const word2pdf = require('word2pdf-promises');
// const word2pdf = require('word2pdf');
var data = require('./data.js');

const LETTERS_DIR = data.filePath;

const { Document, Paragraph, Packer, TextRun, BorderStyle, Borders } = docx;

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

  const ref = new TextRun("Our Ref: PRELISTING/" + letter_data.cardacct + '/' + DATE);
  const paragraphref = new Paragraph();
  ref.bold();
  ref.size(24);
  paragraphref.addRun(ref);
  document.addParagraph(paragraphref);

  document.createParagraph(" ");
  const ddate = new TextRun(dateFormat(new Date(), 'fullDate'));
  const pddate = new Paragraph();
  ddate.size(24);
  pddate.addRun(ddate);
  document.addParagraph(pddate);

  document.createParagraph(" ");

  const name = new TextRun(letter_data.cardname);
  const pname = new Paragraph();
  name.size(20);
  pname.addRun(name);
  document.addParagraph(pname);

  const address = new TextRun(letter_data.address + '- ' + letter_data.rpcode);
  const paddress = new Paragraph();
  address.size(20);
  paddress.addRun(address);
  document.addParagraph(paddress);

  const city = new TextRun(letter_data.city);
  const pcity = new Paragraph();
  city.size(20);
  pcity.addRun(city);
  document.addParagraph(pcity);
  document.createParagraph(" ");

  const dear = new TextRun("Dear sir/madam ");
  const pdear = new Paragraph();
  dear.size(20);
  pdear.addRun(dear);
  document.addParagraph(pdear);
  document.createParagraph(" ");

  const headertext = new TextRun("PRE-LISTING NOTIFICATION ISSUED PURSUANT TO REGULATION 50(1)(a) OF THE CREDIT REFERENCE BUREAU REGULATIONS, 2013");
  const paragraphheadertext = new Paragraph();
  headertext.bold();
  headertext.size(20);
  headertext.underline();
  paragraphheadertext.addRun(headertext);
  paragraphheadertext.center();
  document.addParagraph(paragraphheadertext);

  document.createParagraph(" ");
  const txt = new TextRun("We wish to inform you that in line with the above Regulations, Banks, Microfinance Banks (MFBs) and the Deposit Protection Fund Board (DPFB) are required to share credit information of all their borrowers through licensed Credit Reference Bureaus (CRBs). ");
  const ptxt = new Paragraph();
  txt.size(20);
  ptxt.addRun(txt);
  document.addParagraph(ptxt);

  document.createParagraph(" ");
  const txt1 = new TextRun("A default in your card debt repayment will result in a negative impact on your credit record. If your card debt is classified as Non-Performing as per the Banking Act & Prudential Guidelines and/or as per the Microfinance Act, your credit profile at the CRBs will be adversely affected. ");
  const ptxt1 = new Paragraph();
  txt1.size(20);
  ptxt1.addRun(txt1);
  document.addParagraph(ptxt1);

  document.createParagraph(" ");
  const txt2 = new TextRun("Please note that your card account number " + letter_data.cardacct + ", card number " + letter_data.cardnumber + " is currently in default. It is outstanding at " + letter_data.OUT_BALANCE + " with arrears of " + letter_data.EXP_PMNT + ", having not paid the full installment(s) for 60 days. This card debt continues to accrue interest at a rate of 1.083% per month, on the daily outstanding balance and late payment fees at the rate of 5% on the arrears amount plus an excess fee of Kshs.1,000.00 monthly (if the total balance is above the limit).");
  const ptxt2 = new Paragraph();
  txt2.size(20);
  ptxt2.addRun(txt2);
  document.addParagraph(ptxt2);

  document.createParagraph(" ");
  const txt3 = new TextRun("We hereby notify you that we will proceed to adversely list you with the CRBs if your card debt becomes non-performing. To avoid an adverse listing, you are advised to clear the outstanding arrears within 30 days from the date of this letter. Payment can be made via Mpesa Paybill No. 400200 Account No. CR " + letter_data.cardacct + " ");
  const ptxt3 = new Paragraph();
  txt3.size(20);
  ptxt3.addRun(txt3);
  document.addParagraph(ptxt3);

  document.createParagraph(" ");
  const txt4 = new TextRun("You have a right of access to your credit report at the CRBs and you may dispute any erroneous information. You may request for your report by contacting the CRBs at the following addresses: ");
  const ptxt4 = new Paragraph();
  txt4.size(20);
  ptxt4.addRun(txt4);
  document.addParagraph(ptxt4);

  document.createParagraph(" ");
  //start crb
  const crb = new TextRun("TransUnion CRB                                                   Metropol CRB");
  const pcrb = new Paragraph();
  crb.size(20);
  crb.bold();
  pcrb.addRun(crb);
  document.addParagraph(pcrb);

  const crb1 = new TextRun("2nd Floor, Prosperity House,                                   1st Floor, Shelter Afrique Centre, Upper Hill, Nairobi. ");
  const pcrb1 = new Paragraph();
  crb1.size(20);
  pcrb1.addRun(crb1);
  document.addParagraph(pcrb1);

  const crb2 = new TextRun("Westlands Road, Off Museum Hill,                         P.O Box 35331 - 00200 ");
  const pcrb2 = new Paragraph();
  crb2.size(20);
  pcrb2.addRun(crb2);
  document.addParagraph(pcrb2);

  const crb3 = new TextRun("Westlands, Nairobi. P.O. Box 46406, 00100           NAIROBI, KENYA. ");
  const pcrb3 = new Paragraph();
  crb3.size(20);
  pcrb3.addRun(crb3);
  document.addParagraph(pcrb3);

  const crb4 = new TextRun("NAIROBI, KENYA Telephone: +254 (0) 20          Telephone: +254 (0) 20 2689881/27113575  ");
  const pcrb4 = new Paragraph();
  crb4.size(20);
  pcrb4.addRun(crb4);
  document.addParagraph(pcrb4);

  const crb5 = new TextRun("51799/3751360/2/4/5 Fax: +254 (0) 20 3751344    Fax: +254 (0) 20273572 ");
  const pcrb5 = new Paragraph();
  crb5.size(20);
  pcrb5.addRun(crb5);
  document.addParagraph(pcrb5);

  const crb6 = new TextRun("Email: info@transunion.co.ke                                 Email: creditbureau@metropol.co.ke ");
  const pcrb6 = new Paragraph();
  crb6.size(20);
  pcrb6.addRun(crb6);
  document.addParagraph(pcrb6);

  const crb9 = new TextRun("Website: www.crbafrica.com                                  www.metropolcorporation.com  ");
  const pcrb9 = new Paragraph();
  crb9.size(20);
  // crb9.underline();
  crb9.color("blue")
  pcrb9.addRun(crb9);
  document.addParagraph(pcrb9);

  const crb7 = new TextRun("Please text your name to 21272 and                                                  ");
  const pcrb7 = new Paragraph();
  crb7.size(20);
  pcrb7.addRun(crb7);
  document.addParagraph(pcrb7);

  const crb8 = new TextRun("follow instructions to secure a copy of your CRB report.                                                 ");
  const pcrb8 = new Paragraph();
  crb8.size(20);
  pcrb8.addRun(crb8);
  document.addParagraph(pcrb8);
  //

  document.createParagraph(" ");
  document.createParagraph("Yours sincerely, ");
  // sign
  document.createImage(fs.readFileSync("./sign_rose.png"), 100, 50);
  //sign
  const sign = new TextRun("ROSE KARAMBU ");
  const psign = new Paragraph();
  sign.size(20);
  psign.addRun(sign);
  document.addParagraph(psign);

  const signtext = new TextRun("COLLECTIONS SUPPORT MANAGER.");
  const paragraphsigntext = new Paragraph();
  signtext.bold();
  signtext.underline();
  signtext.size(22);
  paragraphsigntext.addRun(signtext);
  document.addParagraph(paragraphsigntext);

  const packer = new Packer();

  packer.toBuffer(document).then((buffer) => {
    fs.writeFileSync(LETTERS_DIR + letter_data.cardacct + DATE + "prelistingcc.docx", buffer);
    //conver to pdf
    // if pdf format
    if(letter_data.format == 'pdf'){
      convert();
      const convert = () => {
        word2pdf.word2pdf(LETTERS_DIR + letter_data.cardacct + DATE + "prelistingcc.docx")
          .then(data => {
            console.log('data here ...');
            fs.writeFileSync(LETTERS_DIR+ letter_data.cardacct + DATE + 'prelistingcc.pdf', data);
            res.json({result: 'success', message: LETTERS_DIR + letter_data.cardacct + DATE + "prelistingcc.pdf"})
          }, error  => {
            console.log('error ...', error)
            res.json({result: 'error', message: 'Exception occured'});
          })
      }
    } else {
      // res.sendFile(path.join(LETTERS_DIR + letter_data.cardacct + DATE + 'prelistingcc.docx'));
      res.json({result: 'success', message: LETTERS_DIR + letter_data.cardacct + DATE + "prelistingcc.docx"})
    }
  }).catch((err) => {
    console.log(err);
    res.json({result: 'error', message: 'Exception occured'});
  });
});

module.exports = router;