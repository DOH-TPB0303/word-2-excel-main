const mammoth = require('mammoth');
const { JSDOM } = require('jsdom');
const fs = require('fs');

//#region MAIN
async function convertWordTablesToJson(inputDoc, outputJson) {
  try {
    //* Use mammoth to convert the input docx to html
    const result = await mammoth.convertToHtml({ path: inputDoc });

    //* Use JSDOM to extract tables from the html created above
    const tables = extractTablesFromHtml(result.value);
    // console.log('🚀 ~ file: word-2-json.js:14 ~ convertWordTablesToJson ~ tables:', tables)

    //* new translations object
    const translations = {};

    tables.forEach((table) => {
      table.forEach((row) => {
        if (row.length >= 2) {
          const key = row[0].trim();
          //console.log('🚀 ~ file: word-2-json.js:23 ~ table.forEach ~ key:', key)
          const value = row[1].trim();
          //console.log('🚀 ~ file: word-2-json.js:25 ~ table.forEach ~ value:', value)
          translations[key] = value;
        }
      });
    });

    //* Write output to translation.json file
    fs.writeFileSync(outputJson, JSON.stringify(translations, null, 2));
    console.log('Success!');

  } catch (error) {
    console.error('Error: ', error);
  }
}
//#endregion MAIN

//* Use JSDOM to find all <tables> in the html output
//* return those to convertWordTablesToExcel function
function extractTablesFromHtml(html) {
  const tables = [];
  const dom = new JSDOM(html);
  const document = dom.window.document;
  const tableElements = document.querySelectorAll('table');
  //   console.log('🚀 ~ file: process-word.js:29 ~ extractTablesFromHtml ~ tableElements:', tableElements);

  tableCount = 0;
  tableElements.forEach((tableElem, tableIndex) => {
    tableCount++;
    tableName = '';
    if (tableCount === 4) {
      tableName = 'header.';
    } else if (tableCount === 5) {
      tableName = 'vaccinepage.';
    } else if (tableCount === 6) {
      tableName = 'dashboardpage.';
    } else if (tableCount === 7) {
      tableName = 'vaccineform.';
    } else if (tableCount === 8) {
      tableName = 'receivedpage.';
    } else if (tableCount === 9) {
      tableName = 'formatsms.';
    } else if (tableCount === 10) {
      tableName = 'formatnotfoundsms.';
    } else if (tableCount === 11) {
      tableName = 'formathtml.';
    } else if (tableCount === 12) {
      tableName = 'formatnotfoundhtml.';
    } else if (tableCount === 13) {
      tableName = 'qrpage.';
    } else if (tableCount === 14) {
      tableName = 'faqpage.';
    } else if (tableCount === 15) {
      tableName = 'footer.';
    } else if (tableCount === 16) {
      tableName = 'month.';
    }
    const rows = [];
    const rowElements = tableElem.querySelectorAll('tr');
    // console.log('🚀 ~ file: process-word.js:35 ~ tableElements.forEach ~ rowElements:', rowElements);

    rowElements.forEach((rowElem, rowIndex) => {
      const cells = [];
      const cellElements = rowElem.querySelectorAll('td, th');
      //   console.log('🚀 ~ file: process-word.js:39 ~ rowElements.forEach ~ cellElements:', cellElements);

      cellElCount = 1;
      keyText = '';
      cellElements.forEach((cell) => {
        if (cellElCount === 1) {
          keyText = tableName + cell.textContent.trim();
          //if (tableCount > 3) {
          //  console.log('keyText (original) = [' + keyText + ']');
          //}
          if (keyText == 'header.selectlanguage') {
            keyText = 'header.selectlanguage';
          } else if (keyText == 'header.highcontrast') {
            keyText = 'header.highcontrast';
          } else if (keyText == 'header.contactus') {
            keyText = 'header.contactus';
          } else if (keyText == 'header.settings') {
            keyText = 'header.settings';
          } else if (keyText == 'header.dohlogotext**') {
            keyText = '';
          } else if (keyText == 'header.waverifylogo***') {
            keyText = '';
          } else if (keyText == 'header.waverifylogoalt**') {
            keyText = '';
          } else if (keyText == 'vaccinepage.title') {
            keyText = 'vaccinepage.title';
          } else if (keyText == 'vaccinepage.subtitle') {
            keyText = 'vaccinepage.subtitle';
          } else if (keyText == 'dashboardpage.tabtext') {
            keyText = 'dashboardpage.tabtext';
          } else if (keyText == 'dashboardpage.contentheader') {
            keyText = 'dashboardpage.contentheader';
          } else if (keyText == 'dashboardpage.content1') {
            keyText = 'dashboardpage.content1';
          } else if (keyText == 'dashboardpage.content2') {
            keyText = 'dashboardpage.content2';
          } else if (keyText == 'dashboardpage.content3') {
            keyText = 'dashboardpage.content3';
          } else if (keyText == 'dashboardpage.content4') {
            keyText = 'dashboardpage.content4';
          } else if (keyText == 'dashboardpage.Content5') {
            keyText = '';
          } else if (keyText == 'vaccineform.filloutheader') {
            keyText = 'vaccineform.filloutheader';
          } else if (keyText == 'vaccineform.title') {
            keyText = 'vaccineform.title';
          } else if (keyText == 'vaccineform.subtitle') {
            keyText = 'vaccineform.subtitle';
          } else if (keyText == 'vaccineform.lastname') {
            keyText = 'vaccineform.lastname';
          } else if (keyText == 'vaccineform.lastnameErrorMsg2') {
            keyText = 'vaccineform.lastnameErrorMsg';
          } else if (keyText == 'vaccineform.firstname') {
            keyText = 'vaccineform.firstname';
          } else if (keyText == 'vaccineform.firstnameErrorMsg1') {
            keyText = 'vaccineform.firstnameErrorMsg';
          } else if (keyText == 'vaccineform.dateofbirth') {
            keyText = 'vaccineform.dateofbirth';
          } else if (keyText == 'vaccineform.monthLabel') {
            keyText = 'vaccineform.monthLabel';
          } else if (keyText == 'vaccineform.dayLabel') {
            keyText = 'vaccineform.dayLabel';
          } else if (keyText == 'vaccineform.yearLabel') {
            keyText = 'vaccineform.yearLabel';
          } else if (keyText == 'vaccineform.dateofbirthValidationMsg1') {
            keyText = 'vaccineform.dateofbirthErrorMsg1';
          } else if (keyText == 'vaccineform.dateofbirthValidationMsg2') {
            keyText = 'vaccineform.dateofbirthErrorMsg2';
          } else if (keyText == 'vaccineform.dateofbirthValidationMsg3') {
            keyText = 'vaccineform.dateofbirthErrorMsg3';
          } else if (keyText == 'vaccineform.dateofbirthErrorMsg4') {
            keyText = 'vaccineform.dateofbirthErrorMsg4';
          } else if (keyText == 'vaccineform.phoneemailinfo') {
            keyText = 'vaccineform.phoneemailinfo';
          } else if (keyText == 'vaccineform.Phone') {
            keyText = 'vaccineform.Phone';
          } else if (keyText == 'vaccineform.phoneErrorMsg1') {
            keyText = 'vaccineform.phoneErrorMsg1';
          } else if (keyText == 'vaccineform.Email') {
            keyText = 'vaccineform.Email';
          } else if (keyText == 'vaccineform.emailErrorMsg1') {
            keyText = 'vaccineform.emailErrorMsg1';
          } else if (keyText == 'vaccineform.pincode') {
            keyText = 'vaccineform.pincode';
          } else if (keyText == 'vaccineform.pinErrorMsg1') {
            keyText = 'vaccineform.pinErrorMsg1';
          } else if (keyText == 'vaccineform.pinErrorMsg2') {
            keyText = 'vaccineform.pinErrorMsg2';
          } else if (keyText == 'vaccineform.pinErrorMsg3') {
            keyText = 'vaccineform.pinErrorMsg3';
          } else if (keyText == 'vaccineform.pinErrorMsg4') {
            keyText = 'vaccineform.pinErrorMsg4';
          } else if (keyText == 'vaccineform.pinErrorMsg5') {
            keyText = 'vaccineform.pinErrorMsg5';
          } else if (keyText == 'vaccineform.pinErrorMsg6') {
            keyText = 'vaccineform.pinErrorMsg6';
          } else if (keyText == 'vaccineform.pinErrorMsg7') {
            keyText = 'vaccineform.pinErrorMsg7';
          } else if (keyText == 'vaccineform.pinErrorMsg8') {
            keyText = 'vaccineform.pinErrorMsg8';
          } else if (keyText == 'vaccineform.note') {
            keyText = 'vaccineform.note';
          } else if (keyText == 'vaccineform.checkboxdescription') {
            keyText = 'vaccineform.checkboxdescription';
          } else if (keyText == 'vaccineform.agreementErrorMsg1') {
            keyText = 'vaccineform.agreementErrorMsg1';
          } else if (keyText == 'vaccineform.safe') {
            keyText = 'vaccineform.safe';
          } else if (keyText == 'vaccineform.vaccineproductname') {
            keyText = 'vaccineform.vaccineproductname';
          } else if (keyText == 'vaccineform.subsubtitle') {
            keyText = 'vaccineform.subsubtitle';
          } else if (keyText == 'vaccineform.fproductname') {
            keyText = 'vaccineform.fproductname';
          } else if (keyText == 'vaccineform.flotnumber') {
            keyText = 'vaccineform.flotnumber';
          } else if (keyText == 'vaccineform.ffirstvaccinationdate') {
            keyText = 'ffirstvaccinationdate';
          } else if (keyText == 'vaccineform.fhealthcareprofessional') {
            keyText = 'vaccineform.fhealthcareprofessional';
          } else if (keyText == 'vaccineform.subsubsubtitle') {
            keyText = 'vaccineform.subsubsubtitle';
          } else if (keyText == 'vaccineform.sproductname') {
            keyText = '';
          } else if (keyText == 'vaccineform.slotnumber') {
            keyText = '';
          } else if (keyText == 'vaccineform.sfirstvaccinationdate') {
            keyText = '';
          } else if (keyText == 'vaccineform.shealthcareprofessional') {
            keyText = '';
          } else if (keyText == 'vaccineform.submitbutton') {
            keyText = 'vaccineform.submitbutton';
          } else if (keyText == 'vaccineform.note24hour') {
            keyText = 'vaccineform.note24hour';
          } else if (keyText == 'vaccineform.ok') {
            keyText = 'vaccineform.ok';
          } else if (keyText == 'vaccineform.cancel') {
            keyText = 'vaccineform.cancel';
          } else if (keyText == 'receivedpage.title') {
            keyText = 'receivedpage.title';
          } else if (keyText == 'receivedpage.thankyou') {
            keyText = 'receivedpage.thankyou';
          } else if (keyText == 'receivedpage.content1') {
            keyText = 'receivedpage.content1';
          } else if (keyText == 'receivedpage.content2') {
            keyText = 'receivedpage.content2';
          } else if (keyText == 'receivedpage.content3') {
            keyText = 'receivedpage.content3';
          } else if (keyText == 'qrpage.pincode') {
            keyText = '';
          } else if (keyText == 'qrpage.enterPin') {
            keyText = 'vaccineform.enterPin';
          } else if (keyText == 'qrpage.title') {
            keyText = 'qrpage.title';
          } else if (keyText == 'qrpage.vacinfo') {
            keyText = 'qrpage.vacinfo';
          } else if (keyText == 'qrpage.name') {
            keyText = 'qrpage.name';
          } else if (keyText == 'qrpage.dateofbirth') {
            keyText = 'qrpage.dateofbirth';
          } else if (keyText == 'qrpage.dose') {
            keyText = 'qrpage.dose';
          } else if (keyText == 'qrpage.date') {
            keyText = 'qrpage.date';
          } else if (keyText == 'qrpage.type') {
            keyText = 'qrpage.type';
          } else if (keyText == 'qrpage.flotnumber') {
            keyText = 'qrpage.flotnumber';
          } else if (keyText == 'qrpage.howtosave') {
            keyText = 'qrpage.howtosave';
          } else if (keyText == 'qrpage.takeascreenshot') {
            keyText = 'qrpage.takeascreenshot';
          } else if (keyText == 'qrpage.or') {
            keyText = 'qrpage.or';
          } else if (keyText == 'qrpage.printrecord') {
            keyText = 'qrpage.printrecord';
          } else if (keyText == 'qrpage.download') {
            keyText = 'qrpage.download';
          } else if (keyText == 'qrpage.minrequirementsandroid') {
            keyText = 'qrpage.minrequirementsandroid';
          } else if (keyText == 'qrpage.minrequirementsiOS') {
            keyText = 'qrpage.minrequirementsios';
          } else if (keyText == 'qrpage.needhelp') {
            keyText = 'qrpage.needhelp';
          } else if (keyText == 'qrpage.incorrect') {
            keyText = 'qrpage.incorrect';
          } else if (keyText == 'faqpage.title') {
            keyText = 'faqpage.title';
          } else if (keyText == 'faqpage.01question') {
            keyText = 'faqpage.01question';
          } else if (keyText == 'faqpage.01answer') {
            keyText = 'faqpage.01answer';
          } else if (keyText == 'faqpage.02question') {
            keyText = 'faqpage.02question';
          } else if (keyText == 'faqpage.02answer') {
            keyText = 'faqpage.02answer';
          } else if (keyText == 'faqpage.03question') {
            keyText = 'faqpage.03question';
          } else if (keyText == 'faqpage.03answer') {
            keyText = 'faqpage.03answer';
          } else if (keyText == 'faqpage.04question') {
            keyText = 'faqpage.04question';
          } else if (keyText == 'faqpage.04answer') {
            keyText = 'faqpage.04answer';
          } else if (keyText == 'faqpage.05question') {
            keyText = 'faqpage.05question';
          } else if (keyText == 'faqpage.05answer') {
            keyText = 'faqpage.05answer';
          } else if (keyText == 'faqpage.06question') {
            keyText = 'faqpage.06question';
          } else if (keyText == 'faqpage.06answer') {
            keyText = 'faqpage.06answer';
          } else if (keyText == 'faqpage.07question') {
            keyText = 'faqpage.07question';
          } else if (keyText == 'faqpage.07answer') {
            keyText = 'faqpage.07answer';
          } else if (keyText == 'faqpage.08question') {
            keyText = 'faqpage.08question';
          } else if (keyText == 'faqpage.08answer') {
            keyText = 'faqpage.08answer';
          } else if (keyText == 'faqpage.09question') {
            keyText = 'faqpage.09question';
          } else if (keyText == 'faqpage.09answer') {
            keyText = 'faqpage.09answer';
          } else if (keyText == 'faqpage.10question') {
            keyText = 'faqpage.10question';
          } else if (keyText == 'faqpage.10answer') {
            keyText = 'faqpage.10answer';
          } else if (keyText == 'faqpage.11question') {
            keyText = 'faqpage.11question';
          } else if (keyText == 'faqpage.11answer') {
            keyText = 'faqpage.11answer';
          } else if (keyText == 'faqpage.12question') {
            keyText = 'faqpage.12question';
          } else if (keyText == 'faqpage.12answer') {
            keyText = 'faqpage.12answer';
          } else if (keyText == 'faqpage.13question') {
            keyText = 'faqpage.13question';
          } else if (keyText == 'faqpage.13answer') {
            keyText = 'faqpage.13answer';
          } else if (keyText == 'faqpage.14question') {
            keyText = 'faqpage.14question';
          } else if (keyText == 'faqpage.14answer') {
            keyText = 'faqpage.14answer';
          } else if (keyText == 'faqpage.15question') {
            keyText = 'faqpage.15question';
          } else if (keyText == 'faqpage.15answer') {
            keyText = 'faqpage.15answer';
          } else if (keyText == 'faqpage.16question') {
            keyText = 'faqpage.16question';
          } else if (keyText == 'faqpage.16answer') {
            keyText = 'faqpage.16answer';
          } else if (keyText == 'faqpage.needhelptitle') {
            keyText = 'faqpage.needhelptitle';
          } else if (keyText == 'faqpage.needhelpcontent01') {
            keyText = 'faqpage.needhelpcontent01';
          } else if (keyText == 'faqpage.needhelpcontent02') {
            keyText = 'faqpage.needhelpcontent02';
          } else if (keyText == 'faqpage.needhelpcontent03') {
            keyText = 'faqpage.needhelpcontent03';
          } else if (keyText == 'faqpage.needhelpcontent04') {
            keyText = 'faqpage.needhelpcontent04';
          } else if (keyText == 'faqpage.needhelpcontent05') {
            keyText = 'faqpage.needhelpcontent05';
          } else if (keyText == 'faqpage.needhelpcontent06') {
            keyText = 'faqpage.needhelpcontent06';
          } else if (keyText == 'footer.home***') {
            keyText = 'footer.home';
          } else if (keyText == 'footer.conditionsofuse') {
            keyText = 'footer.conditionsofuse';
          } else if (keyText == 'footer.privacypolicy') {
            keyText = 'footer.privacypolicy';
          } else if (keyText == 'footer.accessibility') {
            keyText = 'footer.accessibility';
          } else if (keyText == 'footer.faq') {
            keyText = 'footer.faq';
          } else if (keyText == 'footer.contactus') {
            keyText = 'footer.contactus';
          } else if (keyText == 'footer.copyright') {
            keyText = 'footer.copyright';
          } else if (keyText == 'month.january') {
            keyText = 'month.january';
          } else if (keyText == 'month.february') {
            keyText = 'month.february';
          } else if (keyText == 'month.march') {
            keyText = 'month.march';
          } else if (keyText == 'month.april') {
            keyText = 'month.april';
          } else if (keyText == 'month.may') {
            keyText = 'month.may';
          } else if (keyText == 'month.june') {
            keyText = 'month.june';
          } else if (keyText == 'month.july') {
            keyText = 'month.july';
          } else if (keyText == 'month.august') {
            keyText = 'month.august';
          } else if (keyText == 'month.september') {
            keyText = 'month.september';
          } else if (keyText == 'month.october') {
            keyText = 'month.october';
          } else if (keyText == 'month.november') {
            keyText = 'month.november';
          } else if (keyText == 'month.december') {
            keyText = 'month.december';
          } else if (keyText == 'formatsms.text') {
            keyText = '';
          } else if (keyText == 'formatnotfoundsms.text') {
            keyText = '';
          } else if (keyText == 'formathtml.heading') {
            keyText = '';
          } else if (keyText == 'formathtml.infoText') {
            keyText = '';
          } else if (keyText == 'formathtml.viewlink') {
            keyText = '';
          } else if (keyText == 'formathtml.learnMore') {
            keyText = '';
          } else if (keyText == 'formathtml.questions') {
            keyText = '';
          } else if (keyText == 'formathtml.visitFAQ') {
            keyText = '';
          } else if (keyText == 'formathtml.stayInformed') {
            keyText = '';
          } else if (keyText == 'formathtml.viewInfo') {
            keyText = '';
          } else if (keyText == 'formathtml.emailLabel') {
            keyText = '';
          } else if (keyText == 'formatnotfoundhtml.heading') {
            keyText = '';
          } else if (keyText == 'formatnotfoundhtml.infoText') {
            keyText = '';
          } else if (keyText == 'formatnotfoundhtml.nextSteps') {
            keyText = '';
          } else if (keyText == 'formatnotfoundhtml.questions') {
            keyText = '';
          } else if (keyText == 'formatnotfoundhtml.visitFAQ') {
            keyText = '';
          } else if (keyText == 'formatnotfoundhtml.stayInformed') {
            keyText = '';
          } else if (keyText == 'formatnotfoundhtml.viewInfo') {
            keyText = '';
          } else if (keyText == 'formatnotfoundhtml.emailLabel') {
            keyText = '';
          } else {
            keyText = '';
          }
          //if (tableCount > 3) {
          //  console.log('keyText (after transformation) = [' + keyText + ']');
          //}
          cells.push(keyText);
          cellElCount++;
        } else {
          cells.push(cell.textContent.trim());
        }
        
      });
      //if (tableCount > 3) {
      //  console.log('keyText (just before push) = [' + keyText + '] length=' + keyText.trim().length);
      //}
      if (keyText.trim().length > 0) {
        rows.push(cells);
      }      
    });
    if (tableCount > 3) {
      tables.push(rows);
    }    
  });

  return tables;
}

//* Location of your input docx
//* Location of your output json
const inputPath1 = 'example_arm.docx';
const outputPath1 = 'translation_arm.json';
convertWordTablesToJson(inputPath1, outputPath1);

const inputPath2 = 'example_bas.docx';
const outputPath2 = 'translation_bas.json';
convertWordTablesToJson(inputPath2, outputPath2);

const inputPath3 = 'example_cat.docx';
const outputPath3 = 'translation_cat.json';
convertWordTablesToJson(inputPath3, outputPath3);

const inputPath4 = 'example_hbrw.docx';
const outputPath4 = 'translation_hbrw.json';
convertWordTablesToJson(inputPath4, outputPath4);

const inputPath5 = 'example_hc.docx';
const outputPath5 = 'translation_hc.json';
convertWordTablesToJson(inputPath5, outputPath5);

const inputPath6 = 'example_it.docx';
const outputPath6 = 'translation_it.json';
convertWordTablesToJson(inputPath6, outputPath6);

const inputPath7 = 'example_mala.docx';
const outputPath7 = 'translation_mala.json';
convertWordTablesToJson(inputPath7, outputPath7);

//! Run this from command line using `node word-2-json.js`
