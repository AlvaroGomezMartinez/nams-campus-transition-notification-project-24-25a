/**********************************************************************
 *                                                                    *
 *      nams-campus-transition-notification-project-24-25a            *
 *                                                                    *
 * This script is designed to be run from a Google Sheet. It will     *
 * send an email with attachments and links to a list of recipients   *
 * based on the value in the "Campus" column of the sheet.            *
 *                                                                    *
 * The administrator will first create two letters using              *
 * Autocrat. The letters will be saved in a campus specific           *
 * shared Google Drive folder. When the adminstrator is ready to      *
 * send the email, he clicks on the 'Notify Campuses' user menu       *
 * and then the option that is provided by the dropdown. An email     *
 * will be sent to the campus administrators with a link to the       *
 * folder and the two letters.                                        *
 *                                                                    *
 * Project Lead: John Decker, Associate Principal, NAMS               *
 * Script Author: Alvaro Gomez, Academic Technology Coach             *
 *                alvaro.gomez@nisd.net                               *
 *                Office: 1-210-397-9408                              *
 *                Mobile: 1-210-363-1577                              *
 *                                                                    *
 * Last Updated: 09/05/24                                             *
 **********************************************************************/

const EMAIL_SENT_COL = "Return Date";
const DATE_SENT_COL = "Date when the email was sent to campuses";

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📬 Notify Campuses")
    .addItem("Send emails to campuses with a date in column F and a blank in column BD", "sendEmails")
    .addToUi();
}

function sendEmails(
  sheet = SpreadsheetApp.openById(
    "11kbezUuY2o7P0cvgKMW5Kjp72SY2cOiZFa7SjjIfpeE",
  ).getSheetByName("Form Responses 2"),
) {
  const sheetName = sheet.getName();
  const targetSheetname = "Form Responses 2";

  if (sheetName !== targetSheetname) {
    SpreadsheetApp.getUi().alert(
      'This function can only be run from the "' + targetSheetname + '" sheet.',
    );
    return;
  }

  const dataRange = sheet.getDataRange();
  const data = dataRange.getDisplayValues();
  const heads = data.shift();

  const emailSentColIdx = heads.indexOf(DATE_SENT_COL);

  if (emailSentColIdx === -1) {
    SpreadsheetApp.getUi().alert("Required columns are missing.");
  }

  const obj = data.map((r) =>
    heads.reduce((o, k, i) => ((o[k] = r[i] || ""), o), {}),
  );

  // Initialize counters for success and error counts for dialog display
  let successCount = 0;
  let errorCount = 0;
  let errorMessages = []; // Stores the error messages for dialog display

  obj.forEach(function (row, rowIdx) {
    if (row[EMAIL_SENT_COL] !== "" && row[DATE_SENT_COL] === "") {
      try {
        const campusInfo = getInfoByCampus(row["Campus"]);
        const recipients = campusInfo.recipients;
        const driveLink = campusInfo.driveLink;
        const emailTemplate = getGmailTemplateFromDrafts_(row, driveLink);
        const msgObj = fillInTemplateFromObject_(
          emailTemplate.message,
          row,
          driveLink,
        );

        GmailApp.sendEmail(recipients, msgObj.subject, msgObj.text, {
          htmlBody: msgObj.html,
          replyTo: "john.decker@nisd.net",
          cc: "john.decker@nisd.net"
        });

        successCount++;
        sheet.getRange(rowIdx + 2, emailSentColIdx + 1).setValue(new Date());
      } catch (e) {
        Logger.log(`Error on row ${rowIdx + 2}: ${e.message}`);
        errorCount++;
        errorMessages.push(`Row ${rowIdx + 2}: ${e.message}`);
      }
    }
  });

  SpreadsheetApp.getUi().alert(
    `Emails Sent: ${successCount}\nErrors: ${errorCount}\n${errorMessages.join(
      "\n",
    )}`,
  );

  function getInfoByCampus(campusValue) {
    switch (campusValue.toLowerCase()) {
      case "bernal":
        return {
          recipients: [
            "david.laboy@nisd.net",
            "jose.mendez@nisd.net",
            "monica.flores@nisd.net",
            "sally.maher@nisd.net",
          ],
          driveLink: "1jyY1gPUgj2xLd7K5AWto6l3ictRP2xHQ",
        };
      case "briscoe":
        return {
          recipients: [
            "Joe.bishop@nisd.net",
            "nereida.ollendieck@nisd.net",
            "brigitte.rauschuber@nisd.net",
            "veronica.martinez@nisd.net",
          ],
          driveLink: "1gBcGRl700LGfhMPrHx32dZ2j3MwPlqYS",
        };
      case "connally":
        return {
          recipients: ["erica.robles@nisd.net", "monica.ramirez@nisd.net"],
          driveLink: "1dFXFBakVlHRaE2M4SQR9u-hVoFRs73-6",
        };
      case "folks":
        return {
          recipients: "yvette.lopez@nisd.net",
          driveLink: "1iA4d7ju4dU7aOcqr4R6PezzHUG5x1LK9",
        };
      case "garcia":
        return {
          recipients: [
            "mark.lopez@nisd.net",
            "lori.persyn@nisd.net",
            "julie.minnis@nisd.net",
            "mateo.macias@nisd.net",
            "anna.lopez@nisd.net",
          ],
          driveLink: "1nK6IFB5SQkLmJUMo1cFkXob_ngbOECaM",
        };
      case "hobby":
        return {
          recipients: ["gregory.dylla@nisd.net", "jaime.heye@nisd.net"],
          driveLink: "1OmqKg2tjRmGxx_FbdzPm0seMdjaqDr1F",
        };
      case "holmgreen":
        return {
          recipients: ["cheryl.parra@nisd.net", "frank.johnson@nisd.net"],
          driveLink: "1nGHHinv5vS9O9XuVyKg1SPIzpleVbhoh",
        };
      case "jefferson":
        return {
          recipients: [
            "Nicole.aguirreGomez@nisd.net",
            "monica.cabico@nisd.net",
            "maria-1.martinez@nisd.net",
            "tiffany.watkins@nisd.net",
            "catherine.villela@nisd.net",
          ],
          driveLink: "1kKehUlXwjOnwTjTD0BOuw3L0XmvpXRRF",
        };
      case "jones":
        return {
          recipients: [
            "rudolph.arzola@nisd.net",
            "javier.lazo@nisd.net",
            "erica.lashley@nisd.net",
            "nicole.mcevoy@nisd.net",
            "brian.pfeiffer@nisd.net",
          ],
          driveLink: "1BGe1rvOSZOmGkXU95l_ZiyODA2yTBFsZ",
        };
      case "jones magnet":
        return {
          recipients: "xavier.maldonado@nisd.net",
          driveLink: "1L1roRy3NfUXIm7MmC76YeLo3F_9MhqIE",
        };
      case "jordan":
        return {
          recipients: [
            "anabel.romero@nisd.net",
            "Shannon.Zavala@nisd.net",
            "laurel.graham@nisd.net",
            "jessica.marcha@nisd.net",
            "robert.ruiz@nisd.net",
            "ryanne.barecky@nisd.net",
            "adrian.hysten@nisd.net",
            "patti.vlieger@nisd.net",
          ],
          driveLink: "1lOKHdlFWLqDMg4IAOglq1V6om_nHpjnD",
        };
      case "luna":
        return {
          recipients: [
            "leti.chapa@nisd.net",
            "jennifer.cipollone@nisd.net",
            "amanda.king@nisd.net",
            "lisa.richard@nisd.net",
          ],
          driveLink: "1JBEzGoN44aZlfgw2ZOnMewGqTya0y-dN",
        };
      case "neff":
        return {
          recipients: [
            "yvonne.correa@nisd.net",
            "theresa.heim@nisd.net",
            "joseph.castellanos@nisd.net",
            "adriana.aguero@nisd.net",
            "priscilla.vela@nisd.net",
            "michele.adkins@nisd.net",
            "laura-i.sanroman@nisd.net",
            "jessica.montalvo@nisd.net",
          ],
          driveLink: "",
        };
      case "pease":
        return {
          recipients: [
            "Lynda.Desutter@nisd.net",
            "tamara.campbell-babin@nisd.net",
            "tiffany.flores@nisd.net",
            "tanya.alanis@nisd.net",
            "brian.pfeiffer@nisd.net",
          ],
          driveLink: "1MvxU28snFNspmlryjeV27CgNYWsqGn6S",
        };
      case "rawlinson":
        return {
          recipients: [
            "jesus.villela@nisd.net",
            "david.rojas@nisd.net",
            "nicole.buentello@nisd.net",
            "elizabeth.smith@nisd.net",
          ],
          driveLink: "1j71vRF2rN4p75T2SVN_2_DovpfKKPn3s",
        };
      case "rayburn":
        return {
          recipients: [
            "Robert.Alvarado@nisd.net",
            "maricela.garza@nisd.net",
            "aissa.zambrano@nisd.net",
            "brandon.masters@nisd.net",
            "micaela.welsh@nisd.net",
          ],
          driveLink: "10sEkjJf2XZ38zBgrnRrfmj5UwVrE_N5h",
        };
      case "ross":
        return {
          recipients: [
            "mahntie.reeves@nisd.net",
            "christina.lozano@nisd.net",
            "claudia.salazar@nisd.net",
            "jason.padron@nisd.net",
            "katherine.vela@nisd.net",
            "dolores.cardenas@nisd.net",
            "roxanne.romo@nisd.net",
            "rose.vincent@nisd.net",
          ],
          driveLink: "13C0WgGpwhm0N6WMuMplWeKdNbfl7kv9C",
        };
      case "rudder":
        return {
          recipients: [
            "kevin.vanlanham@nisd.net",
            "ximena.hueramedina@nisd.net",
          ],
          driveLink: "1jl0pKTbYn16496dOK-AFt1WlpkW9fy64",
        };
      case "stevenson":
        return {
          recipients: [
            "anthony.allen01@nisd.net",
            "hilary.pilaczynski@nisd.net",
            "johanna.davenport@nisd.net",
            "amanda.cardenas@nisd.net",
            "chaeleen.garcia@nisd.net",
          ],
          driveLink: "1h6V7It6_hVdEjOlul6oIzCW8Jvcw-Xii",
        };
      case "stinson":
        return {
          recipients: "louis.villarreal@nisd.net",
          driveLink: "1hvX1GY7pbgur_hNt58T7uZx_hCW5b496",
        };
      case "straus":
        return {
          recipients: ["araceli.farias@nisd.net", "brandy.bergeron@nisd.net"],
          driveLink: "136mdaz16r8lonY5En6Rdc983hhji32k4",
        };
      case "vale":
        return {
          recipients: [
            "jenna.bloom@nisd.net",
            "brenda.rayburg@nisd.net",
            "daniel.novosad@nisd.net",
            "mary.harrington@nisd.net",
          ],
          driveLink: "1bNQsGIx-zqSXYRXHO3DBAu5sgACA0WPC",
        };
      case "zachry":
        return {
          recipients: [
            "Richard.DeLaGarza@nisd.net",
            "juliana.molina@nisd.net",
            "randolph.neuenfeldt@nisd.net",
            "jaclyn.galvan@nisd.net",
            "veronica.poblano@nisd.net",
            "matthew.patty@nisd.net",
            "monica.perez@nisd.net",
            "chris.benavidez@nisd.net",
          ],
          driveLink: "14SAjOwdYe-LELYw7PWl8WlfyiV2h5Rgj",
        };
      case "test":
        return {
          recipients: ["alvaro.gomez@nisd.net"], //, "john.decker@nisd.net"],
          driveLink: "1I7mkmBa3-sO_eG6f2KlSKyUg-G71ZAYK",
        };
      default:
        return {
          recipients: "",
          driveLink: "",
        };
    }
  }

  function getGmailTemplateFromDrafts_(row, driveLink) {
    return {
      message: {
        subject: "AEP Placement Transition Plan",
        html: `${row["Student Name"]} has nearly completed their assigned placement at NAMS and should be returning to ${row["Campus"]} on or around ${row["Return Date"]}.<br><br>   
              On their last day of placement, they will be given withdrawal documents and the parents/guardians will have been called and told to contact ${row["Campus"]} to set up an appointment to re-enroll and meet with an administrator/counselor.<br><br>  
              Below are links and attachments to a Personalized Transition Plan (with notes from NAMS' assigned social worker), the student's AEP Transition Plan (with grades and notes from their teachers at NAMS), and a link to ${row["Campus"]}'s folder with all of the transition plans for this year.<br><br>
              Please let me know if you have any questions or concerns.<br><br>
              Thank you for all you do,<br>
              JD<br><br>
              <ul>
                <li><a href="${row["Merged Doc URL - Home Campus Transition Plan"]}">Home Campus Transition Plan</a></li>
                <li><a href="${row["Merged Doc URL - Student Transition "]}">Student Transition Plan</a></li>
                <li><a href="https://drive.google.com/drive/folders/${driveLink}">Drive Folder</a></li>
                <li><a href="https://drive.google.com/file/d/1qnyQ8cCxLVM9D6rg4wkyBp6KrXIELfNx/view?usp=sharing">Updates in Special Education</a></li>
              </ul>`,
      },
    };
  }

  function fillInTemplateFromObject_(template, row, driveLink) {
    let template_string = JSON.stringify(template);
    template_string = template_string.replace(/{{[^{}]+}}/g, (key) => {
      if (key === "${driveLink}") {
        return escapeData_(driveLink);
      }
      return escapeData_(row[key.replace(/[{}]+/g, "")] || "", driveLink);
    });
    return JSON.parse(template_string);
  }

  function escapeData_(str) {
    return str
      .replace(/[\\]/g, "\\\\")
      .replace(/[\"]/g, '\\"')
      .replace(/[\/]/g, "\\/")
      .replace(/[\b]/g, "\\b")
      .replace(/[\f]/g, "\\f")
      .replace(/[\n]/g, "\\n")
      .replace(/[\r]/g, "\\r")
      .replace(/[\t]/g, "\\t");
  }
}
