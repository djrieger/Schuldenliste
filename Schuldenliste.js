// Konfiguration:
var formUrl = "URL des Schuldenformulars";
var resultsUrl = "URL der Ergebnis-/Berechnungstabelle";
var accountDetailsSpreadsheetID = "Spreadsheet ID der Kontodaten-Tabelle";
var emailOptions = { name: "Name der Schuldenliste", replyTo: "Absenderadresse für Benachrichtigungen der Schuldenliste" };
var adminEmail = "E-Mail-Adresse für Benachrichtigungen über jeden neuen Eintrag und Abrechnungen";
// -------------


var responsesSheetName = "Formularantworten";
var calculationSheetName = "Berechnungen";
var mailSheetName = "Mailadressen";
var accountDetailsSheetName = "Kontodaten";

function onOpen(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [ {name: "Beträge neu ausrechnen", functionName: "recalculate" }, { name: "Abrechnen", functionName: "bill" } ];
  ss.addMenu("Schuldenliste", menuEntries);
}

function onFormSubmit(e) {
  var doc = SpreadsheetApp.getActive(); 
  var sheet = doc.getSheetByName(responsesSheetName);
  var calculationSheet = doc.getSheetByName(calculationSheetName);
  var mailSheet = doc.getSheetByName(mailSheetName);
  var lastrow = sheet.getLastRow();

  calculationSheet.insertRowBefore(lastrow);
  updateSumFormulas();
  
  var amountCell = sheet.getRange("E" + lastrow);
  if (typeof amountCell.getValue() == 'string') {
    amountCell.setValue(amountCell.getValue().replace(".", ","));
  }  
  
  calculateRow(lastrow, sheet, calculationSheet);
  email(sheet, calculationSheet, mailSheet, lastrow, false);
}

function updateSumFormulas() {
  var doc = SpreadsheetApp.getActive(); 
  var calculationSheet = doc.getSheetByName(calculationSheetName);
  var lastrow = calculationSheet.getLastRow();
  var firstCol = 8;
  var lastCol = calculationSheet.getLastColumn();
  var range = calculationSheet.getRange(lastrow, firstCol, 1, lastCol - firstCol + 1);
  var formulas = range.getFormulas()[0];
  for (i = 0; i < formulas.length; i++) {
    formulas[i] = formulas[i].replace(lastrow - 2, lastrow - 1);
  }
  range.setFormulas([ formulas ]);
}

function testSubmit() {
  var doc = SpreadsheetApp.getActive(); 
  var sheet = doc.getSheetByName(responsesSheetName);
  var calculationSheet = doc.getSheetByName(calculationSheetName);
  var mailSheet = doc.getSheetByName(mailSheetName);
  var lastrow = sheet.getLastRow();

  calculateRow(lastrow, sheet, calculationSheet);
  
  email(sheet, calculationSheet, mailSheet, lastrow, true);
}

function getMailAddresses() {
  var mailSheet = SpreadsheetApp.getActive().getSheetByName(mailSheetName);
  var mailTable = mailSheet.getRange(2, 1, mailSheet.getLastRow() - 1, 2).getValues();
  return mailTable.map(function (elem) { return { user: elem[0], email: elem[1] }; });
}

function email(sheet, calculationSheet, mailSheet, lastrow, adminonly) {  
  var data = sheet.getRange(lastrow, 2, 1, 4).getValues()[0];
  var creditor = data[0], date = data[1], title = data[2], amount = data[3];
  var guestsNum = calculationSheet.getRange(lastrow, 6).getValue();
  var perPerson = calculationSheet.getRange(lastrow, 7).getValue();
 
  var nameHeaders = calculationSheet.getRange(1, 8, 1, calculationSheet.getLastColumn()).getValues()[0];  
  var balances = calculationSheet.getRange(calculationSheet.getLastRow(), 8, calculationSheet.getLastRow(), calculationSheet.getLastColumn()).getValues()[0];
  var backup = "";
  for (var j = 0; j < nameHeaders.length; j++) {
    backup += Utilities.formatString("%s: %.2f €\n", nameHeaders[j], balances[j]);
  }
  
  // E-Mail an Admin
  var message = Utilities.formatString("Es gibt einen neuen Eintrag in der Schuldenliste.\n\n%s hat für %s %.2f € ausgegeben und %d Person(en) als Gäste angegeben.\n\nBackup:\n%s", creditor, title, amount, guestsNum, backup);
  MailApp.sendEmail(adminEmail, "@Admin: " + title + " in Schuldenliste eingetragen", message, emailOptions);
  
  var mailTable = mailSheet.getRange(2, 1, mailSheet.getLastRow() - 1, 2).getValues();
  var counts = sheet.getRange(lastrow, 6, 1, sheet.getLastColumn()).getValues()[0];  
  var userNewBalance = 0;
  
  for (var i = 0; i < mailTable.length; i++) {
    var user = mailTable[i][0];
    var email = mailTable[i][1];
    
    if (adminonly) {
      email = adminEmail;
    }
    
    userNewBalance = balances[nameHeaders.indexOf(user)];
    var newBalanceMessage = Utilities.formatString("Dein neuer Kontostand auf der Schuldenliste beträgt %.2f €.", userNewBalance);
    
    if (user.toLowerCase() == creditor.toLowerCase()) {
      message = Utilities.formatString("Hallo %s,\n\nDein Eintrag für %s wurde in die Schuldenliste eingetragen. %s Zur Kontrolle: %s", user, title, newBalanceMessage, resultsUrl);
      MailApp.sendEmail(email, title + " in Schuldenliste eingetragen", message, emailOptions)
      mailSheet.getRange(i + 2, 3).setValue(new Date());
    } else if (counts[i] > 0) {
      message = Utilities.formatString("Hallo %s,\n\n%s hat für %s %.2f € ausgegeben", user, creditor, title, amount);
      if (counts[i] > 1) {
        message += Utilities.formatString(" und angegeben, dass du für dich und %d weitere Person(en) zahlst", counts[i] - 1);
      }
      message += Utilities.formatString(". Dein Anteil beläuft sich auf %.2f €.\n\n", perPerson * counts[i]);
      message += Utilities.formatString("%s Unter %s kannst du alle Einträge in der Schuldenliste sehen. Hinweis: Die Abrechnung der Schulden erfolgt erst später. Diese E-Mail ist nur dazu da, dass du die eingetragenen Informationen überprüfen kannst. Du hast auch etwas, das du in die Schuldenliste eintragen möchtest? Dann fülle einfach schnell das Formular unter %s aus.", newBalanceMessage, resultsUrl, formUrl);
      MailApp.sendEmail(email, title + " in Schuldenliste eingetragen", message, emailOptions);
      mailSheet.getRange(i + 2, 3).setValue(new Date());
    }
  }
}

function recalculate() {
  var doc = SpreadsheetApp.getActive(); 
  var sheet = doc.getSheetByName(responsesSheetName);
  var calculationSheet = doc.getSheetByName(calculationSheetName);
  var lastrow = sheet.getLastRow();
  
  for (var row = 2; row <= lastrow; row++) {
    calculateRow(row, sheet, calculationSheet);
  }
}

// row sollte entweder die neu in Antworten eingefügte Zeile sein oder eine andere zwischen 2 und der untersten eingefügten Zeile, die kopiert werden soll.
function calculateRow(row, sheet, calculationSheet) {
  // Statische Werte in Berechnungstabelle kopieren:
  var staticValues = sheet.getRange(row, 1, 1, 5).getValues();
  calculationSheet.getRange(row, 1, 1, 5).setValues(staticValues);
  
  var amount = calculationSheet.getRange("E" + row).getValue();
  // Anzahl Gäste
  calculationSheet.getRange("F" + row).setFormula("SUM(Formularantworten!F" + row + ":" + row + ")");
  // Betrag pro Gast
  calculationSheet.getRange("G" + row).setFormula("E" + row + "/F" + row);
  var perGuest = calculationSheet.getRange("G" + row).getValue();
  var cook = sheet.getRange(row, 2).getValue();  
  
  var headersRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var paysForRow = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  var newValues = [[]];
  
  // von Spalte F in Antwortentabelle (hier fangen die Einträge für die Anzahl der jeweiligen Gäste an)
  for (var sourcecol = 6; sourcecol <= sheet.getLastColumn(); sourcecol++) {
    var destcol = sourcecol + 2;
    var currentGuest = headersRow[sourcecol - 1];
    var check = sheet.getRange(1, sourcecol).getValue();
    var paysFor = paysForRow[sourcecol - 1];
    if (paysFor == "")
      paysFor = 0;
    check = sheet.getRange(row, sourcecol).getValue();
    
    if (cook.toLowerCase() == currentGuest.toLowerCase()) {
      newValues[0][sourcecol - 6] = amount - paysFor * perGuest;
      //calculationSheet.getRange(row, destcol).setValue(amount - paysFor * perGuest);
    } else {
      newValues[0][sourcecol - 6] = - paysFor * perGuest;
      //calculationSheet.getRange(row, destcol).setValue(- paysFor * perGuest);
    }
  }
  
  calculationSheet.getRange(row, 8, 1, newValues[0].length).setValues(newValues);
  
  // Neuesten Eintrag gelb markieren:
  var oldRow = calculationSheet.getRange("A" + (row-1) + ":" + (row-1));
  var newRow = calculationSheet.getRange("A" + row + ":" + row);
  oldRow.setBackground("transparent");
  newRow.setBackground("yellow");
}

function computeTransactions() {
  var doc = SpreadsheetApp.getActive(); 
  var calculationSheet = doc.getSheetByName(calculationSheetName);
  var lastrow = calculationSheet.getLastRow();
  // Kontostände
  var balances = calculationSheet.getRange("H" + lastrow + ":" + lastrow).getValues();
  var names = calculationSheet.getRange("H1:1").getValues();
  var totals = new Array();
  for (var col = 1; col <= balances[0].length; col++) {
    var amount = balances[0][col - 1];
    if (amount !== 0) {
      totals.push({ key: names[0][col - 1], value: amount });
    }
  }
  // Salden vom höchsten zum niedrigsten sortieren
  totals.sort(function (a, b) { return a.value < b.value ? 1 : -1; });

  var transactions = new Array();
  // Vom Gläubiger mit dem höchsten Saldo zu dem mit geringstem:
  for (var creditor = 0; creditor < totals.length; creditor++) {  
    if (totals[creditor].value <= 0) 
      continue;
    // Vom Schuldner mit dem niedrigsten Saldo (höchsten Schulden) zu dem mit dem höchsten (niedrigsten Schulden):
    for (var debitor = totals.length - 1; debitor >= 0; debitor--) {
      if (totals[creditor].value <= 0 || totals[debitor].value >= 0) 
        continue;
      var amount = Math.min(totals[creditor].value, Math.abs(totals[debitor].value));
      // neue Transaktion: from muss to amount zahlen.
      transactions.push({from: totals[debitor].key, to: totals[creditor].key, "amount": amount});
      // Salden anpassen
      totals[creditor].value -= amount;
      totals[debitor].value += amount; 
    }
  }
  return transactions;
}

function bill() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Abrechnung gestartet. Einen Moment bitte...');
    
  var doc = SpreadsheetApp.getActive(); 
  var calculationSheet = doc.getSheetByName(calculationSheetName);
  // Berechnungen auf Abrechnungsblatt kopieren
  var formattedDate = Utilities.formatDate(new Date(), "CET", "dd.MM.yyyy HH:mm:ss");
  var billingSheet = calculationSheet.copyTo(doc);
  billingSheet.setName("Abrechnung " + formattedDate);
  doc.setActiveSheet(billingSheet);
  doc.moveActiveSheet(4);
  billingSheet.hideSheet();
  
  // Formeln in Abrechnungsblatt durch feste Werte ersetzen
  var sourceRange = calculationSheet.getDataRange();
  sourceRange.copyValuesToRange(billingSheet, 1, sourceRange.getLastColumn(), 1, sourceRange.getLastRow());
  
  // Transaktionen berechnen und in Abrechnungsblatt schreiben
  var transactions = computeTransactions();
  
  mailBills(calculationSheet, billingSheet, transactions);
  
  billingSheet.setFrozenRows(0);
  billingSheet.insertRowsBefore(1, transactions.length + 5);
  billingSheet.getRange(1, 1).setValue("Abrechnung");
  for (var row = 2; row < transactions.length + 2; row++) {
    var transaction = transactions[row - 2];
    billingSheet.getRange(row, 1).setValue(transaction.from);
    billingSheet.getRange(row, 2).setValue("->");
    billingSheet.getRange(row, 3).setValue(transaction.to);
    billingSheet.getRange(row, 4).setValue(transaction.amount);
  }
  
  calculationSheet.getRange("E2").copyFormatToRange(billingSheet, 4, 4, 2, transactions.length + 2);
  
  billingSheet.showSheet();
}

function getBankTransferDetails() {
  var accountsSheet = SpreadsheetApp.openById(accountDetailsSpreadsheetID).getSheetByName(accountDetailsSheetName);
  var data = accountsSheet.getRange(2, 1, accountsSheet.getLastRow() - 1, accountsSheet.getLastColumn()).getValues();
  return data.filter(function (elem) { return elem[0] !== "" && elem[1] !== "" && elem[2] !== ""; });
}

function prettyConcat(array, accessFun) {
  // Nimm Identitätsfunktion, falls keine Callbackfunktion übergeben wurde:
  accessFun = typeof accessFun !== 'undefined' ? accessFun : function (x) { return x; };
  return array.reduce(function (prev, cur, index) { return prev += (index == 0 ? "" : (index != array.length - 1 ? ", " : " und ")) + accessFun(cur); }, "");
}

function shortenUrl(url) {
  return UrlShortener.Url.insert({ longUrl: url }).id;
}

function mailBills(calculationSheet, billingSheet, transactions) {
  var billingTableShortUrl = shortenUrl(Utilities.formatString("%s#gid=%d", SpreadsheetApp.getActiveSpreadsheet().getUrl(), billingSheet.getSheetId()));
  var emails = getMailAddresses();
  var accountDetails = getBankTransferDetails();
  for (var i = 0; i < emails.length; i++) {
    var me = emails[i].user;
    var myemail = emails[i].email;
    var report = Utilities.formatString("Hallo %s,\n\nes ist Zahltag! Die Schuldenliste wurde abgerechnet. ", me);
    // Behalte nur Schulden, die auf zwei Nachkommastellen gerundet ungleich Null sind
    transactions = transactions.filter(function (elem) { return Utilities.formatString("%.2f", elem.amount) !== "0.00"; });
    var whatIshouldPay = transactions.filter(function (elem) { return elem.from === me; });
    var whatIamPaid = transactions.filter(function (elem) { return elem.to === me; }); 
    if (whatIshouldPay.length === 0 && whatIamPaid.length === 0)
      continue;
    
    var transfers = [];
    if (whatIshouldPay.length > 0) {
      report += "Bitte zahle innerhalb einer Woche ";
      report += prettyConcat(whatIshouldPay.map(function (debt) {
        for (var j = 0; j < accountDetails.length; j++) {
          // Hat der dem ich Geld schulde Kontodaten angegeben?
          if (debt.to === accountDetails[j][0]) {
            transfers.push(accountDetails[j].concat([debt.amount]));
          }
        }
        return Utilities.formatString("an %s %.2f €", debt.to, debt.amount)
      }));
      report += "\n\n"
      
      if (transfers.length > 0) {
        report += Utilities.formatString("%s möchte(n), dass du deine Schulden überweist statt in Bar zu zahlen. Hier die Überweisungsdaten:\n\n", prettyConcat(transfers, function (a) { return a[0]; }));// transfers.reduce(function (prev, cur, index) { return prev += cur[0] + (index != transfers.length - 1 ? ", " : " und "); }, ""));
        report += transfers.map(function (elem) {
            return Utilities.formatString("Kontoinhaber: %s\nIBAN: %s\nBIC: %s\nBetrag: %.2f €\nVerwendungszweck: Abrechnung Schuldenliste %s", elem[1], elem[2], elem[3], elem[4], Utilities.formatDate(new Date(), "CET", "dd.MM.yyyy"));
          }).join("\n\n");
        report += "\n\n";
      }
    }
    
    if (whatIamPaid.length > 0) {
      report += "Du erhältst ";
      report += prettyConcat(whatIamPaid.map(function (debt) {
        return Utilities.formatString("von %s %.2f €", debt.from, debt.amount);
      }));
      report += ".\n\n";
    }
    
    report += Utilities.formatString("Unter %s kannst du die gesamte Abrechnungstabelle inkl. der Zahlungsposten sehen.", billingTableShortUrl);
    MailApp.sendEmail(myemail, emailOptions.name + " abgerechnet", report);
    //MailApp.sendEmail(adminEmail, emailOptions.name + " abgerechnet", report);
  }
  MailApp.sendEmail(adminEmail, "@Admin: " + emailOptions.name + " abgerechnet", Utilities.formatString("Unter %s#gid=%d kannst du die gesamte Abrechnungstabelle inkl. der Zahlungsposten sehen.", SpreadsheetApp.getActiveSpreadsheet().getUrl(), billingSheet.getSheetId()));
}