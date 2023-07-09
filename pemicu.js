function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var sheetName = sheet.getName();
  var invetto = "invetto.id";

  // Set value header (C2)
  sheet.getRange("B2").setValue('=HYPERLINK("https://' + invetto + '"; IMAGE("https://' + invetto + '/images/logo.png"; 4; 45; 95))');
  // Set value header (B3-D3)
  sheet.getRange("B3:D3").setValue("Name Of Guest");

  // Get value B3
  var valueB3 = sheet.getRange("B3").getValue();

  sheet.getRange("B1").setValue("Generate Message and Link - " + sheetName);

  if (range.getColumn() === 2 && range.getRow() >= 4) {
    var row = range.getRow();
    var valueA = sheet.getRange(row, 2).getValue();

    // Min and max length for valueA
    var maxLengthA = 25;
    var minLengthA = 3;

    if (valueA.length < minLengthA || valueA.length > maxLengthA) {
      sheet.getRange(row, 3).setValue("");
      sheet.getRange(row, 4).setValue("");

      Browser.msgBox("'" + valueB3 + "' harus memiliki panjang antara 3 hingga 25 karakter.", Browser.Buttons.OK);
      return;
    }

    var sheetNameModified = sheetName.replace(/\s+/g, "").replace("&", "-").toLowerCase();
    var URL = "https://" + sheetNameModified + "." + invetto + "?to=" + encodeURI(valueA);

    var urlPattern = new RegExp("\\." + invetto, "i");
    if (!urlPattern.test(URL)) {
      Browser.msgBox("URL tidak valid. Hanya URL yang mengandung '" + invetto + "' yang diperbolehkan.");
      return;
    }

    if (valueA !== "") {
      var name = "Kepada Yth.\nBapak/Ibu/Saudara/i\n" + valueA + "\n___________\n";
      var bride = sheetName;

      var generatedValue =
        name +
        "اَلسَّلاَمُ عَلَيْكُمْ وَرَحْمَةُ اللهِ وَبَرَكَاتُه\n" +
        "بسم الله الرحمن الرحيم\n\n" +
        "Tanpa mengurangi rasa hormat, perkenankan kami mengundang Bapak/Ibu/Saudara/i,\n" +
        "teman sekaligus sahabat, untuk menghadiri acara pernikahan anak kami.\n\n" +
        "Berikut link untuk info lengkap dari acara kami:\n" +
        URL +
        "\n\nMerupakan suatu kebahagiaan bagi kami apabila Bapak/Ibu/Saudara/i\n" +
        "berkenan untuk hadir dan memberikan doa restu.\n\n" +
        "Wassalamu’alaikum Wr. Wb.\nTerima Kasih..\n\n" +
        "Hormat kami,\n" +
        bride +
        "\n___________";

      var columnC = sheet.getRange(row, 3);
      columnC.setValue(generatedValue);
      columnC.protect();

      sheet.getRange(row, 4).setValue('=HYPERLINK("' + URL + '"; "Invitation Link")');

      var maxLength = 600;
      if (generatedValue.length > maxLength) {
        sheet.getRange(row, 3).setValue("");
        sheet.getRange(row, 4).setValue("");
        Browser.msgBox("Pesan terlalu panjang. Silakan periksa kembali input Anda.", Browser.inputBox("adbahdbahjs"));
      }
    }
  } else if (range.getColumn() !== 2 || range.getRow() >= 4) {
    // Restore previous value
    range.setValue(e.oldValue);

    // Prevent further edits
    range.protect();

    var adminInstagramLink = "@" + invetto; // Replace with the actual Instagram link
    var contactAdminMessage = "Hubungi Admin Invetto.id bila ingin Mengubah Pesan " + adminInstagramLink + " di Instagram Admin";

    Browser.msgBox(contactAdminMessage, Browser.Buttons.OK_CANCEL);
  }
}
