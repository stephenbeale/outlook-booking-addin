"use strict";

var RECIPIENT = "kingsresidences@kcl.ac.uk";
var SUBJECT = "King's residences booking request";

function getBookingDate() {
  var today = new Date();
  var future = new Date(today.getTime() + 28 * 24 * 60 * 60 * 1000);

  var dayNames = [
    "Sunday", "Monday", "Tuesday", "Wednesday",
    "Thursday", "Friday", "Saturday"
  ];
  var monthNames = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];

  var dayName = dayNames[future.getDay()];
  var dayOfMonth = future.getDate();
  var monthName = monthNames[future.getMonth()];

  return dayName + " " + dayOfMonth + " " + monthName;
}

function getEmailBody(bookingDate) {
  return "Hi there,<br><br>" +
    "I would like to request a room booking as per below:<br><br>" +
    bookingDate + "<br><br>" +
    "Kind regards,<br>" +
    "Steve Beale";
}

function getEmailBodyPlainText(bookingDate) {
  return "Hi there,\n\n" +
    "I would like to request a room booking as per below:\n\n" +
    bookingDate + "\n\n" +
    "Kind regards,\nSteve Beale";
}

function composeEmail() {
  var bookingDate = getBookingDate();
  var htmlBody = getEmailBody(bookingDate);

  if (Office.context && Office.context.mailbox) {
    Office.context.mailbox.displayNewMessageForm({
      toRecipients: [RECIPIENT],
      subject: SUBJECT,
      htmlBody: htmlBody
    });
  } else {
    // Fallback for testing outside Outlook: open mailto link
    var plainBody = getEmailBodyPlainText(bookingDate);
    var mailto = "mailto:" + encodeURIComponent(RECIPIENT) +
      "?subject=" + encodeURIComponent(SUBJECT) +
      "&body=" + encodeURIComponent(plainBody);
    window.open(mailto);
  }
}

function updateUI() {
  var bookingDate = getBookingDate();
  document.getElementById("bookingDate").textContent = bookingDate;

  var previewEl = document.getElementById("emailPreview");
  previewEl.innerHTML =
    "<strong>To:</strong> " + RECIPIENT + "<br>" +
    "<strong>Subject:</strong> " + SUBJECT + "<br><br>" +
    getEmailBody(bookingDate);

  var btn = document.getElementById("composeBtn");
  btn.disabled = false;
  btn.addEventListener("click", composeEmail);
}

Office.onReady(function (info) {
  updateUI();
});
