Office.onReady(function () {
  const insertButton = document.getElementById("insertTable");
  const resetButton = document.getElementById("resetQuestions");

  if (insertButton) insertButton.onclick = insertQuestionsTable;
  if (resetButton) resetButton.onclick = resetQuestions;

  resetQuestions();
});

const DEFAULT_QUESTIONS = [
  "Czy potwierdzasz, że zmiana została omówiona i zaakceptowana przez Członka Zarządu?",
  "Czy potwierdzasz, że przed zgłoszeniem zmiany skonsultowano z osobami lub działami, na które może mieć wpływ i potwierdzono zrozumienie tych zmian?",
  "Czy potwierdzasz, że rozumiesz i wiesz, jak zmiana wpłynie na dane, raporty, dokumenty, procesy biznesowe, rozliczenia i sposób działania systemu?",
  "Czy potwierdzasz, że bierzesz odpowiedzialność biznesową za zasadność zgłoszonej zmiany oraz za skutki wynikające z jej wdrożenia?"
];

function setStatus(message) {
  const s = document.getElementById("status");
  if (s) s.textContent = message;
}

function getQuestions() {
  const textarea = document.getElementById("questions");
  return textarea.value.split(/\r?\n/).map(x => x.trim()).filter(Boolean);
}

function buildQuestionsTableHtml(questions) {
  const rows = questions.map((q, i) => `
<tr>
<td style="border:1px solid #cccccc;padding:7px 9px;">${i+1}. ${q}</td>
<td style="border:1px solid #cccccc;padding:7px 9px;width:150px;text-align:center;">&nbsp;</td>
</tr>`).join("");

  return `
<table cellpadding="0" cellspacing="0" border="0" width="760" style="width:760px;border-collapse:collapse;">
<tr>
<td style="border:1px solid #DF292F;background:#DF292F;color:#fff;padding:7px 9px;font-weight:bold;">Pytanie</td>
<td style="border:1px solid #DF292F;background:#DF292F;color:#fff;padding:7px 9px;font-weight:bold;width:150px;text-align:center;">Odpowiedź TAK/NIE</td>
</tr>
${rows}
</table><br>`;
}

function insertQuestionsTable() {
  const html = buildQuestionsTableHtml(getQuestions());
  Office.context.mailbox.item.body.setSelectedDataAsync(
    html,
    { coercionType: Office.CoercionType.Html }
  );
}

function resetQuestions() {
  const textarea = document.getElementById("questions");
  if (textarea) textarea.value = DEFAULT_QUESTIONS.join("\n");
}
