
Office.onReady(function () {
  document.getElementById("insertTable").onclick = insertQuestionsTable;
  document.getElementById("resetQuestions").onclick = resetQuestions;
  resetQuestions();
});

const DEFAULT_QUESTIONS = [
  "Czy potwierdzasz, że zmiana została omówiona i zaakceptowana przez Członka Zarządu?",
  "Czy potwierdzasz, że przed zgłoszeniem zmiany skonsultowano z osobami lub działami, na które może mieć wpływ i potwierdzono zrozumienie tych zmian?",
  "Czy potwierdzasz, że rozumiesz i wiesz, jak zmiana wpłynie na dane, raporty, dokumenty, procesy biznesowe, rozliczenia i sposób działania systemu?",
  "Czy potwierdzasz, że bierzesz odpowiedzialność biznesową za zasadność zgłoszonej zmiany oraz za skutki wynikające z jej wdrożenia?"
];

function setStatus(msg, err) {
  const s=document.getElementById("status");
  s.textContent=msg;
  s.className=err?"error":"ok";
}

function getQuestions(){
  return document.getElementById("questions").value.split(/\r?\n/).map(x=>x.trim()).filter(Boolean);
}

function buildQuestionsTableHtml(questions){
  const rows = questions.map((q,i)=>`
<tr>
<td style="border:1px solid #cccccc;padding:7px 9px;">${i+1}. ${q}</td>
<td style="border:1px solid #cccccc;padding:7px 9px;width:150px;text-align:center;">&nbsp;</td>
</tr>`).join("");

  return `
<p style="font-family:Calibri,Arial,sans-serif;">Prosimy o uzupełnienie odpowiedzi w poniższej tabeli:</p>
<table cellpadding="0" cellspacing="0" border="0" width="760" style="width:760px;border-collapse:collapse;font-family:Calibri,Arial,sans-serif;">
<tr>
<td style="border:1px solid #DF292F;background:#DF292F;color:#fff;padding:7px 9px;font-weight:bold;">Pytanie</td>
<td style="border:1px solid #DF292F;background:#DF292F;color:#fff;padding:7px 9px;font-weight:bold;width:150px;text-align:center;">Odpowiedź TAK/NIE</td>
</tr>
${rows}
</table><br>`;
}

function insertQuestionsTable(){
  Office.context.mailbox.item.body.setSelectedDataAsync(
    buildQuestionsTableHtml(getQuestions()),
    { coercionType: Office.CoercionType.Html },
    function(result){
      if(result.status===Office.AsyncResultStatus.Succeeded){
        setStatus("Formularz został wstawiony.", false);
      } else {
        setStatus("Błąd wstawiania formularza.", true);
      }
    }
  );
}

function resetQuestions(){
  document.getElementById("questions").value = DEFAULT_QUESTIONS.join("\n");
  setStatus("Przywrócono pytania.", false);
}
