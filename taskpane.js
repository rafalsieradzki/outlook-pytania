/* Tabela pytań Familijna v1.0 */
Office.onReady(function () {
  const insertButton = document.getElementById("insertTable");
  const resetButton = document.getElementById("resetQuestions");

  if (insertButton) insertButton.onclick = insertQuestionsTable;
  if (resetButton) resetButton.onclick = resetQuestions;

  setStatus("Dodatek gotowy.", false);
});

const DEFAULT_QUESTIONS = [
  "Czy akceptują Państwo termin realizacji?",
  "Czy dane do faktury są poprawne?",
  "Czy zamówienie może zostać uruchomione?",
  "Czy potwierdzają Państwo zakres zamówienia?"
];

function setStatus(message, isError) {
  const status = document.getElementById("status");
  if (!status) return;
  status.textContent = message || "";
  status.className = isError ? "error" : "ok";
}

function htmlEncode(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function getQuestions() {
  const textarea = document.getElementById("questions");
  const raw = textarea ? textarea.value : "";
  return raw
    .split(/\r?\n/)
    .map(x => x.trim())
    .filter(Boolean);
}

function buildQuestionsTableHtml(questions) {
  const rows = questions.map((q, index) => `
    <tr>
      <td style="border:1px solid #cccccc;padding:7px 9px;font-size:10.5pt;line-height:14px;color:#222;">
        ${index + 1}. ${htmlEncode(q)}
      </td>
      <td style="border:1px solid #cccccc;padding:7px 9px;font-size:10.5pt;line-height:14px;color:#222;width:150px;text-align:center;">
        TAK / NIE
      </td>
    </tr>`).join("");

  return `
    <div style="font-family:Calibri,Arial,sans-serif;color:#222;margin:0;padding:0;">
      <p style="margin:0 0 8px 0;font-size:11pt;line-height:15px;">
        Prosimy o uzupełnienie odpowiedzi w poniższej tabeli:
      </p>
      <table cellpadding="0" cellspacing="0" border="0" width="650"
             style="width:650px;border-collapse:collapse;font-family:Calibri,Arial,sans-serif;">
        <tr>
          <td style="border:1px solid #DF292F;background:#DF292F;color:#ffffff;padding:7px 9px;font-size:10.5pt;font-weight:bold;">
            Pytanie
          </td>
          <td style="border:1px solid #DF292F;background:#DF292F;color:#ffffff;padding:7px 9px;font-size:10.5pt;font-weight:bold;width:150px;text-align:center;">
            Odpowiedź TAK/NIE
          </td>
        </tr>
        ${rows}
      </table>
      <p style="margin:8px 0 0 0;font-size:9pt;color:#666;line-height:13px;">
        W kolumnie „Odpowiedź TAK/NIE” prosimy wpisać TAK albo NIE.
      </p>
    </div><br>`;
}

function insertQuestionsTable() {
  const questions = getQuestions();
  if (!questions.length) {
    setStatus("Brak pytań do wstawienia.", true);
    return;
  }

  const html = buildQuestionsTableHtml(questions);

  Office.context.mailbox.item.body.setSelectedDataAsync(
    html,
    { coercionType: Office.CoercionType.Html },
    function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        setStatus("Tabela pytań została wstawiona.", false);
      } else {
        const msg = result.error && result.error.message ? result.error.message : "Nieznany błąd Outlook API.";
        setStatus("Nie udało się wstawić tabeli: " + msg, true);
      }
    }
  );
}

function resetQuestions() {
  const textarea = document.getElementById("questions");
  if (textarea) textarea.value = DEFAULT_QUESTIONS.join("\n");
  setStatus("Przywrócono przykładowe pytania.", false);
}
