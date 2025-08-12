let emailText = "";

Office.onReady(() => {
  const saved = localStorage.getItem("openai_key");
  if (saved) document.getElementById("apiKey").value = saved;

  document.getElementById("saveKey").onclick = () => {
    localStorage.setItem("openai_key", document.getElementById("apiKey").value.trim());
    log("Saved API key locally (browser storage).");
  };

  document.getElementById("btnLoad").onclick = loadEmail;
  document.getElementById("btnSummary").onclick = () => runPrompt("Summarize the email briefly with bullets and a one-line key takeaway.");
  document.getElementById("btnDraft").onclick   = () => runPrompt("Draft a polite, concise reply to this email.");
});

function log(msg){ document.getElementById("out").textContent = msg; }

function loadEmail(){
  const item = Office.context.mailbox.item;
  if (!item) return log("No email item.");
  item.body.getAsync("text", (res)=>{
    if (res.status === Office.AsyncResultStatus.Succeeded){
      // basic strip of long quoted blocks
      emailText = (res.value || "").replace(/(^>.*\n?)+/gm,"").trim();
      log("Loaded email ("+emailText.length+" chars). Now click Summarize or Draft.");
    } else {
      log("Failed to read body: " + res.error.message);
    }
  });
}

async function runPrompt(task){
  const key = localStorage.getItem("openai_key");
  if (!key) return log("Missing API key. Save it first.");
  if (!emailText) return log("Load the current email first.");

  log("Thinking…");
  try{
    const resp = await fetch("https://api.openai.com/v1/chat/completions",{
      method:"POST",
      headers:{
        "Authorization":"Bearer "+key,
        "Content-Type":"application/json"
      },
      body: JSON.stringify({
        model: "gpt-4o-mini",
        messages: [
          {role:"system", content:"You are a helpful assistant for email triage. Be concise."},
          {role:"user", content: `${task}\n\nEmail:\n${emailText}`}
        ]
      })
    });
    if (!resp.ok){
      const t = await resp.text();
      throw new Error(`HTTP ${resp.status}: ${t}`);
    }
    const data = await resp.json();
    const out = data.choices?.[0]?.message?.content ?? "(no content)";
    log(out);
  }catch(e){
    log("Error: " + e.message);
  }
}
