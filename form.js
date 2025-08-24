let conversationId = null;
let lastAIReply = "";

console.log("Form.js loaded");

Office.onReady(() => {
  console.log("✅ Office.js is ready");
  const item = Office.context.mailbox.item;
  if (item && item.subject) {
    document.getElementById("emailSubject").innerText = item.subject;
    console.log("📨 Email subject loaded:", item.subject);
  } else {
    document.getElementById("emailSubject").innerText = "(Subject not available)";
    console.warn("⚠️ No subject available.");
  }

  // --- Telegram → Form bridge poller ---
  let bridgePolling = null;
  let lastBridgeTs = 0;
  
  // How to handle Telegram text when the field already has content:
  // "append"  → add below existing text (default)
  // "replace" → overwrite the field
  // "confirm" → ask user to Replace or Append
  const TELEGRAM_CONFLICT_MODE = "append";
  
  async function pollBridgeOnce() {
    try {
      const r = await fetch("http://127.0.0.1:5678/webhook/form-bridge/pull?t=" + Date.now(), {
        method: "GET",
        cache: "no-store"
      });
      if (!r.ok) return;
      const data = await r.json(); // { prompt, tone, chatId, ts } or { prompt: null }
      if (!data || !data.prompt) return;
      if (data.ts && data.ts <= lastBridgeTs) return; // already handled
  
      const input = document.getElementById("instruction"); // <-- make sure ID matches your textarea
      const toneEl = document.getElementById("tone");
      const askBtn = document.getElementById("askBtn");     // <-- make sure ID matches your button
  
      if (!input || !askBtn) {
        console.warn("Form elements not found: #instruction or #askBtn");
        return;
      }
  
      const incoming = (data.prompt || "").trim();
      const current  = (input.value || "").trim();
  
      if (toneEl && data.tone) toneEl.value = data.tone;
  
      if (!current) {
        console.log("[bridge] field empty → set + run");
        input.value = incoming;
      } else {
        if (TELEGRAM_CONFLICT_MODE === "replace") {
          console.log("[bridge] field had text → REPLACE");
          input.value = incoming;
        } else if (TELEGRAM_CONFLICT_MODE === "append") {
          console.log("[bridge] field had text → APPEND");
          input.value = current + "\n\n" + incoming;
        } else {
          console.log("[bridge] field had text → CONFIRM");
          const choice = window.confirm(
            "Telegram sent a new prompt.\n\nReplace existing text?\nOK = Replace, Cancel = Append"
          );
          input.value = choice ? incoming : current + "\n\n" + incoming;
        }
      }
  
      // Auto-run your existing flow
      // askBtn.click(); // enable or disable automatic button click
      lastBridgeTs = data.ts || Date.now();
    } catch (e) {
      console.warn("Bridge poll failed:", e);
    }
  }
  
  function startBridgePolling() {
    if (!bridgePolling) bridgePolling = setInterval(pollBridgeOnce, 2000);
  }
  
  // call this at the end of your Office.onReady block:
  startBridgePolling();

});

async function fetchEmailBody() {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    if (item.body && item.body.getAsync) {
      item.body.getAsync(Office.CoercionType.Text, result => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject("Could not fetch email body.");
        }
      });
    } else {
      reject("Body not available.");
    }
  });
}

// async function codify(text) {
  // const response = await fetch("http://localhost:5000/codify", {
    // method: "POST",
    // headers: { "Content-Type": "application/json" },
    // body: JSON.stringify({ body: text })
  // });
  // const result = await response.json();
  // return result.body;
// }

async function sendToN8N(payload) {
  const response = await fetch("http://localhost:5678/webhook/auto-reply", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  if (payload.type === "finalize") return {}; // Ignore response
  return await response.json();
}

function escapeHTML(str) {
  return str.replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;");
}

document.getElementById("askBtn").onclick = async () => {
  console.log("🟦 Ask AI button clicked");
  const prompt = document.getElementById("instruction").value;
  const tone = document.getElementById("tone").value;
  const subject = document.getElementById("emailSubject").innerText;
  const history = document.getElementById("history");

  // Disable buttons
  document.getElementById("draftBtn").disabled = true;
  document.getElementById("cancelBtn").disabled = true;
  document.getElementById("askBtn").disabled = true;

  // Show spinner overlay without clearing history
  const spinnerOverlay = document.createElement("div");
  spinnerOverlay.className = "spinner-overlay";
  spinnerOverlay.innerHTML = `<div class="spinner"></div>`;
  document.getElementById("history-wrapper").appendChild(spinnerOverlay);

  try {
    // Get conversation ID or generate a new one
    if (!conversationId) {
      const timestamp = new Date().toISOString().replace(/[-:.TZ]/g, "");
      conversationId = `conv-${timestamp}-${Math.floor(Math.random() * 10000)}`;
    }

    const type = history.innerText.includes("🧾 Prompt") ? "continue" : "start";

    // Fetch email body only once at start
    let emailBody = "";
    if (type === "start") {
      emailBody = await fetchEmailBody();
    }

    // 🔹 Codify emailBody + prompt via Python
    const codifyResponse = await fetch("http://localhost:5000/codify", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        conversationId,
        type,
        body: emailBody,
        subject: subject,
        customInstructions: prompt
      })
    });

    const codifyResult = await codifyResponse.json();
    const encodedBody = codifyResult.body || "";
    const encodedSubject = codifyResult.subject || "";
    const encodedPrompt = codifyResult.fullPrompt || "";

    // 🔹 Send codified body + prompt to n8n
    const payload = {
      type,
      conversationId,
      subject: encodedSubject,
      body: encodedBody,
      from: "user@outlook.com",
      promptStyle: tone,
      customInstructions: encodedPrompt
    };

    const n8nResponse = await sendToN8N(payload);
    const aiReply = n8nResponse.replyContent || "(No response)";

    // 🔹 Decode AI reply
    const decodeResponse = await fetch("http://localhost:5000/decodify", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        conversationId,
        body: aiReply
      })
    });

    const decodedResult = await decodeResponse.json();
    const decodedReply = decodedResult.body || "(Decoding failed)";

    lastAIReply = decodedReply;

    // 🔹 Append history (with raw prompt, not encoded)
    // history.innerText += `\n\n🧾 Prompt [${tone}]: ${prompt}\n✉️ Response: ${decodedReply}`;
    
    history.insertAdjacentHTML("beforeend",
  `  <div style="color: #0078D4;"><strong>🧾 Prompt [${tone}]:</strong> ${escapeHTML(prompt)}</div>
     <div>✉️ Response: ${escapeHTML(decodedReply)}</div><br>`
    );

    history.scrollTop = history.scrollHeight;

    document.getElementById("instruction").value = '';
    
  //} catch (error) {
    //console.error(error);
    //history.innerText += `\n\n❌ Error: ${error}`;
  //}
//};

  } catch (error) {
      console.error(error);
      history.innerHTML = `<div class="spinner"></div>\n\n❌ Error: ${error}`;
    } finally {
      // Re-enable buttons
      document.getElementById("draftBtn").disabled = false;
      document.getElementById("cancelBtn").disabled = false;
      document.getElementById("askBtn").disabled = false;
      const existingOverlay = document.querySelector(".spinner-overlay");
      if (existingOverlay) existingOverlay.remove();
    }
};

document.getElementById("draftBtn").onclick = async () => {
  console.log("🟨 Create Draft button clicked");
  if (!lastAIReply) {
  document.getElementById("history").innerText += "\n\n⚠️ No AI reply available to draft.";
  return;
  }

  if (conversationId) {
    await sendToN8N({ type: "finalize", conversationId });
  }

  Office.context.mailbox.item.displayReplyAllForm({
    htmlBody: `<p>${lastAIReply.trim().replace(/^\=+/g, "").replace(/\n/g, "<br>")}</p>`
  });
};

document.getElementById("cancelBtn").onclick = async () => {
  console.log("🟥 Cancel button clicked");

  if (conversationId) {
    await sendToN8N({ type: "finalize", conversationId });
  }

  document.getElementById("instruction").value = '';
  document.getElementById("history").innerText = `❌ Conversation cancelled and memory cleared.`;
  conversationId = null;
  lastAIReply = "";
};







