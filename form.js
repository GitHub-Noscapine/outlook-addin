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

async function codify(text) {
  const response = await fetch("http://localhost:5000/codify", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ body: text })
  });
  const result = await response.json();
  return result.body;
}

async function sendToN8N(payload) {
  const response = await fetch("http://localhost:5678/webhook/auto-reply", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  if (payload.type === "finalize") return {}; // Ignore response
  return await response.json();
}

document.getElementById("askBtn").onclick = async () => {
  console.log("🟦 Ask AI button clicked");
  const prompt = document.getElementById("instruction").value;
  const tone = document.getElementById("tone").value;
  const subject = document.getElementById("emailSubject").innerText;
  const history = document.getElementById("history");

  try {
    const emailBody = await fetchEmailBody();
    const codifiedBody = await codify(emailBody);

    // Generate conversation ID on first request
    if (!conversationId) {
      const timestamp = new Date().toISOString().replace(/[-:.TZ]/g, "");
      conversationId = `conv-${timestamp}-${Math.floor(Math.random() * 10000)}`;
    }

    const payload = {
      type: conversationId ? (history.innerText.includes("🧾 Prompt") ? "continue" : "start") : "start",
      conversationId,
      subject: subject,
      body: codifiedBody,
      from: "user@outlook.com",
      promptStyle: tone,
      customInstructions: prompt
    };

    const response = await sendToN8N(payload);

    const aiReply = response.replyContent || "(No response)";
    lastAIReply = aiReply;

    history.innerText += `\n\n🧾 Prompt [${tone}]: ${prompt}\n✉️ Response: ${aiReply}`;
    document.getElementById("instruction").value = '';

  } catch (error) {
    history.innerText += `\n\n❌ Error: ${error}`;
  }
};

document.getElementById("draftBtn").onclick = async () => {
  console.log("🟨 Create Draft button clicked");
  if (!lastAIReply) {
    alert("⚠️ No AI reply to use as draft.");
    return;
  }

  if (conversationId) {
    await sendToN8N({ type: "finalize", conversationId });
  }

  Office.context.mailbox.displayReplyForm({
    htmlBody: `<p>${lastAIReply}</p>`
  });
};

document.getElementById("cancelBtn").onclick = () => {
  console.log("🟥 Cancel button clicked");
  document.getElementById("instruction").value = '';
  document.getElementById("history").innerText += `\n\n❌ Cancelled by user.`;
  conversationId = null;
  lastAIReply = "";
};
