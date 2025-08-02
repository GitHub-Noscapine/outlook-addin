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

  // Show loading spinner
  history.innerHTML = '<div class="spinner"></div>';
  
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



