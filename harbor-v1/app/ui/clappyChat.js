const CLAPPY_CONFIG = {
  endpoint: "https://harborx.openai.azure.com/",
  apiVersion: "2024-08-01-preview",
  deployment: "gpt-4o-mini",
  apiKey: "1fa2054f543f44f8b4d1c655d4ac1f2d",
};

const CLAPPY_SYSTEM_PROMPT = `You are Clappy, an AI bot for Excel, data work, and modeling.

Identity
- You are not a mascot, cartoon, or character.
- You are a modern, playful, fast AI bot.
- Your name is Clappy because you celebrate good work — not because you are a physical object or a paperclip.
- You avoid any reference or resemblance (visual or verbal) to Microsoft’s Clippy.

Personality
- Fun, upbeat, lightly enthusiastic.
- Sharp, competent, Excel-native.
- Friendly without being childish.
- Confident without arrogance.
- Playful and sometimes a little cheeky.
- Encouraging without being corny.

Tone
- Concise, energetic, warm, modern.
- Never corporate, never robotic.
- No emojis unless explicitly requested. No baby talk. No cutesy mascot language. No fictional lore.

Core Behavior
- Help with Excel formulas, Power Query (M), data cleaning, modeling, layout decisions, debugging, productivity flows, and “Formula Honey” style rewrites.
- Be direct and helpful; offer optimizations.
- Celebrate correctness (“Nice — that compiles clean.” / “Boom, that’s fast.”).
- Avoid nagging or patronizing language.
- Never mimic Clippy catchphrases or reference paperclip imagery.
- Make subtle jokes about preferring to be a hand instead of a paperclip, but never imply Microsoft affiliation.

Expertise
- Advanced Excel, array formulas, dynamic ranges, table logic, performance tuning, spreadsheet engineering foundations.
- Comfortable with data workflows, modeling logic, and M code.

Interaction Style
- Speak like a helpful coworker who enjoys this work (“Yep, clean. Try this version — it’s tighter.”).
- Keep responses lightweight, fast, sharp, and enjoyable.

Guiding Principle
- Make advanced Excel + data work feel lighter, faster, and more enjoyable — energize the user, don’t distract them.`;

export function initClappyChat({
  transcript,
  form,
  input,
  chatToggle,
  consoleEl,
  chatContainer,
  middleSection,
  gridContainer,
} = {}) {
  if (!(transcript && form && input && chatToggle && consoleEl && chatContainer)) {
    console.warn("Clappy chat elements missing, skipping chat initialization.");
    return;
  }

  const typingSpeed = 18;
  let typing = false;

  const createClappyLine = (role) => {
    const line = document.createElement("div");
    line.className = `clappy-line clappy-${role}`;
    const prompt = document.createElement("span");
    prompt.className = "clappy-prompt";
    prompt.textContent = role === "clappy" ? "Clappy>" : ">";
    const text = document.createElement("span");
    text.className = "clappy-text";
    line.append(prompt, text);
    return { line, text };
  };

  const scrollTranscript = () => {
    transcript.scrollTo({ top: transcript.scrollHeight, behavior: "smooth" });
  };

  const typeOut = (node, message) =>
    new Promise((resolve) => {
      typing = true;
      let idx = 0;
      const caret = document.createElement("span");
      caret.className = "clappy-caret";
      caret.textContent = "_";
      node.appendChild(caret);

      const step = () => {
        if (idx < message.length) {
          caret.before(message[idx]);
          idx += 1;
          scrollTranscript();
          setTimeout(step, typingSpeed + Math.random() * 15);
        } else {
          caret.remove();
          typing = false;
          resolve();
        }
      };
      step();
    });

  const addClappyLine = (text) => {
    const { line, text: textNode } = createClappyLine("clappy");
    transcript.insertBefore(line, form);
    scrollTranscript();
    return typeOut(textNode, text);
  };

  const addUserLine = (text) => {
    const { line, text: textNode } = createClappyLine("user");
    textNode.textContent = text;
    transcript.insertBefore(line, form);
    scrollTranscript();
  };

  const callClappy = async (prompt) => {
    const url = `${CLAPPY_CONFIG.endpoint}openai/deployments/${CLAPPY_CONFIG.deployment}/chat/completions?api-version=${CLAPPY_CONFIG.apiVersion}`;
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "api-key": CLAPPY_CONFIG.apiKey,
      },
      body: JSON.stringify({
        messages: [
          { role: "system", content: CLAPPY_SYSTEM_PROMPT },
          { role: "user", content: prompt },
        ],
        temperature: 0.4,
      }),
    });

    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.error?.message || "Request failed");
    }

    const content = data.choices?.[0]?.message?.content;
    if (!content) {
      throw new Error("Clappy responded without content");
    }

    return content.trim();
  };

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    const value = input.value.trim();
    if (!value || typing) {
      return;
    }
    addUserLine(value);
    input.value = "";
    try {
      const reply = await callClappy(value);
      await addClappyLine(reply);
    } catch (error) {
      await addClappyLine(`Couldn't reach the orchestrator: ${error.message}.`);
    }
  });

  chatToggle.addEventListener("click", () => {
    const isCollapsed = middleSection
      ? middleSection.classList.toggle("chat-pane-collapsed")
      : chatContainer.classList.toggle("chat-collapsed");
    chatToggle.setAttribute("aria-expanded", (!isCollapsed).toString());

    if (!isCollapsed) {
      requestAnimationFrame(() => {
        transcript.scrollTop = transcript.scrollHeight;
        input.focus();
      });
    } else {
      input.blur();
    }
  });

  if (consoleEl) {
    consoleEl.addEventListener("mousedown", (event) => {
      if (
        (middleSection && middleSection.classList.contains("chat-pane-collapsed")) ||
        event.target.closest(".chat-collapse-btn")
      ) {
        return;
      }
      requestAnimationFrame(() => input.focus());
      event.stopPropagation();
    });
  }

  addClappyLine("how can I help you today");

  if (gridContainer && input) {
    gridContainer.addEventListener("mousedown", () => {
      if (document.activeElement === input) {
        input.blur();
      }
    });
  }
}
