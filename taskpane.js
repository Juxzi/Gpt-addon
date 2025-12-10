// taskpane.js

const STORAGE_KEY_API = "chatgpt_word_addin_api_key";
const DEFAULT_MODEL = "gpt-4.1-mini";

let apiKeyInput,
  saveApiKeyButton,
  apiKeyStatus,
  useSelectionButton,
  clearPromptButton,
  promptTextarea,
  modeSelect,
  sendButton,
  insertAsCommentButton,
  resultTextarea,
  replaceSelectionButton,
  insertAtCursorButton,
  statusArea,
  modelInfoChip;

Office.onReady(() => {
  cacheDom();
  bindEvents();
  loadApiKeyFromStorage();
  setStatus("Prêt.");
  if (modelInfoChip) modelInfoChip.textContent = DEFAULT_MODEL;
});

function cacheDom() {
  apiKeyInput = document.getElementById("apiKeyInput");
  saveApiKeyButton = document.getElementById("saveApiKeyButton");
  apiKeyStatus = document.getElementById("apiKeyStatus");

  useSelectionButton = document.getElementById("useSelectionButton");
  clearPromptButton = document.getElementById("clearPromptButton");
  promptTextarea = document.getElementById("promptTextarea");

  modeSelect = document.getElementById("modeSelect");

  sendButton = document.getElementById("sendButton");
  insertAsCommentButton = document.getElementById("insertAsCommentButton");

  resultTextarea = document.getElementById("resultTextarea");
  replaceSelectionButton = document.getElementById("replaceSelectionButton");
  insertAtCursorButton = document.getElementById("insertAtCursorButton");

  statusArea = document.getElementById("statusArea");
  modelInfoChip = document.getElementById("modelInfoChip");
}

function bindEvents() {
  if (saveApiKeyButton) {
    saveApiKeyButton.addEventListener("click", onSaveApiKey);
  }
  if (useSelectionButton) {
    useSelectionButton.addEventListener("click", onUseSelectionToPanel);
  }
  if (clearPromptButton) {
    clearPromptButton.addEventListener("click", () => {
      promptTextarea.value = "";
      setStatus("Texte effacé.");
    });
  }
  if (sendButton) {
    sendButton.addEventListener("click", onSendToChatGPT_SelectionFirst);
  }
  if (insertAsCommentButton) {
    insertAsCommentButton.addEventListener("click", onInsertAsComment);
  }
  if (replaceSelectionButton) {
    replaceSelectionButton.addEventListener("click", onReplaceSelectionFromResult);
  }
  if (insertAtCursorButton) {
    insertAtCursorButton.addEventListener("click", onInsertAtCursorFromResult);
  }
}

/* ------------ Gestion clé API ------------ */

function loadApiKeyFromStorage() {
  try {
    const stored = localStorage.getItem(STORAGE_KEY_API);
    if (stored && apiKeyInput) {
      apiKeyInput.value = "••••••••••••••••";
      apiKeyInput.dataset.hasStoredKey = "true";
      setApiKeyStatus("Clé API déjà enregistrée.", "ok");
    }
  } catch (e) {
    setApiKeyStatus("Impossible d’accéder au stockage local.", "warning");
  }
}

function onSaveApiKey() {
  if (!apiKeyInput) return;

  let value = apiKeyInput.value.trim();

  if (!value && apiKeyInput.dataset.hasStoredKey === "true") {
    setApiKeyStatus("Clé déjà enregistrée.", "ok");
    return;
  }

  if (!value || !value.startsWith("sk-")) {
    setApiKeyStatus("Clé API invalide. Elle doit commencer par “sk-”.", "error");
    return;
  }

  try {
    localStorage.setItem(STORAGE_KEY_API, value);
    apiKeyInput.value = "••••••••••••••••";
    apiKeyInput.dataset.hasStoredKey = "true";
    setApiKeyStatus("Clé API sauvegardée.", "ok");
  } catch (e) {
    setApiKeyStatus("Erreur lors de la sauvegarde de la clé API.", "error");
  }
}

function getEffectiveApiKey() {
  if (!apiKeyInput) return null;

  const raw = apiKeyInput.value.trim();
  if (raw && raw.startsWith("sk-")) {
    return raw;
  }

  try {
    const stored = localStorage.getItem(STORAGE_KEY_API);
    return stored || null;
  } catch (e) {
    return null;
  }
}

function setApiKeyStatus(message, type) {
  if (!apiKeyStatus) return;
  apiKeyStatus.textContent = message;
  apiKeyStatus.style.color =
    type === "ok" ? "#047857" : type === "error" ? "#b91c1c" : "#6b7280";
}

/* ------------ Status global ------------ */

function setStatus(message) {
  if (!statusArea) return;
  statusArea.textContent = message;
}

/* ------------ Word : sélection / insertion ------------ */

/**
 * Bouton "Utiliser la sélection Word" -> copie la sélection dans le textarea
 * (utile si tu veux voir/modifier le texte dans le panneau).
 */
async function onUseSelectionToPanel() {
  setStatus("Récupération de la sélection...");
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      const text = selection.text || "";
      if (!text) {
        setStatus("Aucun texte sélectionné dans le document.");
        return;
      }
      promptTextarea.value = text;
      setStatus("Sélection copiée dans le panneau.");
    });
  } catch (error) {
    console.error(error);
    setStatus("Erreur lors de la récupération de la sélection.");
  }
}

/**
 * Insertion du résultat (zone résultat -> remplace sélection)
 * Option secondaire : si tu veux manipuler à la main le résultat.
 */
async function onReplaceSelectionFromResult() {
  const text = (resultTextarea && resultTextarea.value) || "";
  if (!text) {
    setStatus("Aucun résultat à insérer.");
    return;
  }

  setStatus("Remplacement de la sélection...");
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertText(text, Word.InsertLocation.replace);
      await context.sync();
    });
    setStatus("Sélection remplacée.");
  } catch (error) {
    console.error(error);
    setStatus("Erreur lors du remplacement de la sélection.");
  }
}

/**
 * Insertion du résultat à la position du curseur
 */
async function onInsertAtCursorFromResult() {
  const text = (resultTextarea && resultTextarea.value) || "";
  if (!text) {
    setStatus("Aucun résultat à insérer.");
    return;
  }

  setStatus("Insertion dans le document...");
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertText(text, Word.InsertLocation.end);
      await context.sync();
    });
    setStatus("Texte inséré dans le document.");
  } catch (error) {
    console.error(error);
    setStatus("Erreur lors de l’insertion dans le document.");
  }
}

/**
 * Insertion du résultat en commentaire
 */
async function onInsertAsComment() {
  const text = (resultTextarea && resultTextarea.value) || "";
  if (!text) {
    setStatus("Aucun résultat à insérer en commentaire.");
    return;
  }

  setStatus("Insertion en commentaire...");
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertComment(text);
      await context.sync();
    });
    setStatus("Commentaire ajouté.");
  } catch (error) {
    console.error(error);
    setStatus("Erreur lors de l’ajout du commentaire.");
  }
}

/* ------------ Construction du prompt ------------ */

function buildPrompt(mode, userText) {
  userText = userText.trim();
  switch (mode) {
    case "rewrite":
      return (
        "Réécris le texte suivant en français dans un style clair, professionnel et fluide, " +
        "sans changer le sens, adapté à des documents Word d’entreprise :\n\n" +
        userText
      );
    case "correct":
      return (
        "Corrige uniquement l’orthographe, la grammaire et la ponctuation du texte suivant en français, " +
        "sans changer la formulation ni le ton, sauf si nécessaire pour la clarté :\n\n" +
        userText
      );
    case "summarize":
      return (
        "Résume en français le texte suivant en quelques phrases claires et structurées, " +
        "adaptées à un document Word professionnel :\n\n" +
        userText
      );
    case "translate_en":
      return (
        "Traduits le texte suivant en anglais professionnel, fluide et naturel. " +
        "Préserve le sens et le ton d’origine :\n\n" +
        userText
      );
    case "custom":
    default:
      // Ici tu peux écrire directement un prompt “complexe” dans Word,
      // il sera envoyé tel quel.
      return userText;
  }
}

/* ------------ Appel OpenAI + remplacement auto ------------ */

/**
 * Comportement principal :
 *  - récupère d'abord la sélection dans Word
 *  - si elle est vide, tente d'utiliser le textarea du panneau
 *  - envoie à ChatGPT
 *  - remplace AUTOMATIQUEMENT la sélection par la réponse
 */
async function onSendToChatGPT_SelectionFirst() {
  const apiKey = getEffectiveApiKey();
  if (!apiKey) {
    setStatus("Aucune clé API valide. Renseigne et sauvegarde ta clé.");
    setApiKeyStatus("Aucune clé API trouvée.", "error");
    return;
  }

  let baseText = "";

  setStatus("Lecture de la sélection dans Word...");
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();
      baseText = selection.text || "";
    });
  } catch (error) {
    console.error(error);
    setStatus("Erreur lors de la lecture de la sélection.");
    return;
  }

  // Si pas de sélection : fallback sur le textarea
  if (!baseText.trim() && promptTextarea) {
    baseText = promptTextarea.value || "";
  }

  if (!baseText.trim()) {
    setStatus(
      "Aucun texte à envoyer. Écris ton texte / prompt dans Word, sélectionne-le, puis clique sur Envoyer."
    );
    return;
  }

  const mode = modeSelect ? modeSelect.value : "rewrite";
  const prompt = buildPrompt(mode, baseText);

  const aiResponse = await callOpenAIAndGetText(apiKey, prompt);
  if (!aiResponse) {
    // Erreur déjà gérée dans callOpenAIAndGetText
    return;
  }

  // Affiche aussi la réponse dans la zone résultat (optionnel)
  if (resultTextarea) {
    resultTextarea.value = aiResponse;
  }

  // Remplace automatiquement la sélection dans Word par la réponse
  setStatus("Remplacement du texte sélectionné par la réponse de l’IA...");
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertText(aiResponse, Word.InsertLocation.replace);
      await context.sync();
    });
    setStatus("Texte remplacé par le résultat de l’IA.");
  } catch (error) {
    console.error(error);
    setStatus(
      "Réponse reçue, mais erreur lors du remplacement dans le document. Tu peux copier-coller depuis la zone Résultat."
    );
  }
}

/**
 * Appel à l’API OpenAI qui retourne simplement le texte
 */
async function callOpenAIAndGetText(apiKey, prompt) {
  if (!sendButton) return null;

  setStatus("Appel à l’API OpenAI en cours...");
  sendButton.disabled = true;
  const originalLabel = sendButton.textContent;
  sendButton.textContent = "En cours...";

  try {
    const body = {
      model: DEFAULT_MODEL,
      messages: [
        {
          role: "system",
          content:
            "Tu es un assistant IA qui aide à rédiger, corriger et améliorer des textes pour des documents Word professionnels en français.",
        },
        {
          role: "user",
          content: prompt,
        },
      ],
    };

    const response = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      const errorText = await response.text().catch(() => "");
      console.error("OpenAI API error:", response.status, errorText);
      setStatus(
        "Erreur de l’API OpenAI (" +
          response.status +
          "). Vérifie ta clé ou ton quota."
      );
      return null;
    }

    const data = await response.json();
    const message = data.choices?.[0]?.message?.content || "";
    if (!message) {
      setStatus("Réponse vide de l’IA.");
      return null;
    }

    setStatus("Réponse reçue.");
    return message;
  } catch (error) {
    console.error("Erreur lors de l’appel OpenAI:", error);
    setStatus("Erreur réseau ou problème lors de l’appel à OpenAI.");
    return null;
  } finally {
    sendButton.disabled = false;
    sendButton.textContent = originalLabel;
  }
}
