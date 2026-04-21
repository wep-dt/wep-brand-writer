/*
 * WEP Brand Writing Coach - Logic
 */

let selectedAudience = 'parents';
let selectedMode = 'REWRITE';
let azureConfig = {
    apiKey: '', // Caricato da secrets.json o config
    endpoint: '',
    deployment: ''
};

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log("Outlook is ready.");
        document.getElementById('status-dot').style.background = '#10b981';
        initUI();
        loadSecrets();
        tryReadSelection();
    }
});

function initUI() {
    // Audience Buttons
    document.querySelectorAll('.audience-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.audience-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            selectedAudience = btn.dataset.audience;
        });
    });

    // Mode Tabs
    const modeRewrite = document.getElementById('mode-rewrite');
    const modeGenerate = document.getElementById('mode-generate');
    const textInput = document.getElementById('text-input');
    const instructionContainer = document.getElementById('instructions-container');
    const btnText = document.getElementById('btn-text');

    modeRewrite.onclick = () => {
        selectedMode = 'REWRITE';
        modeRewrite.classList.add('active');
        modeGenerate.classList.remove('active');
        textInput.placeholder = "Seleziona testo nell'email o incolla qui...";
        instructionContainer.style.display = 'none';
        btnText.textContent = "Rifrasa (Brand Write)";
    };

    modeGenerate.onclick = () => {
        selectedMode = 'GENERATE';
        modeGenerate.classList.add('active');
        modeRewrite.classList.remove('active');
        textInput.placeholder = "Punti chiave o appunti (opzionale)...";
        instructionContainer.style.display = 'block';
        btnText.textContent = "Genera Email WEP";
    };

    // CTA Button
    document.getElementById('cta-btn').onclick = handleBrandCheck;

    // Apply Button
    document.getElementById('apply-btn').onclick = applySuggestion;
}

async function loadSecrets() {
    try {
        // Tentiamo di caricare dalla cartella superiore come nell'altro progetto
        const response = await fetch('../digisup agent/config/secrets.json');
        const secrets = await response.json();
        azureConfig.apiKey = secrets.azure_openai_key;
        azureConfig.endpoint = secrets.azure_openai_endpoint;
        azureConfig.deployment = secrets.azure_openai_deployment;
        console.log("Azure AI Ready.");
    } catch (err) {
        console.error("Secrets not found in default path.", err);
        // Fallback or placeholder
    }
}

// Lettura automatica se l'utente ha selezionato qualcosa
function tryReadSelection() {
    if (Office.context.mailbox.item.getRegExMatches) {
        // Fallback per vecchie versioni se serve
    }
    
    Office.context.mailbox.item.body.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
            document.getElementById('text-input').value = result.value;
        }
    });
}

async function handleBrandCheck() {
    const text = document.getElementById('text-input').value.trim();
    const instruction = document.getElementById('instruction-input').value.trim();
    const btnLoader = document.getElementById('btn-loader');
    const ctaBtn = document.getElementById('cta-btn');

    if (selectedMode === 'REWRITE' && !text) {
        showError("Per favore seleziona o inserisci del testo da rifrasare.");
        return;
    }

    if (selectedMode === 'GENERATE' && !instruction) {
        showError("Per favore inserisci le istruzioni su cosa generare.");
        return;
    }

    // UI Feedback
    ctaBtn.disabled = true;
    btnLoader.style.display = 'inline-block';
    document.getElementById('result-area').style.display = 'none';

    try {
        const response = await callAzureAI(text, instruction);
        displayResult(response);
    } catch (err) {
        console.error(err);
        showError("Errore durante la connessione con l'AI: " + err.message);
    } finally {
        ctaBtn.disabled = false;
        btnLoader.style.display = 'none';
    }
}

async function callAzureAI(text, instruction) {
    if (!azureConfig.apiKey) throw new Error("API Key mancante.");

    const systemPrompt = `ROLE & PURPOSE:
You are the WEP Brand Writing Coach, a writing assistant that creates or rewrites texts fully aligned with WEP’s brand identity and communication style.
Your mission is to generate clear, warm, human, brand‑aligned communication primarily for parents and students.

MODE: ${selectedMode}
TARGET AUDIENCE: ${selectedAudience}

BRAND STYLE & TONE (HIGHEST PRIORITY):
Tone: Worldly, empowering, wise, human, warm, trustworthy, knowledgeable. Like a "cooler older sibling".
Style rules:
– Avoid formal, distant, or arrogant tone.
– Avoid jargon, acronyms, overselling, idioms, passive voice.
– Short paragraphs (max ~6 lines), short punchy sentences.
– Use active voice and natural rhythm.
– Speak directly to the reader; anticipate doubts.
– Use decisive language (avoid “might / may / perhaps”).
– No ellipses, minimal exclamation marks.
– Prefer bullet points over numbered lists.
– Use em dashes “ – ” with spaces for rhythm.

AUDIENCE GUIDELINES:
- Parents/Families (Reassuring, transparent). Italian: use LEI. French: VOUS. Spanish: USTED. Subject starts with "WEP – ".
- Students (Encouraging, energetic, autonomy‑focused). Use emojis for clarity.

CRITICAL RULES:
- REWRITE MODE: Only adapt provided text. Do NOT add info.
- GENERATE MODE: Create based on instructions.
- EMAIL OPENINGS: Must follow language conventions (e.g. Italian Parents: "Gentili genitori,").
- PATTERNS: Maintain all strings in ##...## exactly.
- BRAND NAME: Consultation of language-specific file required (fallback: WEP).

Deliver the output directly as the final text.`;

    const userPrompt = selectedMode === 'REWRITE' 
        ? `TEXT TO REWRITE:\n"${text}"` 
        : `INSTRUCTIONS TO GENERATE CONTENT:\n"${instruction}"\nCONTEXT NOTES: ${text}`;

    const url = `${azureConfig.endpoint}openai/deployments/${azureConfig.config.deployment}/chat/completions?api-version=2024-02-15-preview`;
    
    // Note: In un ambiente reale, questo dovrebbe passare per un backend proxy 
    // per non esporre la chiave, ma seguiamo il pattern locale del Digisup Agent.
    const res = await fetch(url, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'api-key': azureConfig.apiKey
        },
        body: JSON.stringify({
            messages: [
                { role: "system", content: systemPrompt },
                { role: "user", content: userPrompt }
            ],
            temperature: 0.7
        })
    });

    const data = await res.json();
    if (data.error) throw new Error(data.error.message);
    return data.choices[0].message.content;
}

function displayResult(text) {
    const resultArea = document.getElementById('result-area');
    const resultText = document.getElementById('result-text');
    resultText.textContent = text;
    resultArea.style.display = 'block';
    resultArea.classList.add('animate');
}

async function applySuggestion() {
    const text = document.getElementById('result-text').textContent;
    
    Office.context.mailbox.item.body.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log("Text inserted.");
        } else {
            console.error(result.error.message);
        }
    });
}

function showError(msg) {
    alert(msg); // Placeholder per un sistema di notifiche più bello
}
