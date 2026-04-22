/*
 * WEP Brand Writing Coach - Web & Outlook JS
 */

let selectedAudience = 'parents';
let selectedMode = 'REWRITE';
let azureConfig = { 
    apiKey: '', 
    endpoint: 'https://digisup-openai.openai.azure.com', 
    deployment: 'gpt-4o-mini' 
};

// Inizializzazione
document.addEventListener("DOMContentLoaded", function() {
    initUI();
    const savedKey = localStorage.getItem('wep_temp_key');
    if (savedKey) { 
        azureConfig.apiKey = savedKey;
        document.getElementById('temp-key').value = savedKey;
    }
});

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        loadSecrets();
    }
});

function initUI() {
    // Gestione Audience
    const audienceButtons = document.querySelectorAll('.audience-btn');
    audienceButtons.forEach(btn => {
        btn.addEventListener('click', function() {
            audienceButtons.forEach(b => b.classList.remove('active'));
            this.classList.add('active');
            selectedAudience = this.dataset.audience;
        });
    });

    // Gestione Modalità
    const modeRewrite = document.getElementById('mode-rewrite');
    const modeGenerate = document.getElementById('mode-generate');
    const instructionContainer = document.getElementById('instructions-container');

    modeRewrite.addEventListener('click', function() {
        selectedMode = 'REWRITE';
        this.classList.add('active');
        modeGenerate.classList.remove('active');
        instructionContainer.style.display = 'none';
        document.getElementById('btn-text').textContent = "Brand Check";
    });

    modeGenerate.addEventListener('click', function() {
        selectedMode = 'GENERATE';
        this.classList.add('active');
        modeRewrite.classList.remove('active');
        instructionContainer.style.display = 'block';
        document.getElementById('btn-text').textContent = "Genera Email";
    });

    // Impostazioni
    document.getElementById('settings-toggle').onclick = () => {
        const panel = document.getElementById('settings-panel');
        panel.style.display = panel.style.display === 'none' ? 'block' : 'none';
    };

    document.getElementById('save-key-btn').onclick = () => {
        const key = document.getElementById('temp-key').value;
        if (key) {
            azureConfig.apiKey = key;
            localStorage.setItem('wep_temp_key', key);
            alert("Chiave configurata!");
            document.getElementById('settings-panel').style.display = 'none';
        }
    };

    document.getElementById('cta-btn').onclick = handleBrandCheck;
    document.getElementById('apply-btn').onclick = applySuggestion;
}

async function loadSecrets() {
    try {
        const response = await fetch('config/secrets.json');
        const secrets = await response.json();
        azureConfig.apiKey = secrets.azure_openai_key;
        azureConfig.endpoint = secrets.azure_openai_endpoint;
        azureConfig.deployment = secrets.azure_openai_deployment;
    } catch (e) { console.log("Secrets non trovati via fetch."); }
}

async function handleBrandCheck() {
    const text = document.getElementById('text-input').value;
    const instruction = document.getElementById('instruction-input').value;
    const btnLoader = document.getElementById('btn-loader');
    
    if (!azureConfig.apiKey) {
        alert("Configura la API Key cliccando sull'ingranaggio in alto!");
        return;
    }

    document.getElementById('cta-btn').disabled = true;
    btnLoader.style.display = 'inline-block';

    try {
        const response = await callAzureAI(text, instruction);
        document.getElementById('result-text').textContent = response;
        document.getElementById('result-area').style.display = 'block';
    } catch (err) {
        alert("Errore AI: " + err.message);
    } finally {
        document.getElementById('cta-btn').disabled = false;
        btnLoader.style.display = 'none';
    }
}

async function callAzureAI(text, instruction) {
    const systemPrompt = `Sei il WEP Brand Writing Coach. 
    REGOLE: Tono caldo, umano, professionale (Cool older sibling). 
    Frasi brevi, voce attiva, no gergo. 
    AUDIENCE: ${selectedAudience}. Se genitori (Italiano): dai del LEI. 
    MODE: ${selectedMode}. 
    Se email a genitori: Oggetto deve iniziare con 'WEP – '. 
    Mantieni intatti i pattern ##...##.`;

    const url = `${azureConfig.endpoint}/openai/deployments/${azureConfig.deployment}/chat/completions?api-version=2024-02-15-preview`;
    
    const res = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'api-key': azureConfig.apiKey },
        body: JSON.stringify({
            messages: [
                { role: "system", content: systemPrompt },
                { role: "user", content: `Testo: ${text}\nIstruzioni: ${instruction}` }
            ]
        })
    });

    if (!res.ok) throw new Error("Chiamata Azure fallita.");

    const data = await res.json();
    return data.choices[0].message.content;
}

function applySuggestion() {
    const text = document.getElementById('result-text').textContent;
    if (window.Office && Office.context && Office.context.mailbox) {
        Office.context.mailbox.item.body.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text });
    } else {
        navigator.clipboard.writeText(text);
        alert("Testo copiato negli appunti!");
    }
}
