/*
 * WEP Brand Writing Coach - JS con gestione Configurazione
 */

let selectedAudience = 'parents';
let selectedMode = 'REWRITE';
let azureConfig = { 
    apiKey: '', 
    endpoint: 'https://digisup-openai.openai.azure.com', 
    deployment: 'gpt-4o-mini' 
};

document.addEventListener("DOMContentLoaded", function() {
    initUI();
    loadLocalConfig();
});

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        loadSecrets();
    }
});

function initUI() {
    const audienceButtons = document.querySelectorAll('.audience-btn');
    audienceButtons.forEach(btn => {
        btn.onclick = () => {
            audienceButtons.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            selectedAudience = btn.dataset.audience;
        };
    });

    const modeRewrite = document.getElementById('mode-rewrite');
    const modeGenerate = document.getElementById('mode-generate');
    const instructionContainer = document.getElementById('instructions-container');

    modeRewrite.onclick = () => {
        selectedMode = 'REWRITE';
        modeRewrite.classList.add('active');
        modeGenerate.classList.remove('active');
        instructionContainer.style.display = 'none';
        document.getElementById('btn-text').textContent = "Brand Check";
    };

    modeGenerate.onclick = () => {
        selectedMode = 'GENERATE';
        modeGenerate.classList.add('active');
        modeRewrite.classList.remove('active');
        instructionContainer.style.display = 'block';
        document.getElementById('btn-text').textContent = "Genera Email";
    };

    // Impostazioni
    document.getElementById('settings-toggle').onclick = () => {
        const panel = document.getElementById('settings-panel');
        panel.style.display = panel.style.display === 'none' ? 'block' : 'none';
    };

    document.getElementById('save-key-btn').onclick = () => {
        azureConfig.apiKey = document.getElementById('temp-key').value;
        azureConfig.endpoint = document.getElementById('temp-endpoint').value;
        azureConfig.deployment = document.getElementById('temp-deploy').value;
        
        localStorage.setItem('wep_config_v2', JSON.stringify(azureConfig));
        alert("Configurazione salvata!");
        document.getElementById('settings-panel').style.display = 'none';
    };

    document.getElementById('cta-btn').onclick = handleBrandCheck;
    document.getElementById('apply-btn').onclick = applySuggestion;
}

function loadLocalConfig() {
    const saved = localStorage.getItem('wep_config_v2');
    if (saved) {
        const parsed = JSON.parse(saved);
        azureConfig = parsed;
        document.getElementById('temp-key').value = parsed.apiKey || '';
        document.getElementById('temp-endpoint').value = parsed.endpoint || 'https://digisup-openai.openai.azure.com';
        document.getElementById('temp-deploy').value = parsed.deployment || 'gpt-4o-mini';
    }
}

async function loadSecrets() {
    try {
        const response = await fetch('config/secrets.json');
        if (response.ok) {
            const secrets = await response.json();
            azureConfig.apiKey = secrets.azure_openai_key;
            azureConfig.endpoint = secrets.azure_openai_endpoint;
            azureConfig.deployment = secrets.azure_openai_deployment;
            localStorage.setItem('wep_config_v2', JSON.stringify(azureConfig));
        }
    } catch (e) { console.log("Impossibile caricare secrets.json automaticamente."); }
}

async function handleBrandCheck() {
    const text = document.getElementById('text-input').value;
    const instruction = document.getElementById('instruction-input').value;
    const btnLoader = document.getElementById('btn-loader');
    
    if (!azureConfig.apiKey) {
        alert("Inserisci la API Key nelle impostazioni (icona ingranaggio)!");
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

    if (!res.ok) {
        const errData = await res.json();
        throw new Error(errData.error ? errData.error.message : "Fallimento chiamata Azure.");
    }

    const data = await res.json();
    return data.choices[0].message.content;
}

function applySuggestion() {
    const text = document.getElementById('result-text').textContent;
    if (window.Office && Office.context && Office.context.mailbox) {
        Office.context.mailbox.item.body.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text });
    } else {
        navigator.clipboard.writeText(text);
        alert("Testo copiato! Incollalo dove vuoi.");
    }
}
