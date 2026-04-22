/*
 * WEP Brand Writing Coach - JS Ultra-Safe
 */

let selectedAudience = 'parents';
let selectedMode = 'REWRITE';
let azureConfig = { 
    apiKey: '', 
    endpoint: 'https://digisup-openai.openai.azure.com', 
    deployment: 'gpt-4o-mini' 
};

// Funzione principale di avvio
function startApp() {
    initUI();
    loadLocalConfig();
    
    if (window.Office) {
        Office.onReady((info) => {
            if (info.host === Office.HostType.Outlook) {
                loadSecrets();
            }
        });
    }
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', startApp);
} else {
    startApp();
}

function initUI() {
    const audienceButtons = document.querySelectorAll('.audience-btn');
    audienceButtons.forEach(btn => {
        btn.onclick = function() {
            audienceButtons.forEach(b => b.classList.remove('active'));
            this.classList.add('active');
            selectedAudience = this.dataset.audience;
        };
    });

    const modeRewrite = document.getElementById('mode-rewrite');
    const modeGenerate = document.getElementById('mode-generate');
    const instructionContainer = document.getElementById('instructions-container');

    if (modeRewrite && modeGenerate) {
        modeRewrite.onclick = function() {
            selectedMode = 'REWRITE';
            modeRewrite.classList.add('active');
            modeGenerate.classList.remove('active');
            if(instructionContainer) instructionContainer.style.display = 'none';
            document.getElementById('btn-text').textContent = "Brand Check";
        };

        modeGenerate.onclick = function() {
            selectedMode = 'GENERATE';
            modeGenerate.classList.add('active');
            modeRewrite.classList.remove('active');
            if(instructionContainer) instructionContainer.style.display = 'block';
            document.getElementById('btn-text').textContent = "Genera Email";
        };
    }

    const settingsToggle = document.getElementById('settings-toggle');
    if (settingsToggle) {
        settingsToggle.onclick = () => {
            const panel = document.getElementById('settings-panel');
            panel.style.display = panel.style.display === 'none' ? 'block' : 'none';
        };
    }

    const saveBtn = document.getElementById('save-key-btn');
    if (saveBtn) {
        saveBtn.onclick = () => {
            azureConfig.apiKey = document.getElementById('temp-key').value;
            azureConfig.endpoint = document.getElementById('temp-endpoint').value;
            azureConfig.deployment = document.getElementById('temp-deploy').value;
            
            localStorage.setItem('wep_config_v2', JSON.stringify(azureConfig));
            alert("Configurazione salvata!");
            document.getElementById('settings-panel').style.display = 'none';
        };
    }

    if(document.getElementById('cta-btn')) document.getElementById('cta-btn').onclick = handleBrandCheck;
    if(document.getElementById('apply-btn')) document.getElementById('apply-btn').onclick = applySuggestion;
}

function loadLocalConfig() {
    try {
        const saved = localStorage.getItem('wep_config_v2');
        if (saved) {
            const parsed = JSON.parse(saved);
            azureConfig = parsed;
            if(document.getElementById('temp-key')) document.getElementById('temp-key').value = parsed.apiKey || '';
            if(document.getElementById('temp-endpoint')) document.getElementById('temp-endpoint').value = parsed.endpoint || '';
            if(document.getElementById('temp-deploy')) document.getElementById('temp-deploy').value = parsed.deployment || '';
        }
    } catch (e) {}
}

async function loadSecrets() {
    try {
        const response = await fetch('config/secrets.json');
        if (response.ok) {
            const secrets = await response.json();
            azureConfig.apiKey = secrets.azure_openai_key;
            azureConfig.endpoint = secrets.azure_openai_endpoint;
            azureConfig.deployment = secrets.azure_openai_deployment;
        }
    } catch (e) {}
}

async function handleBrandCheck() {
    const text = document.getElementById('text-input').value;
    const instruction = document.getElementById('instruction-input').value;
    const btnLoader = document.getElementById('btn-loader');
    
    if (!azureConfig.apiKey) {
        alert("Manca la API Key! Clicca sull'ingranaggio.");
        return;
    }

    document.getElementById('cta-btn').disabled = true;
    if (btnLoader) btnLoader.style.display = 'inline-block';

    try {
        const response = await callAzureAI(text, instruction);
        document.getElementById('result-text').textContent = response;
        document.getElementById('result-area').style.display = 'block';
    } catch (err) {
        alert("Errore AI: " + err.message);
    } finally {
        document.getElementById('cta-btn').disabled = false;
        if (btnLoader) btnLoader.style.display = 'none';
    }
}

async function callAzureAI(text, instruction) {
    const systemPrompt = `Sei il WEP Brand Writing Coach. REGOLE: Tono caldo, umano, professionale. Frasi brevi, voce attiva. AUDIENCE: ${selectedAudience}. Se genitori: dai del LEI. MODE: ${selectedMode}. Mantieni intatti ##...##.`;
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
    if (window.Office && Office.context && Office.context.mailbox && Office.context.mailbox.item) {
        Office.context.mailbox.item.body.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text });
    } else {
        navigator.clipboard.writeText(text);
        alert("Copiato!");
    }
}
