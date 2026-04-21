/*
 * WEP Brand Writing Coach - Final JS
 */

let selectedAudience = 'parents';
let selectedMode = 'REWRITE';
let azureConfig = { apiKey: '', endpoint: '', deployment: '' };

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        initUI();
        loadSecrets();
    }
});

function initUI() {
    document.querySelectorAll('.audience-btn').forEach(btn => {
        btn.onclick = () => {
            document.querySelectorAll('.audience-btn').forEach(b => b.classList.remove('active'));
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
    };

    modeGenerate.onclick = () => {
        selectedMode = 'GENERATE';
        modeGenerate.classList.add('active');
        modeRewrite.classList.remove('active');
        instructionContainer.style.display = 'block';
    };

    document.getElementById('cta-btn').onclick = handleBrandCheck;
    document.getElementById('apply-btn').onclick = applySuggestion;
}

async function loadSecrets() {
    try {
        const response = await fetch('../digisup agent/config/secrets.json');
        const secrets = await response.json();
        azureConfig.apiKey = secrets.azure_openai_key;
        azureConfig.endpoint = secrets.azure_openai_endpoint;
        azureConfig.deployment = secrets.azure_openai_deployment;
    } catch (e) { console.error("Secrets not found."); }
}

async function handleBrandCheck() {
    const text = document.getElementById('text-input').value;
    const instruction = document.getElementById('instruction-input').value;
    const btnLoader = document.getElementById('btn-loader');
    
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

    const data = await res.json();
    return data.choices[0].message.content;
}

function applySuggestion() {
    const text = document.getElementById('result-text').textContent;
    Office.context.mailbox.item.body.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text });
}
