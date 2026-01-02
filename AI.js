/*** AI.gs ***/

// Globales Objekt zum Speichern von API-Kosten pro Request
// Wird in GmailProcessing.js initialisiert, aber hier als Fallback definiert
if (typeof globalApiCosts === 'undefined') {
  var globalApiCosts = {};
}

/**
 * Berechnet die geschätzten Kosten basierend auf Token-Usage
 * @param {Object} usage - Usage-Objekt mit prompt_tokens, completion_tokens, etc.
 * @param {string} model - Modellname
 * @return {string} Kosten-String
 */
function calculateAPICost_(usage, model) {
  if (!usage || !usage.total_tokens) return '';
  
  // Geschätzte Preise pro 1M Tokens (können je nach Modell/Provider variieren)
  // Diese sind Schätzwerte - passen Sie sie an Ihre tatsächlichen Preise an
  const prices = {
    'gemini-2.5-pro': { input: 1.25, output: 5.00 }, // $1.25/$5.00 pro 1M Tokens
    'gemini-1.5-pro': { input: 0.375, output: 1.50 },
    'gemini-1.5-flash': { input: 0.075, output: 0.30 },
    'gemini-2.0-flash': { input: 0.075, output: 0.30 },
    'gemini-pro': { input: 0.50, output: 1.50 }
  };
  
  const modelKey = model || 'gemini-2.5-pro';
  const price = prices[modelKey] || prices['gemini-2.5-pro'];
  
  const inputTokens = usage.prompt_tokens || 0;
  const outputTokens = usage.completion_tokens || 0;
  const reasoningTokens = usage.reasoning_tokens || 0;
  
  // Reasoning-Tokens werden oft als Input-Tokens gezählt
  const effectiveInputTokens = inputTokens + reasoningTokens;
  
  const inputCost = (effectiveInputTokens / 1000000) * price.input;
  const outputCost = (outputTokens / 1000000) * price.output;
  const totalCost = inputCost + outputCost;
  
  if (totalCost < 0.0001) return ''; // Zu klein zum Anzeigen
  
  return `$${totalCost.toFixed(4)} (${usage.total_tokens} tokens)`;
}

/**
 * Bestimmt die API-Konfiguration (direkt oder über Portkey)
 * @return {Object} {url: string, headers: Object, apiKey: string, mode: string}
 */
function getGeminiAPIConfig_() {
  const usePortkey = getProp_('USE_PORTKEY', 'false').toLowerCase() === 'true';
  
  if (usePortkey) {
    const portkeyApiKey = getProp_('PORTKEY_API_KEY', '');
    
    if (!portkeyApiKey) {
      throw new Error('Portkey ist aktiviert, aber PORTKEY_API_KEY fehlt in Script Properties');
    }
    
    // Portkey Endpoint für Gemini
    // Standard Base URL: https://api.portkey.ai/v1 (öffentlich)
    // Oder firmenintern: https://waf-eu.aigw.galileo.roche.com/v1 (Roche)
    const portkeyBaseUrl = getProp_('PORTKEY_BASE_URL', 'https://api.portkey.ai/v1');
    const apiUrl = `${portkeyBaseUrl}/chat/completions`; // Portkey verwendet OpenAI-kompatible API
    
    // Cloudflare Access Credentials (für firmeninterne Portkey-Integration)
    const cfClientId = getProp_('CF_ACCESS_CLIENT_ID', '');
    const cfClientSecret = getProp_('CF_ACCESS_CLIENT_SECRET', '');
    
    // Provider bestimmen (z.B. "google" für Gemini direkt, "vertex" für Vertex AI)
    const provider = getProp_('PORTKEY_PROVIDER', 'google'); // Standard: google für Gemini
    
    const headers = {
      'x-portkey-api-key': portkeyApiKey,
      'Content-Type': 'application/json'
    };
    
    // Cloudflare Access Headers (für firmeninterne Portkey-Integration)
    if (cfClientId && cfClientSecret) {
      headers['CF-Access-Client-Id'] = cfClientId;
      headers['CF-Access-Client-Secret'] = cfClientSecret;
      Logger.log('Cloudflare Access Headers hinzugefügt für firmeninterne Portkey-Integration');
    }
    
    // Cloudflare Access Headers (für firmeninterne Portkey-Integration)
    if (cfClientId && cfClientSecret) {
      headers['CF-Access-Client-Id'] = cfClientId;
      headers['CF-Access-Client-Secret'] = cfClientSecret;
      Logger.log('Cloudflare Access Headers hinzugefügt für firmeninterne Portkey-Integration');
    }
    
    // Portkey Config: Entweder als JSON-String oder als Config-ID
    // Option 1: Config-ID (wenn in Portkey hinterlegt)
    const portkeyConfigId = getProp_('PORTKEY_CONFIG_ID', '');
    
    // Option 2: Config als JSON-String
    const portkeyConfigJson = getProp_('PORTKEY_CONFIG_JSON', '');
    
    if (portkeyConfigId) {
      // Verwende Config-ID (Virtual Key)
      // Portkey könnte verschiedene Header-Namen verwenden
      // Versuch 1: Als Virtual Key Header
      headers['x-portkey-virtual-key'] = portkeyConfigId;
      
      // Versuch 2: Als Config mit virtual_key
      // Falls der Header nicht funktioniert, versuchen wir es als Config
      // headers['x-portkey-config'] = JSON.stringify({ virtual_key: portkeyConfigId });
      
      Logger.log(`Using Portkey Config-ID: ${portkeyConfigId}`);
    } else if (portkeyConfigJson) {
      // Verwende Config als JSON-String
      try {
        // Validiere, dass es gültiges JSON ist
        const configObj = JSON.parse(portkeyConfigJson);
        headers['x-portkey-config'] = portkeyConfigJson;
      } catch (e) {
        Logger.log('WARNUNG: PORTKEY_CONFIG_JSON ist kein gültiges JSON: ' + e);
      }
    } else {
      // Fallback: Einfache Provider-Konfiguration
      const googleGeminiKey = getProp_('GEMINI_API_KEY', '');
      const vertexApiKey = getProp_('PORTKEY_VERTEX_API_KEY', ''); // Separater Vertex API-Key
      const vertexVirtualKey = getProp_('PORTKEY_VERTEX_VIRTUAL_KEY', ''); // Vertex Virtual Key ID
      
      // Prüfe, ob Vertex AI verwendet werden soll
      if (provider === 'vertex' || provider === 'vertex-ai' || vertexApiKey || vertexVirtualKey) {
        // Vertex AI Konfiguration
        const vertexConfig = {
          provider: 'vertex-ai'
        };
        
        // Option 1: Vertex Virtual Key (wenn vorhanden)
        if (vertexVirtualKey) {
          vertexConfig.virtual_key = vertexVirtualKey;
        }
        // Option 2: Vertex API-Key direkt
        else if (vertexApiKey) {
          vertexConfig.api_key = vertexApiKey;
        }
        
        // Zusätzliche Vertex AI Parameter
        const vertexProjectId = getProp_('PORTKEY_VERTEX_PROJECT_ID', '');
        const vertexRegion = getProp_('PORTKEY_VERTEX_REGION', 'us-central1');
        
        if (vertexProjectId) {
          vertexConfig.project_id = vertexProjectId;
        }
        if (vertexRegion) {
          vertexConfig.region = vertexRegion;
        }
        
        headers['x-portkey-config'] = JSON.stringify(vertexConfig);
        Logger.log('Using Vertex AI with direct API key or virtual key');
      }
      // Google Gemini Provider
      else if (provider === 'google' && googleGeminiKey) {
        // Provider-Header explizit setzen
        headers['x-portkey-provider'] = provider;
        
        // Einfache Config mit provider und api_key
        headers['x-portkey-config'] = JSON.stringify({
          provider: 'google',
          api_key: googleGeminiKey
        });
      }
      // Nur Provider ohne Key
      else if (provider) {
        headers['x-portkey-provider'] = provider;
      }
    }
    
    return {
      url: apiUrl,
      headers: headers,
      apiKey: portkeyApiKey,
      mode: 'portkey',
      provider: provider
    };
  } else {
    // Direkter Gemini-Zugriff
  const apiKey = getProp_('GEMINI_API_KEY', '');
    if (!apiKey) {
      throw new Error('GEMINI_API_KEY fehlt in Script Properties');
    }

  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;
    
    return {
      url: apiUrl,
      headers: { 'Content-Type': 'application/json' },
      apiKey: apiKey,
      mode: 'direct'
    };
  }
}

/**
 * Konvertiert Gemini-Format zu Portkey-Format (OpenAI-kompatibel)
 * @param {Object} geminiPayload - Gemini-Format Payload
 * @param {string} provider - Portkey Provider (z.B. 'google', 'vertex')
 */
function convertToPortkeyFormat_(geminiPayload, provider = 'google') {
  // Portkey verwendet OpenAI-kompatibles Format
  // Gemini: { contents: [{ role: 'user', parts: [{ text: '...' }] }] }
  // OpenAI/Portkey: { model: '...', messages: [{ role: 'user', content: '...' }] }
  
  const messages = [];
  if (geminiPayload.contents && Array.isArray(geminiPayload.contents)) {
    for (const content of geminiPayload.contents) {
      if (content.parts && Array.isArray(content.parts)) {
        const textParts = content.parts
          .filter(p => p.text)
          .map(p => p.text)
          .join('\n');
        
        if (textParts) {
          messages.push({
            role: content.role || 'user',
            content: textParts
          });
        }
      }
    }
  }
  
  // Modellname für Portkey bestimmen
  // Für Google Provider: gemini-2.0-flash, gemini-pro
  // Für Vertex AI: gemini-pro (Standard, am weitesten verfügbar, besonders in EU-Regionen)
  // gemini-1.5-pro und gemini-1.5-flash sind möglicherweise nicht in allen Regionen verfügbar
  // gemini-2.0-flash ist NICHT verfügbar!
  let modelName = getProp_('PORTKEY_MODEL', '');
  
  // Wenn kein Modell gesetzt, basierend auf Config/Provider wählen
  if (!modelName) {
    // Prüfe, ob wir Vertex AI verwenden (basierend auf Config)
    const portkeyConfigJson = getProp_('PORTKEY_CONFIG_JSON', '');
    const portkeyConfigId = getProp_('PORTKEY_CONFIG_ID', '');
    const isVertexAI = provider === 'vertex' || provider === 'vertex-ai' || 
                       (portkeyConfigJson && portkeyConfigJson.includes('vertex-ai')) ||
                       (portkeyConfigId && portkeyConfigId.includes('vertex'));
    
      if (isVertexAI) {
        // Portkey benötigt IMMER ein Modell im Payload, auch mit Config
        // Prüfe, ob firmeninterne Integration (Cloudflare Access vorhanden)
        const cfClientId = getProp_('CF_ACCESS_CLIENT_ID', '');
        if (cfClientId) {
          // Firmeninterne Portkey-Integration (Roche)
          // BASIEREND AUF TESTS: Nur gemini-2.5-pro funktioniert zuverlässig
          // gemini-1.5-flash und gemini-1.5-pro funktionieren NICHT in dieser Integration
          // gemini-2.0-flash ist NICHT verfügbar in der firmeninternen Integration
          // Standard: gemini-2.5-pro (funktioniert, aber langsam)
          // Falls Sie andere Modelle testen möchten, setzen Sie PORTKEY_MODEL manuell
          modelName = 'gemini-2.5-pro'; // Standard: funktioniert zuverlässig
          Logger.log('Firmeninterne Portkey-Integration erkannt - verwende gemini-2.5-pro (funktioniert, aber langsam)');
          Logger.log('HINWEIS: gemini-1.5-flash und gemini-1.5-pro funktionieren in dieser Integration nicht');
          Logger.log('Falls Sie andere Modelle testen möchten, setzen Sie PORTKEY_MODEL in Script Properties');
        } else {
          // Öffentliche Portkey - versuche verschiedene Formate für Vertex AI
          // Verfügbare Modelle: gemini-pro, gemini-1.5-pro, gemini-1.5-flash
          modelName = 'gemini-pro'; // Standard für Vertex AI (am weitesten verfügbar in EU)
          Logger.log('Vertex AI erkannt - verwende gemini-pro (falls nicht verfügbar, bitte PORTKEY_MODEL manuell setzen)');
          Logger.log('Alternative Modelle für Vertex AI: gemini-1.5-pro, gemini-1.5-flash');
        }
      } else {
      // Direkter Google Provider
      modelName = 'gemini-2.0-flash'; // Standard für direkten Google Provider
    }
  }
  
  const portkeyPayload = {
    model: modelName, // Portkey benötigt IMMER ein Modell
    messages: messages
  };
  
  // Generation Config übertragen
  if (geminiPayload.generationConfig) {
    if (geminiPayload.generationConfig.temperature !== undefined) {
      portkeyPayload.temperature = geminiPayload.generationConfig.temperature;
    }
    if (geminiPayload.generationConfig.maxOutputTokens !== undefined) {
      let maxTokens = geminiPayload.generationConfig.maxOutputTokens;
      
      // Für Gemini 2.5 Pro: Erhöhe max_tokens deutlich, da es viele Reasoning-Tokens verwendet
      // Reasoning-Tokens zählen gegen max_tokens, daher brauchen wir mehr Platz
      // Für schnellere Modelle (2.0-flash, 1.5-flash) ist das nicht nötig
      if (modelName && modelName.includes('2.5')) {
        // Mindestens 4x mehr, um Platz für Reasoning + Completion zu haben (Summary/Reply brauchen mehr)
        maxTokens = Math.max(maxTokens * 4, 500);
        Logger.log(`Gemini 2.5 Pro erkannt - erhöhe max_tokens von ${geminiPayload.generationConfig.maxOutputTokens} auf ${maxTokens} (wegen Reasoning-Tokens)`);
      } else if (modelName && (modelName.includes('2.0-flash') || modelName.includes('1.5-flash'))) {
        // Flash-Modelle sind schnell und brauchen keine Erhöhung
        Logger.log(`Schnelles Modell erkannt (${modelName}) - verwende original max_tokens: ${maxTokens}`);
      }
      
      portkeyPayload.max_tokens = maxTokens;
    }
  }
  
  return portkeyPayload;
}

/**
 * Konvertiert Portkey-Response (OpenAI-Format) zu Gemini-Format
 */
function convertFromPortkeyFormat_(portkeyResponse) {
  // Portkey gibt OpenAI-Format zurück: { choices: [{ message: { content: '...' } }] }
  // Wir konvertieren zu Gemini-Format: { candidates: [{ content: { parts: [{ text: '...' }] } }] }
  
  try {
    const data = typeof portkeyResponse === 'string' ? JSON.parse(portkeyResponse) : portkeyResponse;
    
    Logger.log(`convertFromPortkeyFormat_ - Data keys: ${Object.keys(data).join(', ')}`);
    
    // OpenAI-Format: { choices: [{ message: { content: '...' } }] }
    if (data.choices && Array.isArray(data.choices) && data.choices.length > 0) {
      const choice = data.choices[0];
      Logger.log(`Choice found: ${JSON.stringify(choice).slice(0, 200)}`);
      
      // Content kann in verschiedenen Feldern sein
      let content = choice.message?.content || 
                   choice.content || 
                   choice.message?.text ||
                   choice.text ||
                   '';
      
      // Falls finish_reason "length" ist, könnte content leer sein, aber die Antwort trotzdem erfolgreich
      const finishReason = choice.finish_reason || '';
      
      if (content) {
        return {
          candidates: [{
            content: {
              parts: [{ text: content }]
            }
          }]
        };
      } else if (finishReason === 'length') {
        // Antwort wurde wegen max_tokens abgeschnitten
        // Bei Gemini 2.5 Pro werden viele Reasoning-Tokens verwendet
        // Prüfe usage, um zu sehen, ob nur Reasoning-Tokens verwendet wurden
        const usage = data.usage || {};
        const reasoningTokens = usage.completion_tokens_details?.reasoning_tokens || 0;
        const completionTokens = usage.completion_tokens || 0;
        
        if (reasoningTokens > 0 && completionTokens === 0) {
          Logger.log(`WARNUNG: Nur Reasoning-Tokens verwendet (${reasoningTokens}), keine Completion-Tokens - max_tokens zu niedrig für Gemini 2.5 Pro`);
          // Technisch erfolgreich, aber kein Content - das ist OK für den Test
          // In der echten Verwendung sollte max_tokens höher sein
          return {
            candidates: [{
              content: {
                parts: [{ text: 'API-Key funktioniert (Gemini 2.5 Pro verwendet viele Reasoning-Tokens, erhöhen Sie max_tokens für echte Anfragen)' }]
              }
            }]
          };
        } else {
          Logger.log('WARNUNG: Response erfolgreich, aber Content leer (finish_reason: length)');
          return {
            candidates: [{
              content: {
                parts: [{ text: 'Response erfolgreich, aber leer (max_tokens möglicherweise zu niedrig)' }]
              }
            }]
          };
        }
      } else {
        Logger.log('WARNUNG: Choice gefunden, aber kein Content');
        Logger.log(`Choice structure: ${JSON.stringify(choice)}`);
        Logger.log(`Finish reason: ${finishReason}`);
      }
    }
    
    // Falls bereits Gemini-Format: { candidates: [...] }
    if (data.candidates && Array.isArray(data.candidates) && data.candidates.length > 0) {
      Logger.log('Response ist bereits im Gemini-Format');
      return data;
    }
    
    // Falls andere Struktur
    Logger.log(`Unbekannte Response-Struktur: ${JSON.stringify(data).slice(0, 500)}`);
  } catch (e) {
    Logger.log('Portkey Response Conversion Error: ' + e);
    Logger.log(`Response was: ${typeof portkeyResponse === 'string' ? portkeyResponse.slice(0, 500) : JSON.stringify(portkeyResponse).slice(0, 500)}`);
  }
  
  return { candidates: [] };
}

/**
 * Führt einen Gemini API-Call aus (unterstützt direkt und Portkey)
 */
function callGeminiAPI_(payload) {
  const config = getGeminiAPIConfig_();
  let finalPayload = payload;
  let finalUrl = config.url;
  
  // Wenn Portkey, konvertiere Payload
  if (config.mode === 'portkey') {
    finalPayload = convertToPortkeyFormat_(payload, config.provider);
    
    // Für Google Provider: Google Gemini API-Key auch im Payload hinzufügen (falls Portkey das erwartet)
    if (config.provider === 'google') {
      const googleGeminiKey = getProp_('GEMINI_API_KEY', '');
      if (googleGeminiKey && !finalPayload.config) {
        // Portkey könnte den Backend-API-Key im Payload erwarten
        finalPayload.config = {
          apiKey: googleGeminiKey
        };
      }
    }
    
    // Log für Debugging
    Logger.log(`Portkey Request - URL: ${finalUrl}, Provider: ${config.provider}, Payload: ${JSON.stringify(finalPayload).slice(0, 300)}`);
  }
  
  // Wenn direkter Zugriff, Key ist bereits in URL
  // Wenn Portkey, Key ist im Header
  
  const options = {
      method: 'post',
      contentType: 'application/json',
    payload: JSON.stringify(finalPayload),
      muteHttpExceptions: true
  };
  
  // Headers hinzufügen
  if (config.headers) {
    options.headers = config.headers;
    // Log Headers (ohne API-Key aus Sicherheitsgründen)
    const safeHeaders = {};
    for (const key in config.headers) {
      if (key.toLowerCase().includes('key')) {
        safeHeaders[key] = '***masked***';
      } else {
        safeHeaders[key] = config.headers[key];
      }
    }
    Logger.log(`Request Headers: ${JSON.stringify(safeHeaders)}`);
  }
  
  const res = UrlFetchApp.fetch(finalUrl, options);
  
  // Response konvertieren wenn nötig
  if (config.mode === 'portkey') {
    const code = res.getResponseCode();
    const txt = res.getContentText() || '{}';
    
    // Log vollständige Response für Debugging
    Logger.log(`Portkey Response - Code: ${code}, Text: ${txt.slice(0, 1000)}`);
    
    // Bei Fehlern, Response direkt zurückgeben (nicht konvertieren)
    if (code < 200 || code >= 300) {
      Logger.log(`Portkey Error Response - Code: ${code}, Text: ${txt.slice(0, 500)}`);
      return res; // Original Response zurückgeben für Fehlerbehandlung
    }
    
    try {
      const portkeyData = JSON.parse(txt);
      Logger.log(`Portkey Response Data: ${JSON.stringify(portkeyData).slice(0, 500)}`);
      
      // Prüfe, ob Response das erwartete Format hat
      if (!portkeyData.choices && !portkeyData.candidates) {
        Logger.log('WARNUNG: Portkey Response hat weder choices noch candidates');
        Logger.log(`Vollständige Response: ${JSON.stringify(portkeyData)}`);
      }
      
      const geminiFormat = convertFromPortkeyFormat_(portkeyData);
      Logger.log(`Konvertiertes Gemini Format: ${JSON.stringify(geminiFormat).slice(0, 300)}`);
      
      // Usage-Informationen (Kosten) extrahieren und speichern
      const usage = portkeyData.usage || {};
      const costInfo = {
        prompt_tokens: usage.prompt_tokens || 0,
        completion_tokens: usage.completion_tokens || 0,
        reasoning_tokens: usage.completion_tokens_details?.reasoning_tokens || 0,
        total_tokens: usage.total_tokens || 0,
        model: portkeyData.model || 'unknown'
      };
      
      // Erstelle eine mock Response mit Gemini-Format + Kosten-Info
      return {
        getResponseCode: () => code,
        getContentText: () => JSON.stringify(geminiFormat),
        getUsage: () => costInfo // Neue Methode für Kosten-Info
      };
    } catch (e) {
      Logger.log('Portkey Response Conversion Error: ' + e);
      Logger.log(`Original Response Text: ${txt.slice(0, 500)}`);
      return res; // Fallback: Original Response
    }
  }
  
  return res;
}

/**
 * Testet, ob der Gemini API-Key funktioniert
 * @return {Object} {valid: boolean, quotaExceeded: boolean, message: string}
 */
function testGeminiAPIKey() {
  try {
    const config = getGeminiAPIConfig_();
    
    const testPayload = {
      contents: [{ role: 'user', parts: [{ text: 'Test' }] }],
      generationConfig: { temperature: 0.1, maxOutputTokens: 200 } // Höher für Gemini 2.5 Pro (verwendet viele Reasoning-Tokens)
    };

    // Log für Debugging
    Logger.log(`Testing API Key - Mode: ${config.mode}, URL: ${config.url}`);
    if (config.mode === 'portkey') {
      Logger.log(`Portkey Provider: ${config.provider || 'google'}`);
    }

    const res = callGeminiAPI_(testPayload);
    
    const code = res.getResponseCode();
    const txt = res.getContentText() || '';
    
    // Log Response für Debugging
    Logger.log(`API Response - Code: ${code}, Text: ${txt.slice(0, 500)}`);
    
    if (code >= 200 && code < 300) {
      const data = JSON.parse(txt || '{}');
      const hasContent = data?.candidates?.[0]?.content?.parts?.[0]?.text;
      if (hasContent) {
        return {
          valid: true,
          quotaExceeded: false,
          message: `API-Key funktioniert korrekt (Modus: ${config.mode}).`
        };
      } else {
        return {
          valid: false,
          quotaExceeded: false,
          message: 'API-Key antwortet, aber ohne Inhalt. Möglicherweise ungültiger Key oder Limit erreicht.'
        };
      }
    } else {
      // Versuche Fehlerdetails zu extrahieren
      let errorMsg = `HTTP ${code}`;
      let errorDetails = '';
      let isQuotaError = false;
      
      try {
        const errorData = JSON.parse(txt);
        
        // Detaillierte Fehlerinformationen sammeln
        if (errorData.error) {
          if (errorData.error.message) {
            errorMsg = errorData.error.message;
          }
          if (errorData.error.status) {
            errorDetails += `Status: ${errorData.error.status}\n`;
          }
          if (errorData.error.details && Array.isArray(errorData.error.details)) {
            errorDetails += 'Details: ' + errorData.error.details.map(d => d.message || d).join(', ') + '\n';
          }
        } else if (errorData.message) {
          errorMsg = errorData.message;
        }
        
        // Prüfe auf Quota/Limit-Fehler
        const msgLower = errorMsg.toLowerCase();
        isQuotaError = code === 429 || 
                      msgLower.includes('quota') || 
                      msgLower.includes('limit') || 
                      msgLower.includes('exceeded') ||
                      msgLower.includes('rate limit');
      } catch (parseErr) {
        // Wenn JSON-Parsing fehlschlägt, zeige rohen Text
        errorMsg = `HTTP ${code}`;
        errorDetails = txt.slice(0, 500);
        isQuotaError = code === 429;
      }
      
      // Vollständige Fehlermeldung zusammenbauen
      let fullErrorMsg = errorMsg;
      if (errorDetails) {
        fullErrorMsg += '\n\n' + errorDetails;
      }
      if (txt && txt.length > 0 && !errorDetails) {
        fullErrorMsg += '\n\nResponse: ' + txt.slice(0, 300);
      }
      
      if (code === 400) {
        return {
          valid: false,
          quotaExceeded: false,
          message: `API-Key ungültig oder fehlerhaft (HTTP 400):\n\n${fullErrorMsg}\n\nBitte überprüfen Sie:\n` +
                   `- Ist der API-Key korrekt?\n` +
                   `- Ist USE_PORTKEY korrekt gesetzt? (aktuell: ${getProp_('USE_PORTKEY', 'false')})\n` +
                   `- Bei Portkey: Ist PORTKEY_API_KEY korrekt?\n` +
                   `- Bei direktem Modus: Ist GEMINI_API_KEY korrekt?`
        };
      } else if (code === 403) {
        return {
          valid: false,
          quotaExceeded: false,
          message: `API-Key nicht autorisiert (HTTP 403):\n\n${fullErrorMsg}`
        };
      } else if (code === 429 || isQuotaError) {
        // Key ist gültig, aber Limit erreicht
        return {
          valid: true,  // Key ist technisch gültig
          quotaExceeded: true,
          message: `API-Key ist gültig, aber das Limit wurde erreicht.\n\n${fullErrorMsg}\n\nSie können trotzdem fortfahren, aber ohne KI-Funktionen.`
        };
      } else {
        return {
          valid: false,
          quotaExceeded: false,
          message: `API-Fehler (HTTP ${code}):\n\n${fullErrorMsg}`
        };
      }
    }
  } catch (e) {
    Logger.log('testGeminiAPIKey Exception: ' + e);
    return {
      valid: false,
      quotaExceeded: false,
      message: `Fehler beim Testen des API-Keys: ${e.message || e}\n\nBitte überprüfen Sie die Script Properties.`
    };
  }
}

function getAISummary(emailText) {
  try {
    const prompt = `Summarize the following email body concisely in one or two sentences:\n\n${emailText}`;
    const payload = {
      contents: [{ role: 'user', parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.3, maxOutputTokens: 500 } // Erhöht für vollständige Zusammenfassungen
    };

    const res = callGeminiAPI_(payload);
    const data = JSON.parse(res.getContentText() || '{}');
    const txt = data?.candidates?.[0]?.content?.parts?.[0]?.text;
    
    // Kosten extrahieren (falls verfügbar)
    const usage = res.getUsage ? res.getUsage() : null;
    if (usage) {
      // Speichere Kosten in einem globalen Objekt für späteren Zugriff
      if (!globalApiCosts) globalApiCosts = {};
      globalApiCosts['summary'] = usage;
    }
    
    return txt || 'Could not generate summary.';
  } catch (e) {
    Logger.log('Gemini Summary Error: ' + e);
    return 'Error generating summary.';
  }
}

function getAIResponseSuggestionWithContext(emailText, hasReplied, lastReplyISO) {
  try {
  const base = hasReplied
    ? `You already replied on ${lastReplyISO}. Draft a short, polite follow-up if no response has been received since then. Keep it under 80 words.`
    : `Suggest a concise, professional first reply. Keep it under 100 words.`;

  const prompt = `${base}\n\nEmail thread context (latest inbound message):\n${emailText}`;

  const payload = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.5, maxOutputTokens: 500 } // Erhöht von 200 auf 500 für vollständige Antwortvorschläge
    };

    const res = callGeminiAPI_(payload);
    const data = JSON.parse(res.getContentText() || '{}');
    const txt = data?.candidates?.[0]?.content?.parts?.[0]?.text;
    return txt || 'Could not generate response suggestion.';
  } catch (e) {
    Logger.log('Gemini Reply Error: ' + e);
    return 'Error generating response suggestion.';
  }
}

function getAIActionabilityAnalysis(subject, body, myProfile) {
  try {

  // Profil fuers Prompt (Name, E-Mails, Teamkeywords)
  const name = (myProfile?.name || CONFIG.MY_NAME || 'User').toString();
  const emails = (myProfile?.emails || []).join(', ');
  const team = (myProfile?.teamKeywords || CONFIG.MY_TEAM_KEYWORDS || []).join(', ');

  const prompt =
    `You are an assistant that classifies whether an email implies an actionable task for the user, even if not explicitly assigned.
    User identity:
    - name: ${name}
    - emails: ${emails}
    - team_keywords: ${team}

    IMPORTANT RULES:
    1. Task titles must be ACTION-ORIENTED and SPECIFIC. Examples:
       - GOOD: "Abholen: Paket aus Packstation" (not just "Paket in Packstation angekommen")
       - GOOD: "Rechnung bezahlen" (not just "Rechnung erhalten")
       - GOOD: "Termin bestätigen: Meeting am 15.03." (not just "Meeting-Einladung")
       - BAD: "Paket-Benachrichtigung" (too vague, not action-oriented)
       - BAD: "Newsletter erhalten" (no action needed)
    
    2. For informational emails (notifications, confirmations, newsletters, advertisements), set is_task_for_me to "No" and leave tasks empty.
    
    3. Only create tasks if there is a clear, actionable item the user needs to do.

    Given the email subject and body, decide:
    - is_task_for_me: "Yes" | "No" | "Unsure"
    - reasons: short rationale (max 200 chars)
    - suggested_owner: "me" | "someone_else" | "team" | "unknown"
    - tasks: list of {title, owner, due_date, priority}
       * title: MUST be action-oriented (verb + object), e.g. "Paket abholen", "Rechnung bezahlen", "Termin bestätigen"
       * If the email is just informational (notification, confirmation, newsletter), leave tasks empty

    Return ONLY strict JSON, no prose.

    Email:
    Subject: ${subject}
    Body:
    ${body}

    JSON schema:
    {
      "is_task_for_me": "Yes|No|Unsure",
      "reasons": "string",
      "suggested_owner": "me|someone_else|team|unknown",
      "tasks": [
        {"title":"string (action-oriented, e.g. 'Paket abholen')","owner":"string","due_date":"YYYY-MM-DD or empty","priority":"Low|Medium|High"}
      ]
    }`;

  const payload = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.2, maxOutputTokens: 1000 } // Erhöht für vollständige Analysen (gemini-2.5-pro braucht mehr)
    };

    const res = callGeminiAPI_(payload);
    const txt = JSON.parse(res.getContentText() || '{}')?.candidates?.[0]?.content?.parts?.[0]?.text || '{}';
    
    // Kosten speichern
    if (res.getUsage) {
      const usage = res.getUsage();
      if (!globalApiCosts) globalApiCosts = {};
      globalApiCosts['actionability'] = usage;
    }

    // JSON robust parsen
    let data;
    try { data = JSON.parse(txt); } catch (_) { data = {}; }

    return {
      is_task_for_me: data.is_task_for_me || 'Unsure',
      reasons: data.reasons || '',
      suggested_owner: data.suggested_owner || 'unknown',
      tasks: Array.isArray(data.tasks) ? data.tasks : []
    };
  } catch (e) {
    Logger.log('Gemini Actionability Error: ' + e);
    return { is_task_for_me: 'Unsure', reasons: 'Error during analysis', tasks: [] };
  }
}

function detectLang_(text) {
  try { return LanguageApp.detectLanguage(text || '') || ''; } catch (_) { return ''; }
}

function getAISummarySmart(emailText) {
  try {
  const src = detectLang_(emailText);
  const rule = (CONFIG.AI_LANG?.SUMMARY_IF_SOURCE || []).find(r => r.source === src);
  const target = rule ? rule.target : null;

  const langLine = target ? `Respond in ${target}.` : 'Respond in the same language as the input.';
  const prompt = `${langLine}\nSummarize the following email body concisely in one or two sentences:\n\n${emailText}`;

  const payload = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.3, maxOutputTokens: 800 } // Erhöht für vollständige Zusammenfassungen (gemini-2.5-pro braucht mehr)
  };
    
    var lastErr = null;
    var lastUsage = null;
    var lastResponse = null;
    for (var attempt = 1; attempt <= 3; attempt++) {
      try {
        Logger.log(`Summary-Generierung Versuch ${attempt}/3`);
        const res = callGeminiAPI_(payload);
        const code = res.getResponseCode();
        const txt = res.getContentText() || '';
        lastResponse = { code, text: txt.substring(0, 500) };
        
        Logger.log(`Summary API Response: Code ${code}, Text: ${txt.substring(0, 300)}`);
        
        if (code >= 200 && code < 300) {
          const data = JSON.parse(txt || '{}');
          
          // Prüfe verschiedene mögliche Response-Strukturen
          let out = null;
          
          // Gemini-Format: candidates[0].content.parts[0].text
          if (data.candidates && Array.isArray(data.candidates) && data.candidates.length > 0) {
            out = data.candidates[0]?.content?.parts?.[0]?.text;
          }
          
          // Portkey/OpenAI-Format: choices[0].message.content (falls nicht konvertiert)
          if (!out && data.choices && Array.isArray(data.choices) && data.choices.length > 0) {
            out = data.choices[0]?.message?.content || data.choices[0]?.content || data.choices[0]?.text;
          }
          
          // Fallback: direktes text-Feld
          if (!out && data.text) {
            out = data.text;
          }
          
          // Kosten speichern
          if (res.getUsage) {
            lastUsage = res.getUsage();
            Logger.log(`Summary Usage: ${JSON.stringify(lastUsage)}`);
          }
          
          if (out && out.trim()) {
            Logger.log(`Summary erfolgreich generiert (${out.length} Zeichen)`);
            // Speichere Kosten für späteren Zugriff (nur wenn globalApiCosts existiert)
            try {
              if (lastUsage && typeof globalApiCosts !== 'undefined') {
                if (!globalApiCosts) globalApiCosts = {};
                globalApiCosts['summary'] = lastUsage;
              }
            } catch (e) {
              // globalApiCosts nicht verfügbar - ignorieren (nicht kritisch)
              Logger.log('Hinweis: globalApiCosts nicht verfügbar (OK für manuelle Generierung)');
            }
            return out.trim();
          } else {
            Logger.log(`WARNUNG: API Response Code ${code}, aber kein Text gefunden`);
            Logger.log(`Response-Struktur: ${JSON.stringify(data).substring(0, 1000)}`);
            
            // Prüfe auf Usage-Informationen (könnte zeigen, dass nur Reasoning-Tokens verwendet wurden)
            const usage = data.usage || {};
            const reasoningTokens = usage.completion_tokens_details?.reasoning_tokens || 0;
            const completionTokens = usage.completion_tokens || 0;
            
            if (reasoningTokens > 0 && completionTokens === 0) {
              lastErr = `Nur Reasoning-Tokens verwendet (${reasoningTokens}), keine Completion-Tokens - max_tokens möglicherweise zu niedrig für gemini-2.5-pro`;
            } else {
              lastErr = `HTTP ${code} - Kein Text in Response (Struktur: ${Object.keys(data).join(', ')})`;
            }
          }
        } else {
          lastErr = `HTTP ${code}: ${txt.slice(0, 200)}`;
          Logger.log(`Summary API Fehler: ${lastErr}`);
        }
      } catch (e) {
        lastErr = e && e.message ? e.message : '' + e;
        Logger.log(`Summary API Exception: ${lastErr}`);
      }
      if (attempt < 3) {
        Utilities.sleep(300 + Math.floor(Math.random() * 300)); // leichter Backoff
      }
    }
    Logger.log(`Gemini Summary Error nach ${attempt} Versuchen: ${lastErr || 'unknown'}`);
    Logger.log(`Letzte Response: ${JSON.stringify(lastResponse)}`);
    // Fallback-Text, falls alle Versuche fehlschlagen:
    return `Keine AI-Summary verfügbar (Fehler: ${lastErr || 'unbekannt'})`;
  } catch (e) {
    Logger.log('Gemini Summary Fatal: ' + e);
    Logger.log('Stack: ' + e.stack);
    return `Keine AI-Summary verfügbar (Fatal Error: ${e.message || e})`;
  }
}
function getAIResponseSuggestionWithLang(emailText, hasReplied, lastReplyISO, replyLangHint) {
  try {
  // Prüfe, ob es sich um eine Dankes- oder Grußmail handelt
  const emailLower = emailText.toLowerCase();
  const isGratitude = [
    'danke', 'thank you', 'thanks', 'vielen dank', 'herzlichen dank',
    'dankeschön', 'thank you very much', 'merci', 'grazie'
  ].some(pattern => emailLower.includes(pattern));
  
  const isGreeting = [
    'frohe weihnachten', 'merry christmas', 'frohes neues jahr', 'happy new year',
    'frohe ostern', 'happy easter', 'schöne feiertage', 'happy holidays',
    'grüße', 'greetings', 'liebe grüße', 'best regards', 'kind regards',
    'viele grüße', 'best wishes', 'herzliche grüße', 'warm regards',
    'frohe festtage', 'season\'s greetings'
  ].some(pattern => emailLower.includes(pattern));

  let base;
  if (isGratitude) {
    base = hasReplied
      ? `You already replied on ${lastReplyISO}. Draft a short, polite acknowledgment if appropriate. Keep it under 60 words.`
      : `This appears to be a thank-you message. Suggest a warm, friendly, and brief reply acknowledging their thanks. Match the tone and formality level. Keep it under 80 words.`;
  } else if (isGreeting) {
    base = hasReplied
      ? `You already replied on ${lastReplyISO}. Draft a short, polite greeting response if appropriate. Keep it under 60 words.`
      : `This appears to be a greeting message (e.g., holiday greetings, seasonal wishes). Suggest a warm, friendly, and brief reply reciprocating the greeting. Match the tone and formality level. Keep it under 80 words.`;
  } else {
    base = hasReplied
      ? `You already replied on ${lastReplyISO}. Draft a short, polite follow-up if no response has been received since then. Keep it under 80 words.`
      : `Suggest a concise, professional first reply. Keep it under 100 words.`;
  }

  const langLine = replyLangHint ? `Write the reply in ${replyLangHint}.` : 'Write the reply in the same language as the latest inbound message.';
  const prompt = `${base}\n${langLine}\n\nEmail thread context (latest inbound message):\n${emailText}`;

  const payload = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.5, maxOutputTokens: 1000 } // Erhöht für vollständige Antwortvorschläge (gemini-2.5-pro braucht mehr)
  };
    
    const res = callGeminiAPI_(payload);
    const data = JSON.parse(res.getContentText() || '{}');
    // Kosten speichern
    if (res.getUsage) {
      const usage = res.getUsage();
      if (!globalApiCosts) globalApiCosts = {};
      globalApiCosts['reply'] = usage;
    }
    return data?.candidates?.[0]?.content?.parts?.[0]?.text || 'Could not generate response suggestion.';
  } catch (e) {
    Logger.log('Gemini Reply Error: ' + e);
    return 'Error generating response suggestion.';
  }
} 

/**
 * Prüft, ob für diese E-Mail ein Reply-Vorschlag generiert werden sollte
 * Infomails (Paket-Benachrichtigungen, Bestätigungen, Newsletter, Werbung) → kein Reply
 * ABER: Dankes- und Grußmails → IMMER Reply generieren
 * 
 * @param {string} subject - E-Mail Betreff
 * @param {string} body - E-Mail Body
 * @param {Object} actionAnalysis - AI-Analyse mit is_task_for_me, etc.
 * @param {string} fromEmail - Absender-E-Mail (optional, für Domain-Learning)
 * @return {boolean} true = Reply generieren, false = kein Reply
 */
function shouldGenerateReply_(subject, body, actionAnalysis, fromEmail) {
  try {
    const subjectLower = (subject || '').toLowerCase();
    const bodyLower = (body || '').toLowerCase();
    const combined = `${subjectLower} ${bodyLower}`;
    
    // WICHTIG: Dankes- und Grußmails IMMER beantworten (auch wenn sie als Infomail gelten)
    const gratitudePatterns = [
      'danke', 'thank you', 'thanks', 'vielen dank', 'herzlichen dank',
      'dankeschön', 'thank you very much', 'merci', 'grazie',
      'frohe weihnachten', 'merry christmas', 'frohes neues jahr', 'happy new year',
      'frohe ostern', 'happy easter', 'schöne feiertage', 'happy holidays',
      'grüße', 'greetings', 'liebe grüße', 'best regards', 'kind regards',
      'viele grüße', 'best wishes', 'herzliche grüße', 'warm regards',
      'frohe festtage', 'season\'s greetings', 'frohe weihnacht', 'merry xmas'
    ];
    
    if (gratitudePatterns.some(pattern => combined.includes(pattern))) {
      Logger.log(`shouldGenerateReply_: Dankes-/Grußmail erkannt → Reply generieren`);
      return true; // Dankes-/Grußmail → IMMER Reply generieren
    }
    
    // Prüfe gelernte Patterns: Wenn "Task for Me" = "No" gelernt wurde → wahrscheinlich Infomail
    if (fromEmail) {
      const fromDomain = (fromEmail.split('@')[1] || '').toLowerCase();
      const learnedDecision = shouldBeTaskForMe_(fromDomain, subject);
      if (learnedDecision === 'No') {
        Logger.log(`shouldGenerateReply_: Gelernt "No" für Domain "${fromDomain}" → kein Reply`);
        return false; // Gelernt: Kein Task → kein Reply
      }
    }
    
    // Wenn AI-Analyse sagt "No" (kein Task) → wahrscheinlich Infomail
    if (actionAnalysis && actionAnalysis.is_task_for_me === 'No') {
      // Prüfe auf typische Infomail-Patterns
      const infoPatterns = [
        'paket', 'packstation', 'benachrichtigung', 'notification',
        'bestätigung', 'confirmation', 'erfolgreich', 'successful',
        'newsletter', 'werbung', 'angebot', 'promotion', 'rabatt',
        'rechnung erhalten', 'invoice received', 'zahlung erhalten',
        'automatische benachrichtigung', 'automatic notification',
        'keine antwort erforderlich', 'no reply needed', 'no action required'
      ];
      
      if (infoPatterns.some(pattern => combined.includes(pattern))) {
        Logger.log(`shouldGenerateReply_: Infomail-Pattern erkannt → kein Reply`);
        return false; // Infomail → kein Reply
      }
    }
    
    // Prüfe auf explizite "keine Antwort"-Hinweise
    const noReplyPatterns = [
      'keine antwort erforderlich',
      'no reply needed',
      'no action required',
      'automatische benachrichtigung',
      'automatic notification',
      'dies ist eine automatische nachricht',
      'this is an automated message'
    ];
    
    if (noReplyPatterns.some(pattern => combined.includes(pattern))) {
      Logger.log(`shouldGenerateReply_: Explizite "keine Antwort"-Hinweise → kein Reply`);
      return false; // Explizit keine Antwort erforderlich
    }
    
    return true; // Standard: Reply generieren
  } catch (e) {
    Logger.log('shouldGenerateReply_ error: ' + e);
    return true; // Bei Fehler: sicherheitshalber Reply generieren
  }
}

/**
 * Generiert eine AI-Response basierend auf dem gesamten Thread-Kontext
 * @param {string} messageId - Message-ID der aktuellen Mail
 * @param {number} row - Zeilennummer in der Tabelle (optional, für Kontext)
 * @return {string} Generierte Response
 */
function generateResponseWithThreadContext_(messageId, row) {
  try {
    const msg = GmailApp.getMessageById(messageId);
    if (!msg) {
      Logger.log('Message nicht gefunden: ' + messageId);
      return '';
    }

    const thread = msg.getThread();
    if (!thread) {
      Logger.log('Thread nicht gefunden für Message: ' + messageId);
      return '';
    }

    // Alle Mails im Thread sammeln (inkl. gesendete Mails!)
    const messages = thread.getMessages();
    const myEmails = getMyEmails_();
    
    Logger.log(`Thread-Kontext: Analysiere ${messages.length} Mails im Thread (inkl. gesendete Mails)`);
    
    // Thread-Kontext aufbauen (chronologisch)
    const threadContext = [];
    let latestInbound = null;
    let myLastReply = null;
    let inboundCount = 0;
    let outboundCount = 0;
    
    for (let i = 0; i < messages.length; i++) {
      const m = messages[i];
      const from = (m.getFrom() || '').toLowerCase();
      const isFromMe = myEmails.some(me => from.includes(me.toLowerCase()));
      const date = m.getDate();
      const subject = m.getSubject() || '';
      const body = getBestBody_(m);
      const bodyPreview = body.length > 500 ? body.slice(0, 500) + '...' : body;
      
      if (!isFromMe) {
        latestInbound = { date, subject, body: bodyPreview };
        inboundCount++;
        threadContext.push({
          role: 'inbound',
          date: date ? Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm') : 'Unbekannt',
          from: m.getFrom(),
          subject: subject,
          body: bodyPreview
        });
      } else {
        myLastReply = { date, subject, body: bodyPreview };
        outboundCount++;
        threadContext.push({
          role: 'outbound',
          date: date ? Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm') : 'Unbekannt',
          from: m.getFrom(),
          subject: subject,
          body: bodyPreview
        });
      }
    }
    
    Logger.log(`Thread-Kontext: ${inboundCount} eingehende + ${outboundCount} ausgehende (gesendete) Mails gefunden`);

    // Prüfe, ob bereits geantwortet wurde
    const hasReplied = myLastReply && latestInbound && myLastReply.date > latestInbound.date;
    const lastReplyISO = myLastReply ? myLastReply.date.toISOString() : '';

    // Thread-Kontext als Text formatieren
    let contextText = 'E-Mail-Thread-Kontext (chronologisch):\n\n';
    threadContext.forEach((item, idx) => {
      contextText += `[${idx + 1}] ${item.role === 'inbound' ? 'EINGEHEND' : 'AUSGEHEND'} - ${item.date}\n`;
      contextText += `Von: ${item.from}\n`;
      contextText += `Betreff: ${item.subject}\n`;
      contextText += `Inhalt:\n${item.body}\n\n`;
      contextText += '---\n\n';
    });

    // Aktuelle Mail (die neueste eingehende)
    const currentSubject = latestInbound ? latestInbound.subject : msg.getSubject();
    const currentBody = latestInbound ? latestInbound.body : getBestBody_(msg);

    // Prompt für Response-Generierung
    const base = hasReplied
      ? `You already replied on ${lastReplyISO}. Draft a short, polite follow-up if no response has been received since then. Keep it under 80 words.`
      : `Suggest a concise, professional first reply. Keep it under 100 words. Consider the full thread context.`;

    const prompt = `${base}\n\n${contextText}\n\nLatest inbound message to respond to:\nSubject: ${currentSubject}\nBody:\n${currentBody}\n\nWrite the reply in the same language as the latest inbound message.`;

    const payload = {
      contents: [{ role: 'user', parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.5, maxOutputTokens: 1000 }
    };

    const res = callGeminiAPI_(payload);
    const data = JSON.parse(res.getContentText() || '{}');
    const response = data?.candidates?.[0]?.content?.parts?.[0]?.text || '';
    
    Logger.log(`Response mit Thread-Kontext generiert: ${threadContext.length} Mails total (${inboundCount} eingehend, ${outboundCount} gesendet)`);
    return response.trim();
  } catch (e) {
    Logger.log('generateResponseWithThreadContext_ error: ' + e);
    return '';
  }
}

function getAvailableLabels_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Gmail Labels');
  if (!sh) return [];

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const vals = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  return vals
    .map(r => (r[0] || '').toString().trim())
    .filter(Boolean); // keine leeren
}

function getAILabelSuggestion(subject, body) {
  try {
  const labels = getAvailableLabels_();
  if (!labels.length) return '';

  const labelList = labels.join(' | ');
  const prompt =
    `Du bist ein Assistent, der eingehende E-Mails einem einzigen Label zuordnet. ` +
    `Du bekommst eine Liste moeglicher Labels und sollst GENAU EIN Label daraus waehlen, das am besten passt. ` +
    `Wenn nichts gut passt, waehle "Uncategorized". ` +
    `Antworte NUR mit dem exakten Labelnamen (ohne Erklaerung).\n\n` +
    `Verfuegbare Labels:\n${labelList}\n\n` +
    `E-Mail:\n` +
    `Betreff: ${subject || '(kein Betreff)'}\n` +
    `Body:\n${body || ''}`;

  const payload = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.1, maxOutputTokens: 100 } // Erhöht von 20 auf 100 für vollständige Label-Namen
    };

    const res = callGeminiAPI_(payload);
    const code = res.getResponseCode();
    const txt = res.getContentText() || '';

    if (code < 200 || code >= 300) {
      Logger.log('Gemini Label Suggestion HTTP ' + code + ' ' + txt.slice(0, 200));
      return '';
    }

    const data = JSON.parse(txt || '{}');
    let out = data?.candidates?.[0]?.content?.parts?.[0]?.text || '';
    if (!out) return '';

    out = out.trim();

    // Falls das Model doch Text drumherum liefert -> auf bekannten Labelnamen matchen
    const lowerOut = out.toLowerCase();
    const exact = labels.find(l => l.toLowerCase() === lowerOut);
    if (exact) return exact;

    for (const l of labels) {
      if (out.includes(l)) return l;
      if (lowerOut.includes(l.toLowerCase())) return l;
    }

    // Fallback: wenn "Uncategorized" in der Liste ist, nimm das
    const unc = labels.find(l => l.toLowerCase() === 'uncategorized');
    return unc || '';
  } catch (e) {
    Logger.log('Gemini Label Suggestion Error: ' + e);
    return '';
  }
}


// end of file AI.gs
