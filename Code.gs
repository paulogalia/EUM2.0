/**
 * ============================================================================
 * EUM MANAGER 3.0 - CÉREBRO COMPLETO (VERSÃO FINAL INTEGRADA)
 * * Módulos Integrados:
 * 1. Wizard Mágico (IA Cascata 2.5 -> 2.0 -> 1.5)
 * 2. Explorador com Motor Clínico (Cálculo Dinâmico de Colunas)
 * 3. Motores Clínicos (Renal BSA + Posologia NLP)
 * 4. Diagnóstico de IA e Gestão de Estado Single-Tenant
 * ============================================================================
 */

// --- CONSTANTES GLOBAIS ---
const MASTER_DB_ID = "1zKqYVR9seTPy3eyX5CR2xuXZtqSwwvERQn_KX3GK5JM"; // Opcional
const GEMINI_API_KEY = "AIzaSyDBAuw8Q99IWNY1s9cZWv37UJJo9DJaCGo"; // Insira sua chave aqui

// ============================================================================
// 1. INICIALIZAÇÃO E MENU
// ============================================================================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🚀 EUM App')
    .addItem('Abrir Painel de Controle', 'abrirDashboard')
    .addSeparator()
    .addItem('🕵️ Diagnosticar Erros', 'diagnosticarErroReferencia')
    .addToUi();
}

function abrirDashboard() {
  const html = HtmlService.createTemplateFromFile('Dashboard').evaluate()
    .setTitle('EUM Manager 3.0')
    .setWidth(1200)
    .setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(html, 'EUM Manager');
}

function diagnosticarErroReferencia() {
  const ui = SpreadsheetApp.getUi();
  ui.alert("Diagnóstico", "Verifique se as abas 'Exames_Referencia' e 'Doses_Referencia' existem.", ui.ButtonSet.OK);
}

// ============================================================================
// 2. API DE CONFIGURAÇÃO E ESTADO (SINGLE-TENANT)
// ============================================================================

function apiGetInitialState() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().map(function(s) { return s.getName(); });
  const props = PropertiesService.getScriptProperties();
  
  // Carrega a configuração ÚNICA salva
  const savedConfig = JSON.parse(props.getProperty('EUM_CONFIG_MASTER') || '{}');
  
  const status = {
    hasCore: !!(savedConfig?.core?.abaFatos),
    hasExames: !!(savedConfig?.exames?.aba),
    hasRams: !!(savedConfig?.rams?.aba)
  };
  
  return { sheets: sheets, config: savedConfig, status: status };
}

function apiSaveConfig(newConfig) {
  try {
    // Preserva a análise da IA se não for enviada na atualização
    const oldConfig = JSON.parse(PropertiesService.getScriptProperties().getProperty('EUM_CONFIG_MASTER') || '{}');
    if (!newConfig.aiAnalysis && oldConfig.aiAnalysis) {
      newConfig.aiAnalysis = oldConfig.aiAnalysis;
    }

    // SINGLE-TENANT: Sobrescreve a chave única
    PropertiesService.getScriptProperties().setProperty('EUM_CONFIG_MASTER', JSON.stringify(newConfig));
    
    return { 
      sucesso: true, 
      status: {
        hasCore: !!(newConfig.core.abaFatos),
        hasExames: !!(newConfig.exames.active),
        hasRams: !!(newConfig.rams.active)
      }
    };
  } catch (e) { 
    return { sucesso: false, erro: e.message };
  }
}

// ============================================================================
// 3. WIZARD MÁGICO (IA COM CASCATA OTIMIZADA)
// ============================================================================

function apiMagicSetup(pdfBase64, matrixDados, nomeArquivo) {
  console.log("INICIO: apiMagicSetup iniciado para: " + nomeArquivo);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. CRIAR ABA E IMPORTAR DADOS
  let sheetName = "Dados_" + (nomeArquivo || "Import").replace(/[^a-zA-Z0-9]/g, "_").substring(0, 15);
  if (ss.getSheetByName(sheetName)) sheetName += "_" + Math.floor(Math.random()*1000);
  
  const sheet = ss.insertSheet(sheetName);
  
  if (matrixDados && matrixDados.length > 0) {
    sheet.getRange(1, 1, matrixDados.length, matrixDados[0].length).setValues(matrixDados);
  } else {
    return { erro: "Arquivo de dados vazio." };
  }

  const csvHeaders = matrixDados[0];
  console.log("PROGRESSO: Aba criada. Iniciando IA...");

  // 2. ANÁLISE DE IA
  if (!GEMINI_API_KEY || GEMINI_API_KEY.includes("AQUI")) return { erro: "Chave API inválida." };

  const prompt = `
    Atue como Consultor Sênior EUM Manager.
    INPUTS: 
    1. Cabeçalhos dos Dados (Aba '${sheetName}'): ${JSON.stringify(csvHeaders)}
    2. Protocolo do Estudo (PDF anexo).
    
    TAREFA:
    Analise a compatibilidade e mapeie colunas para as variáveis internas.
    Crie um checklist de capacidades vs lacunas (gaps).
    
    RETORNE APENAS JSON VÁLIDO (Sem markdown):
    {
      "studyName": "Nome Curto do Estudo",
      "studySummary": "Resumo executivo em 1 frase",
      "preset": "safety" | "efficacy" | "protocol" | "full",
      "analysis": { 
         "capabilities": ["Item 1", "Item 2"], 
         "gaps": ["Falta X", "Falta Y"] 
      },
      "mapping": { 
         "colProntFatos": "NomeColuna", "colMed": "NomeColuna", "colDtIni": "NomeColuna",
         "colDoseUni": "NomeColuna", "colApraz": "NomeColuna", "colDose24h": "NomeColuna",
         "colPeso": "NomeColuna", "colAltura": "NomeColuna", "colCreat": "NomeColuna",
         "colProntDim": "NomeColuna", "colNasc": "NomeColuna", "colSexo": "NomeColuna"
      }
    }
  `;

  const payload = {
    contents: [{
      parts: [
        { text: prompt },
        { inline_data: { mime_type: "application/pdf", data: pdfBase64 } }
      ]
    }]
  };

  // ESTRATÉGIA DE CASCATA: Tenta o mais moderno -> mais rápido -> mais estável
  const modelos = ["gemini-2.5-flash", "gemini-2.0-flash-exp", "gemini-1.5-flash", "gemini-1.5-pro"];

  for (let i = 0; i < modelos.length; i++) {
    let modelo = modelos[i];
    try {
      console.log(`TENTATIVA IA: Usando modelo ${modelo}...`);
      
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelo}:generateContent?key=${GEMINI_API_KEY}`;
      const options = {
        method: 'post', contentType: 'application/json',
        payload: JSON.stringify(payload), muteHttpExceptions: true
      };
      
      const resp = UrlFetchApp.fetch(url, options);
      const content = JSON.parse(resp.getContentText());
      
      if (content.error) throw new Error(content.error.message);
      
      const txt = content.candidates[0].content.parts[0].text;
      
      // PARSER ROBUSTO (Ignora texto antes/depois do JSON)
      const jsonStart = txt.indexOf('{');
      const jsonEnd = txt.lastIndexOf('}');
      
      if(jsonStart === -1 || jsonEnd === -1) throw new Error("IA não retornou JSON válido.");
      
      const jsonString = txt.substring(jsonStart, jsonEnd + 1);
      const resultadoIA = JSON.parse(jsonString);
      
      // Adiciona o nome da aba criada ao resultado
      resultadoIA.createdSheet = sheetName;
      
      console.log("SUCESSO: JSON processado com " + modelo);
      return resultadoIA;

    } catch (e) { 
      console.error(`ERRO IA (${modelo}): ` + e.message);
      // Loop continua para o próximo modelo
    }
  }
  
  return { erro: "Falha na análise IA. A aba '" + sheetName + "' foi criada, mas sem mapeamento." };
}

// --- 4. HELPERS IA (DIAGNÓSTICO E UTILITÁRIOS) ---

function apiCheckGeminiModels() {
  if (!GEMINI_API_KEY) return { sucesso: false, erro: "Sem Chave" };
  try {
    const res = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/models?key=${GEMINI_API_KEY}`, {muteHttpExceptions:true});
    const json = JSON.parse(res.getContentText());
    
    if(json.error) throw new Error(json.error.message);
    
    const available = json.models
      .filter(m => m.supportedGenerationMethods && m.supportedGenerationMethods.includes("generateContent"))
      .map(m => m.name.replace("models/", ""))
      .sort();
      
    return { sucesso: true, modelos: available };
  } catch (e) { return { sucesso: false, erro: e.message }; }
}

function apiInterpretarIntencaoEstudo(texto) {
  const prompt = `Consultor EUM. Usuário: "${texto}". Retorne JSON: {"preset": "full", "nome_sugerido": "Estudo", "resumo": "...", "explicacao": "..."}`;
  return chamarGeminiSimples_(prompt);
}

function apiInterpretarGraficoIA(pergunta, colunasDisponiveis) {
  // Injeta colunas virtuais no contexto da IA
  colunasDisponiveis.push("CALC_RENAL_TFG", "CALC_DOSE_KG", "CALC_IDADE");
  
  const prompt = `
    Analista Plotly. 
    Colunas Disponíveis: ${JSON.stringify(colunasDisponiveis)}. 
    Pergunta: "${pergunta}".
    
    Se a pergunta envolver função renal, dose por peso ou idade, use as colunas CALC_*.
    
    RETORNE JSON: 
    { "viavel": true, "config": { "x": "col", "y": "col", "type": "bar|line|scatter", "aggregation": "count|avg" }, "explicacao": "..." }
  `;
  return chamarGeminiSimples_(prompt);
}

function chamarGeminiSimples_(prompt) {
  const modelos = ["gemini-2.0-flash-exp", "gemini-1.5-flash"];
  for (let i = 0; i < modelos.length; i++) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelos[i]}:generateContent?key=${GEMINI_API_KEY}`;
      const opts = { method: 'post', contentType: 'application/json', payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }), muteHttpExceptions: true };
      const res = UrlFetchApp.fetch(url, opts);
      const json = JSON.parse(res.getContentText());
      
      if (json.error) throw new Error(json.error.message);
      const txt = json.candidates[0].content.parts[0].text;
      const jsonStart = txt.indexOf('{');
      const jsonEnd = txt.lastIndexOf('}');
      
      return JSON.parse(txt.substring(jsonStart, jsonEnd+1));
    } catch (e) { console.log(e); }
  }
  return { erro: "Erro IA Texto." };
}

// ============================================================================
// 5. MOTORES CLÍNICOS (RENAL + POSOLOGIA AVANÇADA)
// ============================================================================

function calcularFuncaoRenal_(creat, idadeAnos, sexo, peso, altura) {
  if (!creat || isNaN(creat) || creat <= 0) return { valor: null, metodo: "Sem Creatinina", estagio: "N/D" };
  
  let resultado = 0, metodo = "";
  let bsa = 0, clcr_absoluto = null;

  // 1. Calcula BSA (Du Bois) se possível
  if (peso > 0 && altura > 0) {
    // Altura em cm? Fórmula Du Bois usa cm e kg -> 0.007184 * W^0.425 * H^0.725
    bsa = 0.007184 * Math.pow(peso, 0.425) * Math.pow(altura, 0.725);
  }

  // 2. PEDIATRIA (<18)
  if (idadeAnos < 18) {
    if (altura && !isNaN(altura) && altura > 0) {
      resultado = (0.413 * altura) / creat; // Schwartz Bedside
      metodo = "Schwartz (Ped)";
    } else {
      return { valor: null, metodo: "Falta Altura (Ped)", estagio: "Inconclusivo" };
    }
  } 
  // 3. ADULTO (CKD-EPI 2021)
  else {
    const isF = String(sexo).toUpperCase().startsWith("F");
    const k = isF ? 0.7 : 0.9, a = isF ? -0.241 : -0.302;
    const f1 = Math.pow(Math.min(creat/k, 1), a);
    const f2 = Math.pow(Math.max(creat/k, 1), -1.200);
    const f3 = Math.pow(0.9938, idadeAnos);
    const f4 = isF ? 1.012 : 1.0;
    
    resultado = 142 * f1 * f2 * f3 * f4; // mL/min/1.73m²
    metodo = "CKD-EPI 2021";
  }

  // 4. Desnormalização (mL/min)
  if (bsa > 0) {
    clcr_absoluto = resultado * (bsa / 1.73);
  } else if (peso > 0 && idadeAnos >= 18) {
    // Backup Cockcroft-Gault
    const isF = String(sexo).toUpperCase().startsWith("F");
    let cg = ((140 - idadeAnos) * peso) / (72 * creat);
    if (isF) cg *= 0.85;
    clcr_absoluto = cg;
    metodo += " (+CG Dose)";
  } else {
    clcr_absoluto = resultado; // Proxy
  }

  return { 
    valor: parseFloat(resultado.toFixed(1)), 
    valor_absoluto: parseFloat(clcr_absoluto.toFixed(1)),
    metodo: metodo, 
    estagio: classificarEstagioRenal_(resultado),
    bsa: bsa ? parseFloat(bsa.toFixed(2)) : null
  };
}

function classificarEstagioRenal_(v) {
  if (v >= 90) return "G1 (Normal)";
  if (v >= 60) return "G2 (Leve)";
  if (v >= 45) return "G3a (Leve-Mod)";
  if (v >= 30) return "G3b (Mod-Grave)";
  if (v >= 15) return "G4 (Grave)";
  return "G5 (Falência)";
}

function interpretarPosologia_(doseTxt, aprazTxt) {
  const dStr = String(doseTxt||"0").replace(',','.').toLowerCase().trim();
  const aStr = String(aprazTxt||"").toLowerCase().trim();
  
  const dMatch = dStr.match(/(\d+(\.\d+)?)/);
  let dose = dMatch ? parseFloat(dMatch[0]) : 0;
  if(dose === 0) return 0;

  let f = 1;
  const mIntervalo = aStr.match(/(\d+)\s*(?:\/|-|em|a cada|h|:)\s*(\d+)/); // "8/8"
  const mCadaHora = aStr.match(/(?:q|cada)\s*(\d+)\s*h?/); // "q8h"

  if (mIntervalo && parseFloat(mIntervalo[2]) > 0) f = 24 / parseFloat(mIntervalo[2]);
  else if (mCadaHora && parseFloat(mCadaHora[1]) > 0) f = 24 / parseFloat(mCadaHora[1]);
  else if (/(12|bid|2x|duas|manh. e noite)/.test(aStr)) f = 2;
  else if (/(8|tid|3x|tres|três)/.test(aStr)) f = 3;
  else if (/(6|qid|4x|quatro)/.test(aStr)) f = 4;
  else if (/(4|6x|seis)/.test(aStr)) f = 6;
  else if (/(unica|única|1x|uma|od|24h)/.test(aStr)) f = 1;
  
  return parseFloat((dose * f).toFixed(2));
}

// ============================================================================
// 6. DATA FETCHERS (COM EXPLORADOR DINÂMICO)
// ============================================================================

/**
 * Função API para o Explorador. 
 * Suporta colunas virtuais: CALC_RENAL_TFG, CALC_DOSE_KG, CALC_IDADE.
 */
function apiGetRawExplorerData(sheetName, columns) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { sucesso: false, erro: "Aba não encontrada." };
  
  // Carrega Configuração para saber onde estão as colunas clínicas
  const props = PropertiesService.getScriptProperties();
  const config = JSON.parse(props.getProperty('EUM_CONFIG_MASTER') || '{}');
  const core = config.core || {};

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  
  // Índices das Colunas Essenciais (para cálculos)
  const idxCreat = core.colCreat ? headers.indexOf(core.colCreat) : -1;
  const idxPeso = core.colPeso ? headers.indexOf(core.colPeso) : -1;
  const idxAltura = core.colAltura ? headers.indexOf(core.colAltura) : -1;
  const idxDoseUni = core.colDoseUni ? headers.indexOf(core.colDoseUni) : -1;
  const idxApraz = core.colApraz ? headers.indexOf(core.colApraz) : -1;
  const idxProntFatos = headers.indexOf(core.colProntFatos);

  // Prepara mapa de pacientes (Join com Dimensão se necessário)
  let mapPacientes = {};
  if (core.abaDim && core.colProntDim) {
     const sheetDim = ss.getSheetByName(core.abaDim);
     if(sheetDim) {
       const dataDim = sheetDim.getDataRange().getValues();
       const hDim = dataDim.shift();
       const iPront = hDim.indexOf(core.colProntDim);
       const iNasc = hDim.indexOf(core.colNasc);
       const iSexo = hDim.indexOf(core.colSexo);
       dataDim.forEach(r => {
         mapPacientes[String(r[iPront]).trim()] = { 
           nasc: r[iNasc] ? new Date(r[iNasc]) : null, 
           sexo: r[iSexo] 
         };
       });
     }
  }
  
  const result = [];
  
  // Itera sobre os dados
  for (let i = 1; i < data.length; i++) {
    const rowObj = {}; 
    let hasData = false;
    const row = data[i];
    
    // Contexto do Paciente
    const pront = idxProntFatos > -1 ? String(row[idxProntFatos]).trim() : null;
    const pacData = mapPacientes[pront] || {};
    
    columns.forEach((col, k) => {
      
      // --- CÁLCULOS DINÂMICOS ---
      if (col === 'CALC_RENAL_TFG') {
        const creat = idxCreat > -1 ? parseFloat(String(row[idxCreat]).replace(',','.')) : 0;
        const peso = idxPeso > -1 ? parseFloat(String(row[idxPeso]).replace(',','.')) : 0;
        const altura = idxAltura > -1 ? parseFloat(String(row[idxAltura]).replace(',','.')) : 0;
        
        let idade = 0;
        if (pacData.nasc) idade = (new Date() - pacData.nasc) / (365.25 * 24 * 60 * 60 * 1000);
        
        // Chama o motor clínico atualizado
        const renal = calcularFuncaoRenal_(creat, idade, pacData.sexo || 'M', peso, altura);
        rowObj[col] = renal.valor; 
        hasData = true;
      } 
      else if (col === 'CALC_DOSE_KG') {
        const dose = interpretarPosologia_(
          idxDoseUni > -1 ? row[idxDoseUni] : "0", 
          idxApraz > -1 ? row[idxApraz] : ""
        );
        const peso = idxPeso > -1 ? parseFloat(String(row[idxPeso]).replace(',','.')) : 0;
        rowObj[col] = (peso > 0) ? parseFloat((dose/peso).toFixed(2)) : 0;
        hasData = true;
      }
      else if (col === 'CALC_IDADE') {
        if (pacData.nasc) {
           rowObj[col] = parseFloat(((new Date() - pacData.nasc) / (365.25 * 86400000)).toFixed(1));
           hasData = true;
        } else {
           rowObj[col] = 0;
        }
      }
      // --- COLUNAS REAIS ---
      else {
        const idx = headers.indexOf(col);
        if (idx > -1) {
          const val = row[idx];
          if (typeof val === 'string' && val.trim() !== '' && !isNaN(parseFloat(val.replace(',','.')))) { 
            rowObj[col] = parseFloat(val.replace(',','.')); 
          } else { 
            rowObj[col] = val; 
          }
          if (val !== "") hasData = true;
        }
      }
    });
    
    if (hasData) result.push(rowObj);
  }
  return { sucesso: true, dados: result };
}

function apiGetEficaciaData(params) {
  const state = apiGetInitialState();
  if (!state.status.hasExames) return { sucesso: false, erro: "Exames não configurados." };

  try {
    const cfgCore = state.config.core;
    const cfgEx = state.config.exames;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ONE_DAY = 86400000;

    // 1. Carrega Pacientes
    const dimData = ss.getSheetByName(cfgCore.abaDim).getDataRange().getValues();
    const hDim = dimData.shift();
    const idxProntD = hDim.indexOf(cfgCore.colProntDim), idxNasc = hDim.indexOf(cfgCore.colNasc), idxSexo = hDim.indexOf(cfgCore.colSexo);
    const pacMeta = {};
    dimData.forEach(r => { 
      pacMeta[String(r[idxProntD]).trim()] = { n: new Date(r[idxNasc]), s: String(r[idxSexo]||"").toUpperCase().charAt(0) }; 
    });

    // 2. Carrega Eventos e Clínicos
    const fatosData = ss.getSheetByName(cfgCore.abaFatos).getDataRange().getValues();
    const hFatos = fatosData.shift();
    const idxProntF = hFatos.indexOf(cfgCore.colProntFatos), idxMed = hFatos.indexOf(cfgCore.colMed), idxDtIni = hFatos.indexOf(cfgCore.colDtIni);
    const idxDoseUni=hFatos.indexOf(cfgCore.colDoseUni), idxApraz=hFatos.indexOf(cfgCore.colApraz), idxDose24=hFatos.indexOf(cfgCore.colDose24h);
    const idxPeso=hFatos.indexOf(cfgCore.colPeso), idxAltura=hFatos.indexOf(cfgCore.colAltura), idxCreat=hFatos.indexOf(cfgCore.colCreat);

    const diaZero = {};
    fatosData.forEach(r => {
      if (r[idxMed] === params.medicamento) {
        const p = String(r[idxProntF]).trim();
        const dt = new Date(r[idxDtIni]);
        if (!isNaN(dt.getTime())) {
          let doseDia = 0;
          if(idxDose24 > -1 && r[idxDose24]) doseDia = parseFloat(String(r[idxDose24]).replace(',','.'));
          else if(idxDoseUni > -1) doseDia = interpretarPosologia_(r[idxDoseUni], idxApraz>-1 ? r[idxApraz] : "");
          
          const peso = idxPeso > -1 ? parseFloat(String(r[idxPeso]).replace(',','.')) : 0;
          const altura = idxAltura > -1 ? parseFloat(String(r[idxAltura]).replace(',','.')) : 0;
          const creat = idxCreat > -1 ? parseFloat(String(r[idxCreat]).replace(',','.')) : 0;

          if (!diaZero[p] || dt < diaZero[p].dt) {
             diaZero[p] = { dt: dt, doseDia: doseDia, peso: peso, altura: altura, creat: creat };
          }
        }
      }
    });

    // 3. Processa Exames
    const exData = ss.getSheetByName(cfgEx.aba).getDataRange().getValues();
    const hEx = exData.shift();
    const idxP=hEx.indexOf(cfgEx.colPront), idxN=hEx.indexOf(cfgEx.colNome), idxV=hEx.indexOf(cfgEx.colValor), idxD=hEx.indexOf(cfgEx.colData);
    const idxMin=hEx.indexOf("REF_MIN"), idxMax=hEx.indexOf("REF_MAX");

    const plotData = [];
    exData.forEach(r => {
      const p = String(r[idxP]).trim();
      if (diaZero[p] && normalize_(r[idxN]) === normalize_(params.exameAlvo)) {
         const dtEx = new Date(r[idxD]);
         let val = r[idxV]; if (typeof val === 'string') val = parseFloat(val.replace(',', '.'));
         
         if (!isNaN(dtEx.getTime()) && !isNaN(val)) {
           const meta = pacMeta[p] || {};
           const dias = Math.floor((dtEx - diaZero[p].dt) / ONE_DAY);
           let idade = 0; if(meta.n) idade = parseFloat(((dtEx - meta.n)/(365.25*ONE_DAY)).toFixed(1));
           
           const dc = diaZero[p];
           // Usa creatinina atual se o exame for creatinina
           let creatAtual = normalize_(r[idxN]).includes("creat") ? val : dc.creat;

           const renal = calcularFuncaoRenal_(creatAtual, idade, meta.s, dc.peso, dc.altura);
           const doseKg = (dc.doseDia && dc.peso) ? (dc.doseDia / dc.peso) : 0;
           
           plotData.push({ 
             x: dias, y: val, prontuario: p, 
             refMin: (idxMin>-1&&r[idxMin]!=="")?parseFloat(r[idxMin]):null, 
             refMax: (idxMax>-1&&r[idxMax]!=="")?parseFloat(r[idxMax]):null,
             sexo: meta.s, idade: idade, 
             doseTotal: dc.doseDia, doseKg: parseFloat(doseKg.toFixed(2)), 
             renal: renal 
           });
         }
      }
    });
    return { sucesso: true, dados: plotData };
  } catch (e) { return { sucesso: false, erro: e.message }; }
}

function apiGetSegurancaData(params) {
  const state = apiGetInitialState();
  if (!state.status.hasRams) return { sucesso: false, erro: "RAMs não configuradas." };
  try {
    const cfgC = state.config.core;
    const cfgR = state.config.rams;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const fatosData = ss.getSheetByName(cfgC.abaFatos).getDataRange().getValues();
    const hF = fatosData.shift();
    const idxPF = hF.indexOf(cfgC.colProntFatos), idxMed = hF.indexOf(cfgC.colMed), idxDI = hF.indexOf(cfgC.colDtIni);
    const coorte = {};
    fatosData.forEach(r => { if(r[idxMed]===params.medicamento) { const p=String(r[idxPF]).trim(); const d=new Date(r[idxDI]); if(!coorte[p] || d<coorte[p]) coorte[p]=d; }});
    
    const ramData = ss.getSheetByName(cfgR.aba).getDataRange().getValues();
    const hR = ramData.shift();
    const idxG=hR.indexOf(cfgR.colGrav), idxC=hR.indexOf(cfgR.colCaus);
    const idxDesc = cfgR.colDesc ? hR.indexOf(cfgR.colDesc) : hR.findIndex(c=>c.toUpperCase().includes("DESC")||c.toUpperCase().includes("RAM")||c.toUpperCase().includes("TIPO"));
    const idxDtR = cfgR.colData ? hR.indexOf(cfgR.colData) : hR.findIndex(c=>c.toUpperCase().includes("DATA"));
    const idxPR = hR.findIndex(c=>c.toUpperCase().includes("PRONT"));
    const events = [];
    ramData.forEach(r => {
      const p = String(r[idxPR]).trim();
      if(coorte[p]) {
        const dt = idxDtR>-1 ? new Date(r[idxDtR]) : null;
        const dias = (dt && !isNaN(dt)) ? Math.floor((dt - coorte[p])/86400000) : null;
        events.push({ gravidade: String(r[idxG]||"N/D").trim(), causalidade: String(r[idxC]||"N/D").trim(), descricao: idxDesc>-1 ? String(r[idxDesc]).trim() : "RAM", dias: dias, prontuario: p });
      }
    });
    return { sucesso: true, dados: events, totalExpostos: Object.keys(coorte).length };
  } catch(e) { return { sucesso: false, erro: e.message }; }
}

// ============================================================================
// 7. UTILITÁRIOS E BANCO DE REGRAS
// ============================================================================

function apiGetMedicamentos(sheetName, colName) { return getUniqueValues_(sheetName, colName); }
function apiGetExamesList(sheetName, colName) { return getUniqueValues_(sheetName, colName); }

function apiGetReferenciasDb() { 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Exames_Referencia");
  if (!sheet) { sheet = ss.insertSheet("Exames_Referencia"); sheet.appendRow(["NOME_EXAME", "TIPO_REGRA", "SEXO", "IDADE_MIN_DIAS", "IDADE_MAX_DIAS", "VALOR_MIN", "VALOR_MAX"]); }
  const data = sheet.getDataRange().getValues(); if(data.length<2) return { sucesso: true, dados: [] };
  const h = data.shift().map(c => String(c).toUpperCase().trim());
  const idxNome=h.findIndex(c=>c.includes("NOME")), idxTipo=h.findIndex(c=>c.includes("TIPO")), idxSex=h.indexOf("SEXO"), idxDmin=h.findIndex(c=>c.includes("MIN_DIAS")), idxDmax=h.findIndex(c=>c.includes("MAX_DIAS")), idxVmin=h.findIndex(c=>c.includes("VALOR_MIN")), idxVmax=h.findIndex(c=>c.includes("VALOR_MAX"));
  const out = [];
  data.forEach((row, i) => { if(row[idxNome]) { out.push({ id: i + 2, nome: row[idxNome], tipo: idxTipo > -1 ? (row[idxTipo] || "Padrão") : "Padrão", sexo: row[idxSex] || "A", diasMin: row[idxDmin], diasMax: row[idxDmax], min: row[idxVmin], max: row[idxVmax] }); } });
  return { sucesso: true, dados: out };
}

function apiSalvarReferencia(regra) { 
  const ss = SpreadsheetApp.getActiveSpreadsheet(); let sheet = ss.getSheetByName("Exames_Referencia");
  if(!sheet) return { sucesso: false, erro: "Aba não encontrada." };
  const h = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(c=>String(c).toUpperCase().trim());
  const getIdx = (k) => { let i = h.findIndex(c => c.includes(k)); if(i === -1) { i = sheet.getLastColumn(); sheet.getRange(1, i+1).setValue(k); h.push(k); } return i; };
  const rowData = []; for(let k=0; k<h.length; k++) rowData.push("");
  rowData[getIdx("NOME")] = regra.nome.toUpperCase(); rowData[getIdx("TIPO")] = regra.tipo; rowData[getIdx("SEXO")] = regra.sexo; rowData[getIdx("IDADE_MIN")] = regra.diasMin; rowData[getIdx("IDADE_MAX")] = regra.diasMax; rowData[getIdx("VALOR_MIN")] = regra.min; rowData[getIdx("VALOR_MAX")] = regra.max;
  sheet.appendRow(rowData);
  return { sucesso: true, msg: "Salvo!" };
}

function apiExcluirReferencia(id) { try { SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Exames_Referencia").deleteRow(parseInt(id)); return { sucesso: true }; } catch(e) { return { sucesso: false, erro: e.message }; } }

function apiGetRegrasDose() { 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Doses_Referencia");
  if (!sheet) { sheet = ss.insertSheet("Doses_Referencia"); sheet.appendRow(["MEDICAMENTO", "IDADE_MIN_DIAS", "IDADE_MAX_DIAS", "DOSE_MIN", "DOSE_USUAL", "DOSE_MAX", "UNIDADE"]); }
  const data = sheet.getDataRange().getValues(); if(data.length < 2) return { sucesso: true, dados: [] };
  const h = data.shift().map(c => String(c).toUpperCase().trim());
  const idxMed = h.findIndex(c => c.includes("MEDICAMENTO")), idxDmin = h.findIndex(c => c.includes("IDADE_MIN")), idxDmax = h.findIndex(c => c.includes("IDADE_MAX")), idxVmin = h.findIndex(c => c === "DOSE_MIN"), idxVus = h.findIndex(c => c.includes("USUAL")), idxVmax = h.findIndex(c => c === "DOSE_MAX"), idxUn = h.findIndex(c => c.includes("UNIDADE"));
  const out = [];
  data.forEach((row, i) => { if(row[idxMed]) { out.push({ id: i + 2, med: row[idxMed], diasMin: row[idxDmin]||0, diasMax: row[idxDmax]||36500, doseMin: row[idxVmin], doseUsual: row[idxVus], doseMax: row[idxVmax], unidade: row[idxUn]||"mg/dia" }); } });
  return { sucesso: true, dados: out };
}

function apiSalvarRegraDose(regra) { 
  const ss = SpreadsheetApp.getActiveSpreadsheet(); let sheet = ss.getSheetByName("Doses_Referencia");
  if(!sheet) return { sucesso: false, erro: "Aba não encontrada." };
  const h = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(c=>String(c).toUpperCase().trim());
  const getIdx = (k) => { let i = h.indexOf(k); if(i === -1) { i = sheet.getLastColumn(); sheet.getRange(1, i+1).setValue(k); h.push(k); } return i; };
  const rowData = []; for(let k=0; k<h.length; k++) rowData.push("");
  rowData[getIdx("MEDICAMENTO")] = regra.med.toUpperCase(); rowData[getIdx("IDADE_MIN_DIAS")] = regra.diasMin; rowData[getIdx("IDADE_MAX_DIAS")] = regra.diasMax; rowData[getIdx("DOSE_MIN")] = regra.doseMin; rowData[getIdx("DOSE_USUAL")] = regra.doseUsual; rowData[getIdx("DOSE_MAX")] = regra.doseMax; rowData[getIdx("UNIDADE")] = regra.unidade;
  sheet.appendRow(rowData);
  return { sucesso: true, msg: "Salvo!" };
}

function apiExcluirRegraDose(id) { try { SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Doses_Referencia").deleteRow(parseInt(id)); return { sucesso: true }; } catch(e) { return { sucesso: false, erro: e.message }; } }

function apiExportarParaMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetLocal = ss.getSheetByName("Exames_Referencia");
  if(!sheetLocal) return { sucesso: false, erro: "Não há dados." };
  const dadosLocais = sheetLocal.getDataRange().getValues();
  try {
    const masterSS = SpreadsheetApp.openById(MASTER_DB_ID);
    let sheetMaster = masterSS.getSheetByName("Exames_Referencia");
    if(!sheetMaster) sheetMaster = masterSS.insertSheet("Exames_Referencia");
    sheetMaster.clear();
    sheetMaster.getRange(1, 1, dadosLocais.length, dadosLocais[0].length).setValues(dadosLocais);
    return { sucesso: true, msg: "Sincronizado!" };
  } catch(e) { return { sucesso: false, erro: e.message }; }
}

function apiGetColumns(sheetName) {
  if (!sheetName) return [];
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  return sheet ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].filter(h => h) : [];
}

function getUniqueValues_(sheetName, colName) { 
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if(!sheet) return [];
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = headers.indexOf(colName);
  if(colIndex === -1) return [];
  return [...new Set(sheet.getRange(2, colIndex+1, sheet.getLastRow()-1, 1).getValues().map(v=>String(v[0]).trim()))].filter(v=>v).sort();
}

function apiImportarReferencias() { 
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  try { 
    const master = SpreadsheetApp.openById(MASTER_DB_ID);
    const data = (master.getSheetByName("Exames_Referencia")||master.getSheets()[0]).getDataRange().getValues(); 
    let sheet = ss.getSheetByName("Exames_Referencia") || ss.insertSheet("Exames_Referencia"); 
    sheet.clear().getRange(1,1,data.length,data[0].length).setValues(data); 
    return { sucesso: true, msg: "Importado!" };
  } catch(e) { return { sucesso: false, erro: e.message }; } 
}

function apiProcessarAnaliseExames() { return {sucesso:true, msg:"Use a aba Eficácia."}; }
function normalize_(s) { return String(s||"").toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, ""); }
function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }
