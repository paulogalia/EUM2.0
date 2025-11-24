/**
 * ============================================================================
 * EUM MANAGER 3.0 - CÉREBRO COMPLETO (VERSÃO PRODUÇÃO)
 * * Módulos Integrados:
 * 1. Wizard Mágico (Cria Abas + Analisa PDF) -> apiMagicSetup
 * 2. Explorador de Dados (Text-to-Chart) -> apiInterpretarGraficoIA
 * 3. Motores Clínicos Determinísticos (Renal CKD-EPI/Schwartz + Posologia)
 * 4. Gestão de Regras e Sincronização
 * 5. Processamento de Dados de Eficácia e Segurança
 * ============================================================================
 */

// --- CONSTANTES GLOBAIS ---
const MASTER_DB_ID = "1zKqYVR9seTPy3eyX5CR2xuXZtqSwwvERQn_KX3GK5JM"; 
const GEMINI_API_KEY = "AIzaSyAVzqzfled_SdpmAXCUSH5kg-LwmkFwFvM"; 

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

// ============================================================================
// 2. API DE CONFIGURAÇÃO E ESTADO
// ============================================================================

function apiGetInitialState() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().map(function(s) { return s.getName(); });
  const props = PropertiesService.getScriptProperties();
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
// 3. WIZARD MÁGICO (UPLOAD + IA + CRIAÇÃO DE ABA)
// ============================================================================

/**
 * Recebe PDF (Base64) e DADOS (Matriz).
 * 1. Cria uma nova aba com os dados.
 * 2. Analisa o PDF e os cabeçalhos com IA.
 * 3. Retorna configuração sugerida.
 */
function apiMagicSetup(pdfBase64, matrixDados, nomeArquivo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. CRIAR ABA E IMPORTAR DADOS
  let sheetName = "Dados_" + (nomeArquivo || "Import").replace(/[^a-zA-Z0-9]/g, "_").substring(0, 15);
  // Garante nome único para não dar erro
  if (ss.getSheetByName(sheetName)) sheetName += "_" + Math.floor(Math.random()*1000);
  
  const sheet = ss.insertSheet(sheetName);
  
  // Cola os dados (Matriz) de uma vez
  if (matrixDados && matrixDados.length > 0) {
    sheet.getRange(1, 1, matrixDados.length, matrixDados[0].length).setValues(matrixDados);
  } else {
    return { erro: "Arquivo de dados vazio." };
  }

  // Pega os cabeçalhos reais da nova aba para a IA analisar
  const csvHeaders = matrixDados[0];

  // 2. ANÁLISE DE IA
  if (!GEMINI_API_KEY || GEMINI_API_KEY.includes("AQUI")) return { erro: "Chave API inválida." };
  
  const prompt = `
    Atue como Consultor Sênior EUM Manager.
    
    INPUTS:
    1. Cabeçalhos dos Dados (Acabei de importar na aba '${sheetName}'): ${JSON.stringify(csvHeaders)}
    2. Protocolo do Estudo (PDF anexo).

    TAREFA:
    1. Analise a compatibilidade.
    2. Liste "Capacidades" (O que dá para fazer?) e "Lacunas" (O que falta?).
    3. Mapeie as colunas do CSV para os campos internos:
       - colProntFatos, colMed, colDtIni
       - colDoseUni, colApraz, colDose24h
       - colPeso, colAltura, colCreat
       - colProntDim, colNasc, colSexo

    RETORNE JSON:
    {
      "studyName": "...",
      "studySummary": "...",
      "preset": "safety" | "efficacy" | "protocol" | "full",
      "analysis": { "capabilities": [], "gaps": [] },
      "mapping": { "colProntFatos": "NOME_COLUNA", ... }
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

  // Prioridade: 2.0 Flash (Mais rápido e capaz)
  const modelos = ["gemini-2.0-flash", "gemini-1.5-flash"];

  for (let i = 0; i < modelos.length; i++) {
    let modelo = modelos[i];
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelo}:generateContent?key=${GEMINI_API_KEY}`;
      const options = {
        method: 'post', contentType: 'application/json',
        payload: JSON.stringify(payload), muteHttpExceptions: true
      };
      
      const resp = UrlFetchApp.fetch(url, options);
      const content = JSON.parse(resp.getContentText());
      
      if (content.error) throw new Error(content.error.message);
      
      const txt = content.candidates[0].content.parts[0].text;
      const resultadoIA = JSON.parse(txt.replace(/```json/g, "").replace(/```/g, "").trim());
      
      // Adiciona o nome da aba criada ao resultado para o Frontend saber qual selecionar
      resultadoIA.createdSheet = sheetName;
      
      return resultadoIA;

    } catch (e) { console.log(`Erro IA (${modelo}): ` + e.message); }
  }
  
  // Retorno de emergência se a IA falhar, mas a aba foi criada
  return { erro: "Falha na análise IA, mas a aba '" + sheetName + "' foi criada." };
}

// --- 3.1. FUNÇÃO DE ANÁLISE IA DIRETA (SEM UPLOAD) ---
function apiAnalisarArquivosIA(pdfBase64, csvHeaders) {
  // Esta função serve para o caso de só analisar sem criar aba (ex: re-análise)
  // Reutiliza a mesma lógica interna, mas sem sheet.insertSheet
  if (!GEMINI_API_KEY || GEMINI_API_KEY.includes("AQUI")) return { erro: "Chave API inválida." };
  
  const prompt = `Atue como Consultor Sênior EUM Manager. Inputs: ${JSON.stringify(csvHeaders)} e PDF. Mapeie colunas para colProntFatos, colMed, etc. Retorne JSON: {studyName, studySummary, preset, analysis, mapping}`;
  
  const payload = {
    contents: [{ parts: [ { text: prompt }, { inline_data: { mime_type: "application/pdf", data: pdfBase64 } } ] }]
  };

  const modelos = ["gemini-2.0-flash", "gemini-1.5-flash"];
  for (let i = 0; i < modelos.length; i++) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelos[i]}:generateContent?key=${GEMINI_API_KEY}`;
      const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };
      const resp = UrlFetchApp.fetch(url, options);
      const content = JSON.parse(resp.getContentText());
      if (content.error) throw new Error(content.error.message);
      const txt = content.candidates[0].content.parts[0].text;
      return JSON.parse(txt.replace(/```json/g, "").replace(/```/g, "").trim());
    } catch (e) { console.log("Erro IA Direta: " + e.message); }
  }
  return { erro: "Falha na análise." };
}


// --- 4. WIZARD TEXTO (FALLBACK) ---

function apiInterpretarIntencaoEstudo(texto) {
  if (!texto) return { erro: "Texto vazio" };
  const prompt = `Consultor EUM. Usuário: "${texto}". Retorne JSON: {"preset": "...", "nome_sugerido": "...", "resumo": "...", "explicacao": "..."}`;
  return chamarGeminiSimples_(prompt);
}

// --- 5. EXPLORADOR DE GRÁFICOS (TEXT-TO-CHART) ---

function apiInterpretarGraficoIA(pergunta, colunasDisponiveis) {
  if (!pergunta) return { erro: "Pergunta vazia" };
  
  const prompt = `
    Analista de Dados (Plotly).
    CONTEXTO: Colunas: ${JSON.stringify(colunasDisponiveis)}. Pergunta: "${pergunta}".
    TAREFA:
    1. Identifique Eixo X e Y. Se pedir contagem, Y é null.
    2. Escolha gráfico: bar, line, pie, scatter, box.
    3. Agregação: count, sum, avg.
    
    RETORNE JSON:
    { "viavel": true, "config": { "x": "...", "y": "...", "type": "bar", "aggregation": "count" }, "explicacao": "..." }
  `;

  return chamarGeminiSimples_(prompt);
}

function chamarGeminiSimples_(prompt) {
  const modelos = ["gemini-2.0-flash", "gemini-1.5-flash"];
  for (let i = 0; i < modelos.length; i++) {
    let m = modelos[i];
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${m}:generateContent?key=${GEMINI_API_KEY}`;
      const opts = { method: 'post', contentType: 'application/json', payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }), muteHttpExceptions: true };
      const res = UrlFetchApp.fetch(url, opts);
      const json = JSON.parse(res.getContentText());
      if (json.error) throw new Error(json.error.message);
      return JSON.parse(json.candidates[0].content.parts[0].text.replace(/```json/g, "").replace(/```/g, "").trim());
    } catch (e) { console.log(e); }
  }
  return { erro: "Erro IA Texto." };
}

// ============================================================================
// 6. MOTORES CLÍNICOS DETERMINÍSTICOS
// ============================================================================

/**
 * Motor Renal Híbrido: Schwartz (Ped) + CKD-EPI (Adulto) + Cockcroft (Backup)
 */
function calcularFuncaoRenal_(creat, idadeAnos, sexo, peso, altura) {
  if (!creat || isNaN(creat) || creat <= 0) {
    return { valor: null, metodo: "Sem Creatinina", estagio: "N/D", clcr_dose: null };
  }
  
  let resultado = 0;
  let metodo = "";
  let clcr_alt = null;

  // PEDIATRIA
  if (idadeAnos < 18) {
    if (altura && !isNaN(altura) && altura > 0) {
      // Fórmula Schwartz: 0.413 * Altura (cm) / Creatinina
      resultado = (0.413 * altura) / creat;
      metodo = "Schwartz (Ped)";
    } else {
      return { valor: null, metodo: "Falta Altura", estagio: "Verificar Ref" };
    }
  } 
  // ADULTO
  else {
    const isF = String(sexo).toUpperCase().startsWith("F");
    const k = isF ? 0.7 : 0.9, a = isF ? -0.241 : -0.302;
    const f1 = Math.pow(Math.min(creat/k, 1), a);
    const f2 = Math.pow(Math.max(creat/k, 1), -1.200);
    const f3 = Math.pow(0.9938, idadeAnos);
    const f4 = isF ? 1.012 : 1.0;
    
    resultado = 142 * f1 * f2 * f3 * f4;
    metodo = "CKD-EPI";

    // Backup Cockcroft para dose
    if (peso && peso > 0) {
      let cg = ((140 - idadeAnos) * peso) / (72 * creat);
      if (isF) cg *= 0.85;
      clcr_alt = parseFloat(cg.toFixed(1));
    }
  }

  return { 
    valor: parseFloat(resultado.toFixed(1)), 
    clcr_dose: clcr_alt,
    metodo: metodo, 
    estagio: classificarEstagioRenal_(resultado) 
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
  const d = parseFloat(String(doseTxt||"0").replace(',','.').match(/(\d+(\.\d+)?)/)?.[0] || 0);
  if(d === 0) return 0;
  const a = String(aprazTxt||"").toLowerCase();
  let f = 1;
  const m = a.match(/(\d+)\s*[\/\-]\s*(\d+)/);
  if (m && m[2]>0) f = 24/m[2];
  else if (a.includes('12')||a.includes('bid')) f=2;
  else if (a.includes('8')||a.includes('tid')) f=3;
  else if (a.includes('6')||a.includes('qid')) f=4;
  else if (a.includes('4')) f=6;
  return d * f;
}

// ============================================================================
// 7. DATA FETCHERS (PROCESSAMENTO DE DADOS)
// ============================================================================

function apiGetEficaciaData(params) {
  const state = apiGetInitialState();
  if (!state.status.hasExames) return { sucesso: false, erro: "Exames não configurados." };

  try {
    const cfgCore = state.config.core;
    const cfgEx = state.config.exames;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ONE_DAY = 24*60*60*1000;
    
    // 1. Carrega Dimensões (Pacientes)
    const dimData = ss.getSheetByName(cfgCore.abaDim).getDataRange().getValues();
    const hDim = dimData.shift();
    const idxProntD = hDim.indexOf(cfgCore.colProntDim), idxNasc = hDim.indexOf(cfgCore.colNasc), idxSexo = hDim.indexOf(cfgCore.colSexo);
    
    const pacMeta = {};
    dimData.forEach(function(r) { 
      pacMeta[String(r[idxProntD]).trim()] = { 
        n: new Date(r[idxNasc]), 
        s: String(r[idxSexo]||"").toUpperCase().charAt(0)
      }; 
    });

    // 2. Carrega Fatos (Medicamento, Dose, Peso, Altura, Creatinina)
    const fatosData = ss.getSheetByName(cfgCore.abaFatos).getDataRange().getValues();
    const hFatos = fatosData.shift();
    
    const idxProntF = hFatos.indexOf(cfgCore.colProntFatos);
    const idxMed = hFatos.indexOf(cfgCore.colMed);
    const idxDtIni = hFatos.indexOf(cfgCore.colDtIni);
    
    // Índices de Posologia e Clínicos
    const idxDoseUni = hFatos.indexOf(cfgCore.colDoseUni);
    const idxApraz = hFatos.indexOf(cfgCore.colApraz);
    const idxDose24 = hFatos.indexOf(cfgCore.colDose24h);
    const idxPeso = hFatos.indexOf(cfgCore.colPeso); 
    const idxAltura = hFatos.indexOf(cfgCore.colAltura); 
    const idxCreat = hFatos.indexOf(cfgCore.colCreat);

    const diaZero = {};
    
    // Processa cada linha de prescrição
    fatosData.forEach(function(r) {
      if (r[idxMed] === params.medicamento) {
        const p = String(r[idxProntF]).trim();
        const dt = new Date(r[idxDtIni]);
        if (!isNaN(dt.getTime())) {
          // Lógica Híbrida de Dose
          let doseDia = 0;
          if(idxDose24 > -1 && r[idxDose24]) {
            doseDia = parseFloat(String(r[idxDose24]).replace(',','.'));
          } else if(idxDoseUni > -1) {
            doseDia = interpretarPosologia_(r[idxDoseUni], idxApraz>-1 ? r[idxApraz] : "");
          }
          
          // Dados Clínicos
          const peso = idxPeso > -1 ? parseFloat(String(r[idxPeso]).replace(',','.')) : 0;
          const altura = idxAltura > -1 ? parseFloat(String(r[idxAltura]).replace(',','.')) : 0;
          const creat = idxCreat > -1 ? parseFloat(String(r[idxCreat]).replace(',','.')) : 0;

          if (!diaZero[p] || dt < diaZero[p].dt) {
             diaZero[p] = { dt: dt, doseDia: doseDia, peso: peso, altura: altura, creat: creat };
          }
        }
      }
    });

    // 3. Carrega Exames e Cruza com Dados
    const exData = ss.getSheetByName(cfgEx.aba).getDataRange().getValues();
    const hEx = exData.shift();
    const idxP = hEx.indexOf(cfgEx.colPront);
    const idxN = hEx.indexOf(cfgEx.colNome);
    const idxV = hEx.indexOf(cfgEx.colValor);
    const idxD = hEx.indexOf(cfgEx.colData);
    const idxMin = hEx.indexOf("REF_MIN");
    const idxMax = hEx.indexOf("REF_MAX");

    const plotData = [];
    exData.forEach(function(r) {
      const p = String(r[idxP]).trim();
      // Verifica se é o exame alvo
      if (diaZero[p] && normalize_(r[idxN]) === normalize_(params.exameAlvo)) {
         const dtEx = new Date(r[idxD]);
         let val = r[idxV]; 
         if (typeof val === 'string') val = parseFloat(val.replace(',', '.'));
         
         if (!isNaN(dtEx.getTime()) && !isNaN(val)) {
           const meta = pacMeta[p] || {};
           const dias = Math.floor((dtEx - diaZero[p].dt) / ONE_DAY);
           let idade = 0; 
           if(meta.n) idade = parseFloat(((dtEx - meta.n)/(365.25*ONE_DAY)).toFixed(1));
           
           const dadosClinicos = diaZero[p];
           
           // Usa creatinina do evento ou do exame atual se for creatinina
           let creatParaCalculo = dadosClinicos.creat;
           if(normalize_(r[idxN]).includes("creat")) creatParaCalculo = val;

           const renal = calcularFuncaoRenal_(
             creatParaCalculo, idade, meta.s, dadosClinicos.peso, dadosClinicos.altura
           );

           const doseKg = (dadosClinicos.doseDia && dadosClinicos.peso) ? (dadosClinicos.doseDia / dadosClinicos.peso) : 0;

           plotData.push({ 
             x: dias, y: val, prontuario: p, 
             refMin: (idxMin>-1&&r[idxMin]!=="")?parseFloat(r[idxMin]):null, 
             refMax: (idxMax>-1&&r[idxMax]!=="")?parseFloat(r[idxMax]):null,
             sexo: meta.s, idade: idade, 
             doseTotal: dadosClinicos.doseDia, 
             doseKg: parseFloat(doseKg.toFixed(2)),
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
    const cfgC = state.config.core; const cfgR = state.config.rams;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const fatosData = ss.getSheetByName(cfgC.abaFatos).getDataRange().getValues();
    const hF = fatosData.shift();
    const idxPF = hF.indexOf(cfgC.colProntFatos), idxMed = hF.indexOf(cfgC.colMed), idxDI = hF.indexOf(cfgC.colDtIni);
    const coorte = {};
    fatosData.forEach(function(r) { if(r[idxMed]===params.medicamento) { const p=String(r[idxPF]).trim(); const d=new Date(r[idxDI]); if(!coorte[p] || d<coorte[p]) coorte[p]=d; }});
    
    const ramData = ss.getSheetByName(cfgR.aba).getDataRange().getValues();
    const hR = ramData.shift();
    const idxG=hR.indexOf(cfgR.colGrav), idxC=hR.indexOf(cfgR.colCaus);
    const idxDesc = cfgR.colDesc ? hR.indexOf(cfgR.colDesc) : hR.findIndex(c=>c.toUpperCase().includes("DESC")||c.toUpperCase().includes("RAM")||c.toUpperCase().includes("TIPO"));
    const idxDtR = cfgR.colData ? hR.indexOf(cfgR.colData) : hR.findIndex(c=>c.toUpperCase().includes("DATA"));
    const idxPR = hR.findIndex(c=>c.toUpperCase().includes("PRONT"));
    const events = [];
    ramData.forEach(function(r) {
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

function apiGetPerfilData(p) {
  const state = apiGetInitialState();
  if (!state.status.hasCore) return { sucesso: false, erro: "Core?" };
  const c=state.config.core; const ss=SpreadsheetApp.getActiveSpreadsheet();
  const fd=ss.getSheetByName(c.abaFatos).getDataRange().getValues(); const dd=ss.getSheetByName(c.abaDim).getDataRange().getValues();
  const hf=fd.shift();
  const im=hf.indexOf(c.colMed); const ipf=hf.indexOf(c.colProntFatos);
  const coorte=new Set(); fd.forEach(r=>{ if(r[im]===p.medicamento) coorte.add(String(r[ipf]).trim()); });
  const hd=dd.shift(); const ipd=hd.indexOf(c.colProntDim); const is=hd.indexOf(c.colSexo); const ina=hd.indexOf(c.colNasc);
  const stats={sexo:{}, idade:[]}; const today=new Date();
  dd.forEach(r=>{ if(coorte.has(String(r[ipd]).trim())) {
    const s=r[is]||"N/D"; stats.sexo[s]=(stats.sexo[s]||0)+1;
    const n=new Date(r[ina]); if(!isNaN(n)) stats.idade.push((today-n)/(365.25*86400000));
  }});
  return { sucesso: true, dados: stats };
}

// ============================================================================
// 8. GESTÃO DE REGRAS (EXAMES)
// ============================================================================

function apiGetReferenciasDb() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Exames_Referencia");
  if (!sheet) { sheet = ss.insertSheet("Exames_Referencia"); sheet.appendRow(["NOME_EXAME", "TIPO_REGRA", "SEXO", "IDADE_MIN_DIAS", "IDADE_MAX_DIAS", "VALOR_MIN", "VALOR_MAX"]); }
  const data = sheet.getDataRange().getValues(); if(data.length<2) return { sucesso: true, dados: [] };
  const h = data.shift().map(c => String(c).toUpperCase().trim());
  
  const idxNome=h.findIndex(c=>c.includes("NOME")), idxTipo=h.findIndex(c=>c.includes("TIPO"));
  const idxSex=h.indexOf("SEXO"), idxDmin=h.findIndex(c=>c.includes("MIN_DIAS")), idxDmax=h.findIndex(c=>c.includes("MAX_DIAS"));
  const idxVmin=h.findIndex(c=>c.includes("VALOR_MIN")), idxVmax=h.findIndex(c=>c.includes("VALOR_MAX"));
  
  const out = [];
  data.forEach(function(row, i) {
    if(row[idxNome]) {
       out.push({
         id: i + 2, nome: row[idxNome], tipo: idxTipo > -1 ? (row[idxTipo] || "Padrão") : "Padrão",
         sexo: row[idxSex] || "A", diasMin: row[idxDmin], diasMax: row[idxDmax], min: row[idxVmin], max: row[idxVmax]
       });
    }
  });
  return { sucesso: true, dados: out };
}

function apiSalvarReferencia(regra) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Exames_Referencia");
  if(!sheet) return { sucesso: false, erro: "Aba não encontrada." };
  try {
    const h = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(c=>String(c).toUpperCase().trim());
    const getIdx = (k) => { let i = h.findIndex(c => c.includes(k)); if(i === -1) { i = sheet.getLastColumn(); sheet.getRange(1, i+1).setValue(k); h.push(k); } return i; };
    const rowData = []; for(let k=0; k<h.length; k++) rowData.push("");
    rowData[getIdx("NOME")] = regra.nome.toUpperCase(); rowData[getIdx("TIPO")] = regra.tipo;
    rowData[getIdx("SEXO")] = regra.sexo; rowData[getIdx("IDADE_MIN")] = regra.diasMin;
    rowData[getIdx("IDADE_MAX")] = regra.diasMax; rowData[getIdx("VALOR_MIN")] = regra.min; rowData[getIdx("VALOR_MAX")] = regra.max;
    sheet.appendRow(rowData);
    return { sucesso: true, msg: "Salvo!" };
  } catch(e) { return { sucesso: false, erro: e.message }; }
}

function apiExcluirReferencia(id) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Exames_Referencia").deleteRow(parseInt(id));
    return { sucesso: true };
  } catch(e) { return { sucesso: false, erro: e.message }; }
}

// ============================================================================
// 9. GESTÃO DE REGRAS (DOSES)
// ============================================================================

function apiGetRegrasDose() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Doses_Referencia");
  if (!sheet) { sheet = ss.insertSheet("Doses_Referencia"); sheet.appendRow(["MEDICAMENTO", "IDADE_MIN_DIAS", "IDADE_MAX_DIAS", "DOSE_MIN", "DOSE_USUAL", "DOSE_MAX", "UNIDADE"]); }
  const data = sheet.getDataRange().getValues(); if(data.length < 2) return { sucesso: true, dados: [] };
  const h = data.shift().map(c => String(c).toUpperCase().trim());
  const idxMed = h.findIndex(c => c.includes("MEDICAMENTO"));
  const idxDmin = h.findIndex(c => c.includes("IDADE_MIN")), idxDmax = h.findIndex(c => c.includes("IDADE_MAX"));
  const idxVmin = h.findIndex(c => c === "DOSE_MIN"), idxVus = h.findIndex(c => c.includes("USUAL")), idxVmax = h.findIndex(c => c === "DOSE_MAX");
  const idxUn = h.findIndex(c => c.includes("UNIDADE"));
  const out = [];
  data.forEach(function(row, i) {
    if(row[idxMed]) {
       out.push({
         id: i + 2, med: row[idxMed], diasMin: row[idxDmin]||0, diasMax: row[idxDmax]||36500, doseMin: row[idxVmin], doseUsual: row[idxVus], doseMax: row[idxVmax], unidade: row[idxUn]||"mg/dia"
       });
    }
  });
  return { sucesso: true, dados: out };
}

function apiSalvarRegraDose(regra) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); let sheet = ss.getSheetByName("Doses_Referencia"); if(!sheet) return { sucesso: false, erro: "Aba não encontrada." };
  const h = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(c=>String(c).toUpperCase().trim());
  const getIdx = (k) => { let i = h.indexOf(k); if(i === -1) { i = sheet.getLastColumn(); sheet.getRange(1, i+1).setValue(k); h.push(k); } return i; };
  const rowData = []; for(let k=0; k<h.length; k++) rowData.push("");
  rowData[getIdx("MEDICAMENTO")] = regra.med.toUpperCase(); rowData[getIdx("IDADE_MIN_DIAS")] = regra.diasMin; rowData[getIdx("IDADE_MAX_DIAS")] = regra.diasMax;
  rowData[getIdx("DOSE_MIN")] = regra.doseMin; rowData[getIdx("DOSE_USUAL")] = regra.doseUsual; rowData[getIdx("DOSE_MAX")] = regra.doseMax; rowData[getIdx("UNIDADE")] = regra.unidade;
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
    sheetMaster.clear(); sheetMaster.getRange(1, 1, dadosLocais.length, dadosLocais[0].length).setValues(dadosLocais);
    return { sucesso: true, msg: "Sincronizado!" };
  } catch(e) { return { sucesso: false, erro: e.message }; }
}

// ============================================================================
// 10. UTILITÁRIOS GERAIS
// ============================================================================

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

function apiGetRawExplorerData(sheetName, columns) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { sucesso: false, erro: "Aba não encontrada." };
  const data = sheet.getDataRange().getValues(); const headers = data[0].map(h => String(h).trim());
  const indices = columns.map(col => headers.indexOf(col));
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const rowObj = {}; let hasData = false;
    columns.forEach((col, k) => {
      const val = data[i][indices[k]];
      if (typeof val === 'string' && val.trim() !== '' && !isNaN(parseFloat(val.replace(',','.')))) { rowObj[col] = parseFloat(val.replace(',','.')); } 
      else { rowObj[col] = val; }
      if (val !== "") hasData = true;
    });
    if (hasData) result.push(rowObj);
  }
  return { sucesso: true, dados: result };
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

function apiProcessarAnaliseExames() { return {sucesso:true, msg:"Módulo Legado"}; }
function normalize_(s) { return String(s||"").toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, ""); }
function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }
