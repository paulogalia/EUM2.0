/**
 * ============================================================================
 * EUM - SISTEMA DE INTELIGÊNCIA CLÍNICA (BACKEND FINAL)
 * ============================================================================
 */

const MASTER_DB_ID = "1zKqYVR9seTPy3eyX5CR2xuXZtqSwwvERQn_KX3GK5JM";

// --- 1. INICIALIZAÇÃO ---
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

// --- 2. API & CONFIGURAÇÃO ---
function apiGetInitialState() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets().map(s => s.getName());
  const props = PropertiesService.getScriptProperties();
  const savedConfig = JSON.parse(props.getProperty('EUM_CONFIG_MASTER') || '{}');
  
  const status = {
    hasCore: !!(savedConfig?.core?.abaFatos),
    hasExames: !!(savedConfig?.exames?.aba),
    hasRams: !!(savedConfig?.rams?.aba)
  };
  return { sheets, config: savedConfig, status };
}

function apiSaveConfig(newConfig) {
  try {
    PropertiesService.getScriptProperties().setProperty('EUM_CONFIG_MASTER', JSON.stringify(newConfig));
    return { sucesso: true, status: {
      hasCore: !!(newConfig.core.abaFatos),
      hasExames: !!(newConfig.exames.active),
      hasRams: !!(newConfig.rams.active)
    }};
  } catch (e) { return { sucesso: false, erro: e.message }; }
}

function apiGetColumns(sheetName) {
  if (!sheetName) return [];
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  return sheet ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].filter(h => h) : [];
}

function apiGetMedicamentos(sheetName, colName) { return getUniqueValues_(sheetName, colName); }
function apiGetExamesList(sheetName, colName) { return getUniqueValues_(sheetName, colName); }

// --- 3. PROCESSAMENTO DE DADOS (ESCRITA NA TABELA) ---

function apiImportarReferencias() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const master = SpreadsheetApp.openById(MASTER_DB_ID);
    const data = (master.getSheetByName("Exames_Referencia")||master.getSheets()[0]).getDataRange().getValues();
    let sheet = ss.getSheetByName("Exames_Referencia") || ss.insertSheet("Exames_Referencia");
    sheet.clear().getRange(1,1,data.length,data[0].length).setValues(data);
    return { sucesso: true, msg: "Regras importadas!" };
  } catch(e) { return { sucesso: false, erro: e.message }; }
}

function apiProcessarAnaliseExames() {
  const state = apiGetInitialState();
  if (!state.status.hasExames || !state.status.hasCore) return { sucesso: false, erro: "Configure primeiro." };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfgC = state.config.core; const cfgE = state.config.exames;

  try {
    const refMap = getRefMap_(ss, null);
    const dimData = ss.getSheetByName(cfgC.abaDim).getDataRange().getValues();
    const hDim = dimData.shift();
    const idxP = hDim.indexOf(cfgC.colProntDim), idxN = hDim.indexOf(cfgC.colNasc), idxS = hDim.indexOf(cfgC.colSexo);
    
    const pacMap = {};
    dimData.forEach(r => { if(r[idxP]) pacMap[String(r[idxP]).trim()] = { n: new Date(r[idxN]), s: String(r[idxS]||"").trim().toUpperCase() }; });

    const sheet = ss.getSheetByName(cfgE.aba);
    const data = sheet.getDataRange().getValues();
    const h = data[0];
    const idxPront=h.indexOf(cfgE.colPront), idxNome=h.indexOf(cfgE.colNome), idxVal=h.indexOf(cfgE.colValor), idxDt=h.indexOf(cfgE.colData);
    
    let idxAn = h.indexOf("ANÁLISE"); if(idxAn<0) { idxAn=h.length; sheet.getRange(1,idxAn+1).setValue("ANÁLISE"); }
    let idxMin = h.indexOf("REF_MIN"); if(idxMin<0) { idxMin=h.length+1; sheet.getRange(1,idxMin+1).setValue("REF_MIN"); }
    let idxMax = h.indexOf("REF_MAX"); if(idxMax<0) { idxMax=h.length+2; sheet.getRange(1,idxMax+1).setValue("REF_MAX"); }

    const outAn=[], outMin=[], outMax=[];
    for(let i=1; i<data.length; i++) {
      const r = data[i]; let val = r[idxVal];
      if(typeof val === 'string') val = parseFloat(val.replace(',','.'));
      let resAn="N/A", rMin=null, rMax=null;
      const pac = pacMap[String(r[idxPront]).trim()];

      if(pac && !isNaN(pac.n) && !isNaN(val)) {
        const dtEx = new Date(r[idxDt]);
        if(!isNaN(dtEx)) {
          const dias = Math.floor((dtEx - pac.n)/86400000);
          const regra = getRefRange_(refMap, String(r[idxNome]).trim(), dias, pac.s);
          if(regra) {
            rMin = regra.min; rMax = regra.max;
            if(val < rMin) resAn = "Abaixo"; else if(val > rMax) resAn = "Acima"; else resAn = "Normal";
          } else { resAn = "Sem Ref."; }
        }
      }
      outAn.push([resAn]); outMin.push([rMin]); outMax.push([rMax]);
    }

    sheet.getRange(2, idxAn+1, outAn.length, 1).setValues(outAn);
    sheet.getRange(2, idxMin+1, outMin.length, 1).setValues(outMin);
    sheet.getRange(2, idxMax+1, outMax.length, 1).setValues(outMax);
    return { sucesso: true, msg: `Processado! ${outAn.length} exames.` };
  } catch(e) { return { sucesso: false, erro: e.message }; }
}

// --- 4. MOTORES VISUAIS (DATA FETCHERS) ---

function apiGetEficaciaData(params) {
  const state = apiGetInitialState();
  if (!state.status.hasExames) return { sucesso: false, erro: "Exames não configurados." };

  try {
    const cfgCore = state.config.core; const cfgEx = state.config.exames;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ONE_DAY = 24*60*60*1000;

    // 1. Dimensões (Para filtros laterais)
    const dimData = ss.getSheetByName(cfgCore.abaDim).getDataRange().getValues();
    const hDim = dimData.shift();
    const idxProntD = hDim.indexOf(cfgCore.colProntDim), idxNasc = hDim.indexOf(cfgCore.colNasc), idxSexo = hDim.indexOf(cfgCore.colSexo);
    const pacMeta = {};
    dimData.forEach(r => { pacMeta[String(r[idxProntD]).trim()] = { n: new Date(r[idxNasc]), s: String(r[idxSexo]||"").toUpperCase().charAt(0) }; });

    // 2. Fatos (Para Dose e Data Início)
    const fatosData = ss.getSheetByName(cfgCore.abaFatos).getDataRange().getValues();
    const hFatos = fatosData.shift();
    const idxProntF=hFatos.indexOf(cfgCore.colProntFatos), idxMed=hFatos.indexOf(cfgCore.colMed), idxDtIni=hFatos.indexOf(cfgCore.colDtIni);
    const idxDose = hFatos.findIndex(c => c.toUpperCase().includes("DOSE") || c.toUpperCase().includes("VALOR"));
    
    const diaZero = {};
    fatosData.forEach(r => {
      if (r[idxMed] === params.medicamento) {
        const p = String(r[idxProntF]).trim();
        const dt = new Date(r[idxDtIni]);
        const doseVal = idxDose > -1 ? parseFloat(r[idxDose]) : 0;
        if (!isNaN(dt.getTime())) {
          if (!diaZero[p] || dt < diaZero[p].dt) diaZero[p] = { dt: dt, dose: doseVal||0 };
        }
      }
    });

    // 3. Exames
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
           
           plotData.push({ 
             x: dias, y: val, prontuario: p, 
             refMin: (idxMin>-1&&r[idxMin]!=="")?parseFloat(r[idxMin]):null, 
             refMax: (idxMax>-1&&r[idxMax]!=="")?parseFloat(r[idxMax]):null,
             sexo: meta.s, idade: idade, dose: diaZero[p].dose
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
    
    // Coorte
    const fatosData = ss.getSheetByName(cfgC.abaFatos).getDataRange().getValues();
    const hF = fatosData.shift();
    const idxPF = hF.indexOf(cfgC.colProntFatos), idxMed = hF.indexOf(cfgC.colMed), idxDI = hF.indexOf(cfgC.colDtIni);
    const coorte = {};
    fatosData.forEach(r => { if(r[idxMed]===params.medicamento) { const p=String(r[idxPF]).trim(); const d=new Date(r[idxDI]); if(!coorte[p] || d<coorte[p]) coorte[p]=d; }});
    
    // RAMs
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
        events.push({
          gravidade: String(r[idxG]||"N/D").trim(),
          causalidade: String(r[idxC]||"N/D").trim(),
          descricao: idxDesc>-1 ? String(r[idxDesc]).trim() : "RAM",
          dias: dias,
          prontuario: p
        });
      }
    });
    return { sucesso: true, dados: events, totalExpostos: Object.keys(coorte).length };
  } catch(e) { return { sucesso: false, erro: e.message }; }
}

function apiGetPerfilData(p) { /* Mantido da versão anterior */ 
  const state = apiGetInitialState();
  if (!state.status.hasCore) return { sucesso: false, erro: "Core?" };
  const c=state.config.core; const ss=SpreadsheetApp.getActiveSpreadsheet();
  const fd=ss.getSheetByName(c.abaFatos).getDataRange().getValues(); const dd=ss.getSheetByName(c.abaDim).getDataRange().getValues();
  const hf=fd.shift(); const im=hf.indexOf(c.colMed); const ipf=hf.indexOf(c.colProntFatos);
  const coorte=new Set(); fd.forEach(r=>{ if(r[im]===p.medicamento) coorte.add(String(r[ipf]).trim()); });
  const hd=dd.shift(); const ipd=hd.indexOf(c.colProntDim); const is=hd.indexOf(c.colSexo); const ina=hd.indexOf(c.colNasc);
  const stats={sexo:{}, idade:[]}; const today=new Date();
  dd.forEach(r=>{ if(coorte.has(String(r[ipd]).trim())) {
    const s=r[is]||"N/D"; stats.sexo[s]=(stats.sexo[s]||0)+1;
    const n=new Date(r[ina]); if(!isNaN(n)) stats.idade.push((today-n)/(365.25*86400000));
  }});
  return { sucesso: true, dados: stats };
}

// --- HELPERS ---
function normalize_(s) { return String(s||"").toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, ""); }
function getRefMap_(s,f) { /* Igual v9.1 */ 
  const sheet = s.getSheetByName("Exames_Referencia"); if(!sheet) return [];
  const data = sheet.getDataRange().getValues(); const h = data.shift();
  const find = (k) => h.findIndex(c => k.some(x => normalize_(c).includes(normalize_(x))));
  const iNom=find(["NOME"]), iMinD=find(["MIN_DIAS","IDADE_MIN"]), iMaxD=find(["MAX_DIAS"]), iMinV=find(["VALOR_MIN"]), iMaxV=find(["VALOR_MAX"]), iSex=find(["SEXO"]), iMinAn=find(["ANOS_MIN"]), iMaxAn=find(["ANOS_MAX"]);
  if(iNom<0 || iMinV<0) return [];
  const regras = []; const alvo = normalize_(f);
  data.forEach(r => {
    if(!alvo || normalize_(r[iNom]) === alvo) {
      let minD = parseFloat(r[iMinD]), maxD = parseFloat(r[iMaxD]);
      if(isNaN(minD) && iMinAn>-1 && r[iMinAn]!=="") minD = r[iMinAn]*365;
      if(isNaN(maxD) && iMaxAn>-1 && r[iMaxAn]!=="") maxD = r[iMaxAn]*365;
      if(isNaN(minD)) minD=0; if(isNaN(maxD)) maxD=99999;
      let minV = r[iMinV]; if(typeof minV==='string') minV = parseFloat(minV.replace(',','.'));
      let maxV = r[iMaxV]; if(typeof maxV==='string') maxV = parseFloat(maxV.replace(',','.'));
      let sexRaw = String(r[iSex]||"").toUpperCase();
      let sex = "A"; if(sexRaw.match(/\bM\b/)) sex="M"; else if(sexRaw.match(/\bF\b/)) sex="F";
      regras.push({ n: normalize_(r[iNom]), minD, maxD, min: minV, max: maxV, s: sex });
    }
  });
  return regras;
}
function getRefRange_(regras, nome, dias, sexo) {
  const n = normalize_(nome); const s = String(sexo||"").toUpperCase().charAt(0);
  return regras.find(r => r.n===n && dias>=r.minD && dias<=r.maxD && (r.s==="A" || r.s===s));
}
function getUniqueValues_(s,c) { const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(s); return sh?[...new Set(sh.getRange(2,sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].indexOf(c)+1,sh.getLastRow()-1,1).getValues().map(v=>String(v[0]).trim()))].filter(v=>v).sort():[]; }

function diagnosticarErroReferencia() { /* Igual v9.1 */ 
  const ui=SpreadsheetApp.getUi(); const ss=SpreadsheetApp.getActiveSpreadsheet();
  const exSheet=ss.getSheetByName("Exames"); if(!exSheet) return ui.alert("Aba Exames não encontrada");
  const data=exSheet.getDataRange().getValues(); const h=data[0];
  const idxAnalise=h.indexOf("ANÁLISE"), idxNome=h.findIndex(c=>c.includes("NOME")), idxPront=h.findIndex(c=>c.includes("PRONT")), idxData=h.findIndex(c=>c.includes("DATA"));
  let rowErr=-1; for(let i=1; i<data.length; i++) if(data[i][idxAnalise]==="Sem Ref.") { rowErr=i; break; }
  if(rowErr===-1) return ui.alert("Tudo OK!");
  const p=data[rowErr][idxPront], ex=data[rowErr][idxNome], dt=new Date(data[rowErr][idxData]);
  const state=apiGetInitialState(); const dimSheet=ss.getSheetByName(state.config.core.abaDim);
  const dimData=dimSheet.getDataRange().getValues(); const hD=dimData.shift();
  const iP=hD.indexOf(state.config.core.colProntDim), iN=hD.indexOf(state.config.core.colNasc), iS=hD.indexOf(state.config.core.colSexo);
  let pac=null; for(let r of dimData) if(String(r[iP]).trim()==p) { pac={n:new Date(r[iN]), s:r[iS]}; break; }
  if(!pac) return ui.alert("Paciente não encontrado");
  const dias=Math.floor((dt-pac.n)/86400000);
  const refMap=getRefMap_(ss, ex);
  let msg=`Erro na linha ${rowErr+1}: ${ex} | ${p}\nIdade: ${dias} dias | Sexo: ${pac.s}\n`;
  if(refMap.length===0) msg+="Sem regras para este exame.";
  else {
    let match=false; refMap.forEach(r=>{ if(dias>=r.minD && dias<=r.maxD && (r.s==="A"||r.s===String(pac.s).charAt(0))) match=true; });
    msg += match ? "Regra válida existe. Reprocesse." : "Nenhuma regra cobre esta idade/sexo.";
  }
  ui.alert(msg);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function limparCacheDoSistema() {
  const cache = CacheService.getScriptCache();
  cache.remove('opcoesRamLabPendente');
  cache.remove('opcoesGraficoPendente');
  cache.remove('opcoesPerfilDeUso');
  cache.remove('WEIGHT_MAP');
  // Se usar PropertiesService para configurações permanentes e quiser resetar:
  // PropertiesService.getScriptProperties().deleteProperty('EUM_CONFIG_MASTER');
  
  SpreadsheetApp.getUi().alert('🧹 Cache do Sistema Limpo com Sucesso!');
}
