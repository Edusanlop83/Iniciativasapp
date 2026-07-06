/* ACE · Motor de plantillas HTML descargables (formato B)
   Genera un HTML autocontenido, editable y con exportación a Word (.doc) y PDF (impresión nativa).
   Uso:  const {html, filename} = ACEPlantilla.generar(tipo, data);
         ACEPlantilla.descargar(tipo, data);
   data = { nombre, lider, sponsor, area, fecha, folio, necesidad, objhip, alctec, criterio, prioridad, iniciativas, hitos } */
(function(global){
  "use strict";

  // ---- utilidades ----
  function esc(s){ return String(s==null?"":s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;"); }
  function noAcc(s){ return String(s||"").normalize("NFD").replace(/[\u0300-\u036f]/g,""); }
  function userShort(name){
    var n = noAcc(name||"Usuario").trim().replace(/\s+/g," ");
    if(!n) return "Usuario";
    var p = n.split(" ");
    var base = p.length>=2 ? (p[0][0]+p[p.length-1]) : p[0];
    return base.replace(/[^A-Za-z0-9]/g,"");
  }
  function ymd(fecha){
    var d = fecha ? new Date(fecha+"T00:00:00") : new Date();
    if(isNaN(d)) d = new Date();
    var m=("0"+(d.getMonth()+1)).slice(-2), da=("0"+d.getDate()).slice(-2);
    return ""+d.getFullYear()+m+da;
  }
  function filenameBase(tipo, data){
    return "ACE-"+tipo+"_"+userShort(data&&data.lider)+"_"+ymd(data&&data.fecha);
  }

  // ---- helpers de render (HTML) ----
  function box(value, ph, pre, multi){
    var cls = "box"+(pre?" pre":"")+(multi?" multi":"");
    return '<div class="'+cls+'" contenteditable="true" data-ph="'+esc(ph||"")+'">'+esc(value||"")+'</div>';
  }
  function lbl(t){ return '<div class="lbl">'+esc(t)+'</div>'; }
  function fieldFull(label, value, ph, pre, multi){
    return '<div class="field">'+lbl(label)+box(value,ph,pre,multi)+'</div>';
  }
  function gridRow(items, multi){
    var cells = items.map(function(it){
      return '<div class="gcell">'+lbl(it.label)+box(it.value,it.ph,it.pre,multi)+'</div>';
    }).join("");
    return '<div class="grid g'+items.length+'">'+cells+'</div>';
  }
  function section(n,t){ return '<div class="sec"><span class="sn">'+n+' · </span>'+esc(t)+'</div>'; }
  function note(t){ return '<div class="note">'+esc(t)+'</div>'; }
  function table(headers, rows){
    var th = headers.map(function(h){ return '<th>'+esc(h)+'</th>'; }).join("");
    var tr = rows.map(function(r){
      var tds = r.map(function(c){
        var pre = c && c.pre, v = (c&&c.v)||"";
        return '<td class="'+(pre?"pre":"")+'" contenteditable="true">'+esc(v)+'</td>';
      }).join("");
      return '<tr>'+tds+'</tr>';
    }).join("");
    return '<table class="tbl"><thead><tr>'+th+'</tr></thead><tbody>'+tr+'</tbody></table>';
  }
  function join(arr){ return (arr&&arr.length)? arr.join(" · ") : ""; }

  // ---- callout común ----
  var RESP = ["Responsable del llenado:","Líder del proyecto que realiza la solicitud de registro."];
  function callout(title, lines){
    var body = '<div class="co-t">'+esc(title)+'</div>';
    lines.forEach(function(ln){
      body += '<div class="co-l"><b>'+esc(ln[0])+'</b> '+esc(ln[1])+'</div>';
    });
    return '<div class="callout">'+body+'</div>';
  }

  // ---- cuerpos por tipo ----
  function bodyProyecto(d){
    var s = "";
    s += callout("¿Qué es este documento?", [
      ["Propósito:","charter ligero para formalizar el arranque de un proyecto ACE con la información mínima de gobierno; parte de los datos capturados en la solicitud y completa lo necesario para iniciar."],
      RESP, ["Órgano de gobierno:","PMO → Comité."] ]);
    s += section(1,"Información General");
    s += fieldFull("Nombre del proyecto", d.nombre, "", true, false);
    s += gridRow([{label:"Líder / PO",value:d.lider,pre:true},{label:"Sponsor",value:d.sponsor,pre:true},{label:"Business Owner",value:"",ph:"Por completar",pre:false}]);
    s += gridRow([{label:"Área solicitante",value:d.area,pre:true},{label:"Fecha",value:d.fecha,pre:true},{label:"Folio Monday",value:d.folio,pre:true}]);
    s += section(2,"Propósito y objetivo");
    s += fieldFull("Objetivo del proyecto", d.objhip, "", true, true);
    s += fieldFull("Necesidad / justificación de negocio", d.necesidad, "", true, true);
    s += section(3,"Objetivos de éxito y KPIs");
    s += note("Define entre 2 y 4 objetivos de éxito, cada uno con su métrica y meta.");
    s += table(["#","Objetivo de éxito","KPI","Meta"], [[{v:"1",pre:true},{},{},{}],[{v:"2",pre:true},{},{},{}],[{v:"3",pre:true},{},{},{}],[{v:"4",pre:true},{},{},{}]]);
    s += section(4,"Alcance");
    s += gridRow([{label:"Incluye",value:d.alctec,pre:true},{label:"No incluye",value:"",ph:"Ej. Referidos presenciales, canales no digitales",pre:false}], true);
    s += section(5,"Entregables principales");
    s += fieldFull("Entregables", "", "Lista los entregables clave (producto, integraciones, documentación, capacitación, etc.)", false, true);
    s += section(6,"Roles, responsabilidades e involucrados");
    s += table(["Áreas/Roles","Nombre / Área","Responsabilidad"], [
      [{v:"Encargado de Negocio / Funcional",pre:true},{},{}],
      [{v:"Encargado de Operación",pre:true},{},{}],
      [{v:"Líder Técnico",pre:true},{},{}] ]);
    s += note("El llenado inicial de este documento es responsabilidad del Líder del proyecto solicitante.");
    s += section(7,"Cronograma de alto nivel");
    s += table(["Fase / Hito","Inicio (est.)","Fin (est.)","Entregable asociado"], [[{},{},{},{}],[{},{},{},{}],[{},{},{},{}],[{},{},{},{}]]);
    s += section(8,"Supuestos y restricciones");
    s += gridRow([{label:"Supuestos",value:"",ph:"Ej. Disponibilidad de datos en HubSpot",pre:false},{label:"Restricciones",value:"",ph:"Ej. Presupuesto y ventana de liberación",pre:false}], true);
    s += section(9,"Riesgos iniciales");
    s += table(["#","Riesgo","Impacto","Mitigación"], [[{v:"1",pre:true},{},{},{}],[{v:"2",pre:true},{},{},{}],[{v:"3",pre:true},{},{},{}],[{v:"4",pre:true},{},{},{}]]);
    s += section(10,"Criterios de aceptación");
    s += fieldFull("Criterios de aceptación", d.criterio, "", true, true);
    s += section(11,"Dependencias e integraciones");
    s += fieldFull("Sistemas / terceros involucrados", "", "Ej. EK9/EK10, MuleSoft, HubSpot, Genesys, proveedores externos", false, true);
    s += section(12,"Alineación estratégica");
    s += gridRow([{label:"Iniciativas ACE que aporta",value:join(d.iniciativas),pre:true},{label:"Momentos Hito que habilita",value:join(d.hitos),pre:true}], true);
    s += section(13,"Aprobaciones");
    s += gridRow([{label:"Solicitante",value:"",ph:"Nombre / firma / fecha"},{label:"PMO",value:"",ph:"Nombre / firma / fecha"},{label:"Business Owner",value:"",ph:"Nombre / firma / fecha"}]);
    return s;
  }

  function bodyPiloto(d){
    var s = "";
    s += callout("¿Qué es un piloto?", [
      ["Propósito:","validar en entorno real y acotado la viabilidad operativa de una propuesta antes de escalar a proyecto formal; parte de los datos de la solicitud y completa el plan del piloto."],
      RESP, ["Órgano de gobierno:","PMO → Comité."] ]);
    s += section(1,"Información General");
    s += fieldFull("Nombre del piloto", d.nombre, "", true, false);
    s += gridRow([{label:"Líder / PO",value:d.lider,pre:true},{label:"Sponsor",value:d.sponsor,pre:true},{label:"Business Owner",value:"",ph:"Por completar"}]);
    s += gridRow([{label:"Área solicitante",value:d.area,pre:true},{label:"Fecha",value:d.fecha,pre:true},{label:"Folio Monday",value:d.folio,pre:true}]);
    s += section(2,"Objetivo e hipótesis");
    s += fieldFull("Objetivo del piloto", d.objhip, "", true, true);
    s += fieldFull("Hipótesis a validar", d.objhip, "", true, true);
    s += fieldFull("Necesidad / justificación", d.necesidad, "", true, true);
    s += section(3,"Alcance acotado");
    s += gridRow([{label:"Incluye",value:d.alctec,pre:true},{label:"No incluye",value:"",ph:"Ej. Otras zonas, otros canales, producción"}], true);
    s += section(4,"Población / muestra / entorno");
    s += fieldFull("Usuarios, volumen, zona y sistemas reales involucrados", "", "Ej. 3 asesores, zona norte, campañas Meta reales, HubSpot productivo", false, true);
    s += section(5,"Métricas y criterios go/no-go");
    s += note("El criterio de éxito viene del registro; complétalo con línea base, meta y la decisión asociada.");
    s += table(["Métrica","Línea base","Meta","Decisión (go/no-go)"], [[{v:d.criterio||"",pre:!!d.criterio},{},{},{}],[{},{},{},{}],[{},{},{},{}]]);
    s += section(6,"Duración y plan");
    s += gridRow([{label:"Inicio",value:"",ph:"dd/mm/aaaa"},{label:"Fin",value:"",ph:"dd/mm/aaaa"},{label:"Duración",value:"",ph:"Ej. 6 semanas"}]);
    s += fieldFull("Hitos del piloto", "", "Ej. Semana 1 configuración · 2–5 operación · 6 evaluación", false, true);
    s += section(7,"Recursos");
    s += gridRow([{label:"Personas",value:"",ph:"Ej. 2 asesores"},{label:"Ambientes",value:"",ph:"Ej. HubSpot prod."},{label:"Licencias",value:"",ph:"Ej. 3 bot"},{label:"Herramientas",value:"",ph:"Ej. Meta Ads"}]);
    s += section(8,"Riesgos");
    s += table(["#","Riesgo","Impacto","Mitigación"], [[{v:"1",pre:true},{},{},{}],[{v:"2",pre:true},{},{},{}],[{v:"3",pre:true},{},{},{}]]);
    s += section(9,"Resultados esperados y decisión de escalamiento");
    s += fieldFull("Qué evidencia llevaría a escalar, ajustar o detener", "", "Ej. Si conversión ≥ 15%, se propone proyecto; si < 10%, se detiene", false, true);
    s += section(10,"Alineación estratégica");
    s += gridRow([{label:"Iniciativas ACE que aporta",value:join(d.iniciativas),pre:true},{label:"Momentos Hito que habilita",value:join(d.hitos),pre:true}], true);
    s += section(11,"Aprobaciones");
    s += gridRow([{label:"Solicitante",value:"",ph:"Nombre / firma / fecha"},{label:"PMO",value:"",ph:"Nombre / firma / fecha"},{label:"Business Owner",value:"",ph:"Nombre / firma / fecha"}]);
    return s;
  }

  function bodyPoC(d){
    var s = "";
    s += callout("¿Qué es este subproceso?", [
      ["Objetivo:","validar la viabilidad técnica, funcional, arquitectónica u operativa de una propuesta antes del diseño definitivo; reduce la incertidumbre y facilita decisiones informadas."],
      ["Responsable Principal:","Arquitectura de Soluciones / Consultor Tecnológico."],
      RESP, ["Órgano de Gobierno:","PMO → Comité Tecnológico."] ]);
    s += section(1,"Información General");
    s += fieldFull("Nombre de la PoC", d.nombre, "", true, false);
    s += gridRow([{label:"Responsable / Líder PO",value:d.lider,pre:true},{label:"Sponsor",value:d.sponsor,pre:true},{label:"Business Owner",value:"",ph:"Por completar"}]);
    s += gridRow([{label:"Área solicitante",value:d.area,pre:true},{label:"Fecha",value:d.fecha,pre:true},{label:"Folio Monday",value:d.folio,pre:true}]);
    s += fieldFull("Arquitecto Responsable", "", "Nombre del arquitecto asignado", false, false);
    s += section(2,"Definición de la PoC");
    s += fieldFull("Objetivo — ¿qué queremos validar?", "", "Ej. Confirmar si la integración es viable sin middleware", false, true);
    s += fieldFull("Problema o incertidumbre — ¿qué duda existe?", d.necesidad, "", true, true);
    s += fieldFull("Hipótesis técnica — ¿qué creemos que ocurrirá?", d.objhip, "", true, true);
    s += section(3,"Alcance y tecnología a probar");
    s += fieldFull("Tecnología / componente a probar", d.alctec, "", true, true);
    s += gridRow([{label:"Incluye",value:"",ph:"Ej. Conexión API REST en ambiente de pruebas"},{label:"No incluye",value:"",ph:"Ej. Producción, volumen real"}], true);
    s += section(4,"Criterios de éxito");
    s += fieldFull("Criterios medibles de aceptación", d.criterio, "", true, true);
    s += section(5,"Tiempo estimado y recursos");
    s += gridRow([{label:"Inicio",value:"",ph:"dd/mm/aaaa"},{label:"Fin",value:"",ph:"dd/mm/aaaa"},{label:"Duración",value:"",ph:"días / semanas"}]);
    s += gridRow([{label:"Personas",value:"",ph:"Ej. 1 arquitecto"},{label:"Ambientes",value:"",ph:"Ej. Sandbox"},{label:"Licencias",value:"",ph:"Ej. Trial"},{label:"Herramientas",value:"",ph:"Ej. Postman"}]);
    s += section(6,"Riesgos (máx. 3–4)");
    s += table(["#","Riesgo","Impacto","Mitigación"], [[{v:"1",pre:true},{},{},{}],[{v:"2",pre:true},{},{},{}],[{v:"3",pre:true},{},{},{}]]);
    s += section(7,"Entregables");
    s += table(["Código","Evidencias","Documento","Demo","ADR"], [[{v:"☐"},{v:"☐"},{v:"☐"},{v:"☐"},{v:"☐"}]]);
    s += section(8,"Alineación estratégica");
    s += gridRow([{label:"Iniciativas ACE que aporta",value:join(d.iniciativas),pre:true},{label:"Momentos Hito que habilita",value:join(d.hitos),pre:true}], true);
    s += section(9,"Aprobaciones");
    s += gridRow([{label:"Solicitante",value:"",ph:"Nombre / firma / fecha"},{label:"PMO",value:"",ph:"Nombre / firma / fecha"},{label:"Business Owner",value:"",ph:"Nombre / firma / fecha"}]);
    return s;
  }

  var TITLES = { Proyecto:"Registro de Proyecto — Mini-Charter", Piloto:"Registro de Piloto — Plan de Piloto", PoC:"Registro de Prueba de Concepto (PoC)" };
  var FOOTLBL = { Proyecto:"Registro de Proyecto", Piloto:"Registro de Piloto", PoC:"Registro de PoC" };
  function bodyFor(tipo,d){ return tipo==="Piloto"?bodyPiloto(d):tipo==="PoC"?bodyPoC(d):bodyProyecto(d); }

  // ---- CSS del documento generado (pantalla + impresión) ----
  var CSS = [
"*{box-sizing:border-box;margin:0;padding:0}",
"body{background:#EEEEEA;color:#242424;font-family:'DM Sans',system-ui,Arial,sans-serif;font-size:14px;line-height:1.5;padding:0}",
".toolbar{position:sticky;top:0;z-index:50;background:#0A1138;color:#fff;display:flex;gap:10px;align-items:center;padding:12px 18px}",
".toolbar .tt{font-family:'DM Sans',sans-serif;font-weight:700;font-size:13px;margin-right:auto}",
".toolbar button{font-family:'DM Sans',sans-serif;font-weight:700;font-size:13px;border:0;border-radius:9px;padding:9px 16px;cursor:pointer;background:#E8400C;color:#fff}",
".toolbar button.ghost{background:transparent;border:1.5px solid rgba(255,255,255,.35);color:#fff}",
".page{max-width:820px;margin:20px auto;background:#F4F0EA;padding:0 0 30px}",
".topbar{height:9px;background:#E8400C}",
".head{display:flex;justify-content:space-between;align-items:flex-start;padding:22px 34px 0}",
".head h1{font-family:Fraunces,Georgia,serif;font-weight:700;color:#E8400C;font-size:23px;line-height:1.1}",
".brand{text-align:right}.brand b{font-size:13px;color:#242424}.brand span{display:block;font-size:11px;color:#8A8A8A}",
".callout{margin:16px 34px 0;background:#FBE4D8;border-left:5px solid #E8400C;border-radius:8px;padding:14px 18px}",
".callout .co-t{font-weight:700;color:#242424;margin-bottom:6px}",
".callout .co-l{font-size:13px;color:#242424;margin-top:3px}",
".sec{font-family:Fraunces,Georgia,serif;font-weight:700;color:#1A1A1A;font-size:16px;margin:22px 34px 8px;padding:6px 0 6px 12px;border-left:5px solid #E8400C;border-bottom:1px solid #B8AE9F}",
".sec .sn{color:#E8400C}",
".field{margin:0 34px 12px}",
".lbl{font-weight:700;font-size:11px;letter-spacing:.5px;text-transform:uppercase;color:#8B7050;margin:8px 0 4px}",
".box{border:1.5px solid #B8AE9F;border-radius:7px;background:#fff;padding:8px 10px;min-height:34px;font-size:13px;color:#242424}",
".box.pre{background:#FDEEE4}",
".box.multi{min-height:64px}",
".box:empty:before,.box[data-ph]:empty:before{content:attr(data-ph);color:#b7b2a8;font-style:italic}",
".box:focus{outline:0;border-color:#E8400C;box-shadow:0 0 0 3px rgba(229,74,18,.12)}",
".grid{display:flex;gap:14px;margin:0 34px 12px}",
".grid .gcell{flex:1 1 0}",
".note{margin:4px 34px 10px;font-size:12px;font-style:italic;color:#8B7050}",
".tbl{width:calc(100% - 68px);margin:0 34px 12px;border-collapse:collapse;font-size:12.5px}",
".tbl th{background:#EDE7DE;color:#8B7050;text-transform:uppercase;font-size:10.5px;letter-spacing:.5px;text-align:left;padding:7px 9px;border:1px solid #B8AE9F}",
".tbl td{border:1px solid #B8AE9F;padding:7px 9px;min-height:26px;background:#fff}",
".tbl td.pre{background:#FDEEE4}",
".tbl td:focus{outline:0;box-shadow:inset 0 0 0 2px rgba(229,74,18,.25)}",
".foot{margin:22px 34px 0;padding-top:10px;border-top:1px solid #B8AE9F;font-size:11px;color:#8A8A8A}",
"@media print{ body{background:#fff} .toolbar{display:none!important} .page{margin:0;max-width:none;box-shadow:none} .box:focus,.tbl td:focus{box-shadow:none} .box,.tbl td,.tbl th{-webkit-print-color-adjust:exact;print-color-adjust:exact} }"
  ].join("\n");

  function exportScript(fnameBase){
    return [
"var FN=", JSON.stringify(fnameBase), ";document.title=FN;",
"function dl(b,n){var a=document.createElement('a');a.href=URL.createObjectURL(b);a.download=n;document.body.appendChild(a);a.click();setTimeout(function(){URL.revokeObjectURL(a.href);a.remove();},500);}",
"function toWord(){var css=document.getElementById('acecss').innerHTML;var body=document.getElementById('doc').innerHTML;",
"var h=\"<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word'><head><meta charset='utf-8'><style>\"+css+\"</style></head><body>\"+body+\"</body></html>\";",
"var b=new Blob(['\\ufeff'+h],{type:'application/msword'});dl(b,FN+'.doc');}",
"document.getElementById('btnWord').addEventListener('click',toWord);",
"document.getElementById('btnPdf').addEventListener('click',function(){window.print();});"
    ].join("");
  }

  function generar(tipo, data){
    data = data||{};
    var fnameBase = filenameBase(tipo, data);
    var fecha = data.fecha || ymd(); // para el footer mostramos fecha legible si viene
    var docInner =
      '<div class="topbar"></div>'+
      '<div class="head"><h1>'+esc(TITLES[tipo]||TITLES.Proyecto)+'</h1>'+
        '<div class="brand"><b>PROGRAMA ACE</b><span>Hogares Unión · GIM Holdings</span></div></div>'+
      bodyFor(tipo, data)+
      '<div class="foot">Control de versiones — v1.0 · Plantilla editable · '+esc(FOOTLBL[tipo]||"")+' · Modelo ACE · Folio '+esc(data.folio||"—")+'</div>';
    var toolbar =
      '<div class="toolbar"><span class="tt">Plantilla ACE · '+esc(tipo)+' — edítala y descárgala</span>'+
      '<button id="btnWord" class="ghost">Descargar Word</button>'+
      '<button id="btnPdf">Descargar PDF</button></div>';
    var html =
      '<!DOCTYPE html><html lang="es"><head><meta charset="utf-8">'+
      '<meta name="viewport" content="width=device-width, initial-scale=1">'+
      '<title>'+esc(fnameBase)+'</title>'+
      '<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>'+
      '<link href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,600;9..144,700&family=DM+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">'+
      '<style id="acecss">'+CSS+'</style></head><body>'+
      toolbar+'<div class="page" id="doc">'+docInner+'</div>'+
      '<script>'+exportScript(fnameBase)+'<\/script>'+
      '</body></html>';
    return { html:html, filename:fnameBase+".html" };
  }

  function descargar(tipo, data){
    var g = generar(tipo, data);
    var blob = new Blob([g.html], {type:"text/html;charset=utf-8"});
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob); a.download = g.filename;
    document.body.appendChild(a); a.click();
    setTimeout(function(){ URL.revokeObjectURL(a.href); a.remove(); }, 500);
    return g.filename;
  }

  global.ACEPlantilla = { generar:generar, descargar:descargar, filenameBase:filenameBase, userShort:userShort, ymd:ymd };
})(typeof window!=="undefined"?window:this);
