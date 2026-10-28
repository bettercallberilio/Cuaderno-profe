/* Seguimiento Word en el modelo de cinco columnas aportado por la docente. */
var wordTrackingTemplatePromise;
function loadWordTrackingTemplate(){if(!wordTrackingTemplatePromise)wordTrackingTemplatePromise=fetch('./tracking-word-template.json',{cache:'no-cache'}).then(function(r){if(!r.ok)throw Error('No se pudo cargar el modelo Word. Abre la app con conexión.');return r.json()}).catch(function(e){wordTrackingTemplatePromise=null;throw e});return wordTrackingTemplatePromise}
function trackingWordRows(subject,cutoff){
var names=subject==='all'?PROG_SUBJECTS:[subject],rows=[];
names.forEach(function(name){var gi=reportGroupFor(name),mi=S.subj.indexOf(name),mx=S.mx[name],p=PLANS[name];if(!mx||!p||!S.g[gi])return;
var plan=reportPlan(gi,name),entries=reportEntries(gi,mi,p.start,cutoff),content=entries.filter(function(e){return e.t===0}),current=0;
content.forEach(function(e){if(Number.isInteger(e.u)&&mx.ut[e.u])current=Math.max(current,e.u)});
var key=JSON.stringify([S.g[gi].n,name]),confirmed=S.confirmedUTProgress&&S.confirmedUTProgress[key];
if(confirmed&&iso(new Date(confirmed.at))<=cutoff){var ci=mx.ut.findIndex(function(u){return u[0]===confirmed.code});if(ci>=0)current=ci}
var notes=S.reportNotes&&S.reportNotes[cutoff]&&S.reportNotes[cutoff][name]||{},abbr={'Materias primas':'MPPP','Innovación':'INNO','Digitalización':'DIGI','Presentación y venta al público':'PVPP'}[name]||name;
for(var ui=0;ui<=current;ui++){
var u=mx.ut[ui],st=plan.stats[ui],own=content.filter(function(e){return e.u===ui}),actual=reportActual(entries,ui),expected=0;
if(st&&st.centre!==null)expected=st.dates.filter(function(d){return d<=cutoff}).length;
var lines=['Inicio del módulo: '+sh(p.start)+'.'];
if(st&&st.dates.length){lines.push('UT prevista: '+sh(st.dates[0])+' a '+sh(st.dates[st.dates.length-1])+'.')}else lines.push('Fechas previstas de la UT: reparto de periodos de centro pendiente.');
if(own.length){lines.push('Actividades registradas: '+sh(own[0].d)+' a '+sh(own[own.length-1].d)+'.');lines.push('Periodos registrados: '+actual+(st&&st.centre!==null?' / previstos hasta el corte: '+expected+'; diferencia: '+(actual-expected)+'.':'.'))}else lines.push('Sin actividades registradas de esta UT hasta la fecha de corte.');
if(ui===current&&confirmed&&iso(new Date(confirmed.at))<=cutoff)lines.push('UT activa confirmada: '+u[0]+' ('+sh(iso(new Date(confirmed.at)))+').');
if(content.some(function(e){return Number.isInteger(e.u)&&e.u>current})&&ui===current)lines.push('Hay anotaciones de UT posteriores: revisar su etiquetado.');
var measures=notes[ui]||'Sin medidas correctoras registradas.';
if(!notes[ui]&&st&&st.centre!==null&&actual!==expected)measures+=' La diferencia entre periodos previstos y registrados requiere revisión; no acredita por sí sola retraso docente.';
if(plan.unallocated)measures+=' Pendiente concretar los periodos de centro por UT para comparar la previsión.';
rows.push([secondYear(gi)?'2º':'1º',abbr,u[0]+'\n'+u[1],lines.join('\n'),measures])
}
});return rows
}
function wordTrackingText(s){return rx(String(s)).replace(/\n/g,'</w:t><w:br/><w:t xml:space="preserve">')}
async function buildTrackingWord(subject,cutoff){
if(!/^\d{4}-\d{2}-\d{2}$/.test(cutoff))throw Error('Introduce una fecha de corte válida.');
var template=await loadWordTrackingTemplate(),zip=await JSZip.loadAsync(template.base64,{base64:true});
var xml=await zip.file('word/document.xml').async('string'),profile=reportProfile();
[['COURSE',profile.course],['DEPARTMENT',profile.department],['DATE',new Date(cutoff+'T12:00:00').toLocaleDateString('es-ES',{day:'numeric',month:'long',year:'numeric'})],['TEACHER',profile.teacher]].forEach(function(pair){xml=xml.split('{{'+pair[0]+'}}').join(wordTrackingText(pair[1]))});
var rowRe=/<w:tr\b[^>]*>[\s\S]*?<\/w:tr>/g,sourceRows=xml.match(rowRe)||[],proto=sourceRows.find(function(row){return row.indexOf('{{C0}}')>=0});if(!proto)throw Error('El modelo Word no contiene la fila de seguimiento.');
var rows=trackingWordRows(subject,cutoff);if(!rows.length)throw Error('No hay materias configuradas para este seguimiento.');
xml=xml.replace(proto,rows.map(function(values){var row=proto;values.forEach(function(v,i){row=row.split('{{C'+i+'}}').join(wordTrackingText(v))});return row}).join(''));
zip.file('word/document.xml',xml);return zip.generateAsync({type:'uint8array',compression:'DEFLATE'})
}
var wordTrackingButton=$('downloadTrackingWord');if(wordTrackingButton)wordTrackingButton.onclick=function(){var cutoff=$('reportCutoff').value,subject=$('trackingSubject').value;reportDownloadAction(this,async function(){var bytes=await buildTrackingWord(subject,cutoff);downloadReport(bytes,'Seguimiento_programacion_'+cutoff+'.docx','application/vnd.openxmlformats-officedocument.wordprocessingml.document')})};
loadWordTrackingTemplate().catch(function(){});
