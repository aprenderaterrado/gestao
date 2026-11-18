// APRENDER — Backend (Google Apps Script). Rode setup() uma vez e publique como Web App.
const TZ='America/Sao_Paulo';
const SH={Configuracoes:['chave','valor'],Professores:['id','nome','telefone','email','cpf','disciplina','valor_aula','status','obs'],Alunos:['id','nome','responsavel','telefone','email','matricula','mensalidade','status','obs'],Vinculos:['id','aluno_id','professor_id'],Mensalidades:['id','aluno_id','competencia','vencimento','valor','desconto','juros','status'],Recebimentos:['id','mensalidade_id','data','valor','forma'],Aulas:['id','aluno_id','professor_id','data','qtd','duracao','status','obs'],Despesas:['id','descricao','categoria','valor','vencimento','pagamento','tipo','status','obs'],Repasses:['id','professor_id','competencia','valor','data'],Distribuicoes:['id','competencia','resultado','s1','s2','data'],Fechamentos:['id','competencia','data'],Historico:['id','data','acao','tabela','detalhe']};
const EDIT=['Professores','Alunos','Mensalidades','Recebimentos','Aulas','Despesas','Repasses'];
const DEF={escola_nome:'Aprender',preco:85,escola:22,professor:63,s1_nome:'Sócia 1',s1_pct:65,s2_nome:'Sócia 2',s2_pct:35,venc_dia:10,categorias:'Aluguel,Energia elétrica,Água,Internet,IPTU,Material de uso,Material de limpeza,Manutenção de equipamentos,Manutenção da estrutura,Material didático,Outros gastos',formas:'Dinheiro,PIX,Cartão,Transferência,Boleto'};
function ss_(){const P=PropertiesService.getScriptProperties();let id=P.getProperty('SSID');if(id)return SpreadsheetApp.openById(id);const s=SpreadsheetApp.create('Aprender - Banco de Dados');P.setProperty('SSID',s.getId());return s}
function setup(){const s=ss_(),P=PropertiesService.getScriptProperties();
 for(const n in SH){let sh=s.getSheetByName(n);if(!sh){sh=s.insertSheet(n);sh.appendRow(SH[n]);sh.setFrozenRows(1)}}
 const d=s.getSheetByName('Página1')||s.getSheetByName('Sheet1');if(d)s.deleteSheet(d);
 if(!P.getProperty('ADMIN_PASS'))P.setProperty('ADMIN_PASS','aprender123'); // TROQUE nas Propriedades do script
 const c=s.getSheetByName('Configuracoes');if(c.getLastRow()<2)for(const k in DEF)c.appendRow([k,DEF[k]]);
 if(!ScriptApp.getProjectTriggers().some(t=>t.getHandlerFunction()=='backup'))ScriptApp.newTrigger('backup').timeBased().everyWeeks(1).create();
 Logger.log(s.getUrl())}
function backup(){DriveApp.getFileById(ss_().getId()).makeCopy('Backup Aprender '+Utilities.formatDate(new Date(),TZ,'yyyy-MM-dd'))}
function out_(d){return ContentService.createTextOutput(JSON.stringify({ok:true,d:d})).setMimeType(ContentService.MimeType.JSON)}
function err_(m){return ContentService.createTextOutput(JSON.stringify({ok:false,err:m})).setMimeType(ContentService.MimeType.JSON)}
function doPost(e){try{
 const P=PropertiesService.getScriptProperties(),C=CacheService.getScriptCache(),r=JSON.parse(e.postData.contents);
 if(!P.getProperty('SSID'))setup();
 if(r.a==='login'){if(r.p!==P.getProperty('ADMIN_PASS'))return err_('Senha incorreta');const t=Utilities.getUuid();C.put('t_'+t,'1',21600);return out_(t)}
 if(!C.get('t_'+r.t))return err_('auth');
 const L=LockService.getScriptLock();L.waitLock(20000);try{return out_(route_(r,P))}finally{L.releaseLock()}
}catch(x){return err_(x.message)}}
function route_(r,P){switch(r.a){
 case 'load':{const o={cfg:cfg_(),fechados:rows_('Fechamentos').map(x=>x.competencia)};EDIT.concat(['Vinculos','Distribuicoes']).forEach(n=>o[n]=rows_(n));return o}
 case 'save':return save_(r.n,r.o);
 case 'del':return del_(r.n,r.id);
 case 'vinc':return vinc_(r.aluno_id,r.profs);
 case 'gerar':return gerar_(r.ym);
 case 'dash':return calc_(r.ym);
 case 'fechar':return fechar_(r.ym);
 case 'reabrir':if(r.p!==P.getProperty('ADMIN_PASS'))throw new Error('Senha incorreta');del0_('Fechamentos','competencia',r.ym);log_('reabrir','Fechamentos',r.ym);return 1;
 case 'cfg':for(const k in r.c)setCfg_(k,r.c[k]);log_('config','Configuracoes',JSON.stringify(r.c));return 1;
 case 'senha':if(String(r.novo).length<6)throw new Error('Mínimo 6 caracteres');P.setProperty('ADMIN_PASS',r.novo);return 1}
 throw new Error('Ação inválida')}
function rows_(n){const v=ss_().getSheetByName(n).getDataRange().getValues(),h=v.shift();return v.filter(r=>r[0]!=='').map(r=>{const o={};h.forEach((k,i)=>{let x=r[i];if(x instanceof Date)x=Utilities.formatDate(x,TZ,'yyyy-MM-dd');o[k]=x});return o})}
function cfg_(){const o=Object.assign({},DEF),v=ss_().getSheetByName('Configuracoes').getDataRange().getValues();v.slice(1).forEach(x=>{o[x[0]]=/^(preco|escola|professor|s\d_pct|venc_dia)$/.test(x[0])?+x[1]:String(x[1])});return o}
function setCfg_(k,val){const sh=ss_().getSheetByName('Configuracoes'),v=sh.getDataRange().getValues();for(let i=1;i<v.length;i++)if(v[i][0]===k){sh.getRange(i+1,2).setValue(val);return}sh.appendRow([k,val])}
function log_(a,t,d){ss_().getSheetByName('Historico').appendRow([Utilities.getUuid().slice(0,8),Utilities.formatDate(new Date(),TZ,'yyyy-MM-dd HH:mm'),a,t,typeof d=='string'?d:JSON.stringify(d)])}
function fechado_(ym){return rows_('Fechamentos').some(x=>x.competencia===ym)}
function guard_(n,o){const f={Aulas:'data',Despesas:'vencimento',Recebimentos:'data',Repasses:'competencia'}[n];if(f&&o[f]&&fechado_(String(o[f]).slice(0,7)))throw new Error('Período fechado. Reabra o mês para alterar.')}
function num_(x){const v=+x;if(isNaN(v))throw new Error('Valor numérico inválido');return v}
function save_(n,o){
 if(EDIT.indexOf(n)<0)throw new Error('Tabela não permitida');
 const cur=o.id?rows_(n).find(x=>x.id==o.id):null;if(o.id&&!cur)throw new Error('Registro não encontrado');
 if(cur)guard_(n,cur);o=Object.assign({},cur||{},o);guard_(n,o);
 ['valor','mensalidade','valor_aula','qtd','duracao','desconto','juros'].forEach(k=>{if(o[k]!==''&&o[k]!=null)num_(o[k])});
 if(n==='Professores'||n==='Alunos'){if(!String(o.nome||'').trim())throw new Error('Nome obrigatório')}
 if(n==='Aulas'){
  if(!rows_('Vinculos').some(v=>v.aluno_id==o.aluno_id&&v.professor_id==o.professor_id))throw new Error('Professor não está vinculado a este aluno');
  if(!(+o.qtd>0))throw new Error('Informe a quantidade de aulas');
  if(rows_('Aulas').some(a=>a.id!=o.id&&a.aluno_id==o.aluno_id&&a.professor_id==o.professor_id&&a.data==o.data&&+a.duracao==+o.duracao&&a.status==o.status))throw new Error('Aula já lançada (duplicidade)')}
 if(n==='Recebimentos'&&!(+o.valor>0))throw new Error('Valor inválido');
 const sh=ss_().getSheetByName(n),h=SH[n];
 if(!o.id){o.id=Utilities.getUuid().slice(0,8);sh.appendRow(h.map(k=>o[k]==null?'':o[k]));log_('criar',n,o)}
 else{const ids=sh.getRange(2,1,Math.max(sh.getLastRow()-1,1),1).getValues().flat();sh.getRange(ids.indexOf(o.id)+2,1,1,h.length).setValues([h.map(k=>o[k]==null?'':o[k])]);log_('editar',n,o)}
 return o}
function del0_(n,col,val){const sh=ss_().getSheetByName(n),v=sh.getDataRange().getValues(),c=v[0].indexOf(col);for(let i=v.length-1;i>0;i--)if(v[i][c]==val)sh.deleteRow(i+1)}
function del_(n,id){
 if(EDIT.indexOf(n)<0)throw new Error('Tabela não permitida');
 const cur=rows_(n).find(x=>x.id==id);if(!cur)return 1;guard_(n,cur);
 if(n==='Professores'&&rows_('Aulas').some(a=>a.professor_id==id))throw new Error('Professor possui aulas; inative-o em vez de excluir');
 if(n==='Alunos'&&(rows_('Aulas').some(a=>a.aluno_id==id)||rows_('Mensalidades').some(a=>a.aluno_id==id)))throw new Error('Aluno possui histórico; inative-o em vez de excluir');
 if(n==='Mensalidades'&&rows_('Recebimentos').some(a=>a.mensalidade_id==id))throw new Error('Mensalidade possui recebimentos');
 del0_(n,'id',id);if(n==='Professores')del0_('Vinculos','professor_id',id);if(n==='Alunos')del0_('Vinculos','aluno_id',id);
 log_('excluir',n,cur);return 1}
function vinc_(aid,profs){del0_('Vinculos','aluno_id',aid);const sh=ss_().getSheetByName('Vinculos');(profs||[]).forEach(p=>sh.appendRow([Utilities.getUuid().slice(0,8),aid,p]));log_('vincular','Vinculos',{aid,profs});return 1}
function gerar_(ym){const c=cfg_(),ex=rows_('Mensalidades').filter(m=>m.competencia===ym).map(m=>m.aluno_id);let n=0;
 rows_('Alunos').filter(a=>a.status==='ativo'&&+a.mensalidade>0&&ex.indexOf(a.id)<0).forEach(a=>{save_('Mensalidades',{aluno_id:a.id,competencia:ym,vencimento:ym+'-'+('0'+c.venc_dia).slice(-2),valor:+a.mensalidade,desconto:0,juros:0,status:'pendente'});n++});return n}
function calc_(ym){
 const c=cfg_(),sm=x=>x.reduce((a,b)=>a+b,0),inM=d=>String(d||'').slice(0,7)===ym;
 const profs=rows_('Professores'),aulas=rows_('Aulas').filter(a=>inM(a.data)&&(a.status==='realizada'||a.status==='reposicao')),rep=rows_('Repasses').filter(r=>r.competencia===ym);
 const pl=profs.map(p=>{const q=sm(aulas.filter(a=>a.professor_id==p.id).map(a=>+a.qtd)),dev=q*(+p.valor_aula||c.professor),pg=sm(rep.filter(r=>r.professor_id==p.id).map(r=>+r.valor));return{id:p.id,nome:p.nome,aulas:q,devido:dev,pago:pg,receber:dev-pg}}).filter(p=>p.aulas||p.pago);
 const q=sm(aulas.map(a=>+a.qtd)),devProf=sm(pl.map(p=>p.devido));
 const rec=x=>sm(rows_('Recebimentos').filter(r=>x(r.data)).map(r=>+r.valor));
 const mens=rows_('Mensalidades').filter(m=>m.competencia===ym&&m.status!=='cancelado'),recM=rows_('Recebimentos');
 const pend=sm(mens.map(m=>Math.max(0,+m.valor-(+m.desconto||0)+(+m.juros||0)-sm(recM.filter(r=>r.mensalidade_id===m.id).map(r=>+r.valor)))));
 const desp=rows_('Despesas'),dM=desp.filter(d=>inM(d.vencimento)),dP=desp.filter(d=>d.status==='paga'&&inM(d.pagamento||d.vencimento)),cats={};
 dM.forEach(d=>cats[d.categoria]=(cats[d.categoria]||0)+ +d.valor);
 const recebido=rec(inM),dPago=sm(dP.map(d=>+d.valor)),liq=recebido-devProf-dPago;
 const serie=[];for(let i=11;i>=0;i--){const dt=new Date(+ym.slice(0,4),+ym.slice(5)-1-i,1),m=Utilities.formatDate(dt,TZ,'yyyy-MM');serie.push({m:m,rec:rec(d=>String(d).slice(0,7)===m),desp:sm(desp.filter(d=>d.status==='paga'&&String(d.pagamento||d.vencimento).slice(0,7)===m).map(d=>+d.valor))})}
 return{ym:ym,fechado:fechado_(ym),alunos:rows_('Alunos').filter(a=>a.status==='ativo').length,professores:profs.filter(p=>p.status==='ativo').length,aulas:q,receitaAulas:q*c.preco,escola:q*c.escola,devProf:devProf,recebido:recebido,pendente:pend,despesasTotal:sm(dM.map(d=>+d.valor)),despesasPagas:dPago,liquido:liq,s1:Math.round(liq*c.s1_pct)/100,s2:Math.round(liq*c.s2_pct)/100,profs:pl,cats:cats,serie:serie}}
function fechar_(ym){if(fechado_(ym))throw new Error('Mês já fechado');const k=calc_(ym),id=Utilities.getUuid().slice(0,8),d=Utilities.formatDate(new Date(),TZ,'yyyy-MM-dd');
 ss_().getSheetByName('Distribuicoes').appendRow([id,ym,k.liquido,k.s1,k.s2,d]);ss_().getSheetByName('Fechamentos').appendRow([id,ym,d]);log_('fechar','Fechamentos',k);return k}
