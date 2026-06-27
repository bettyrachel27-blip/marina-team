
const express=require("express"),jwt=require("jsonwebtoken"),bcrypt=require("bcryptjs"),multer=require("multer"),XLSX=require("xlsx"),fs=require("fs"),path=require("path");
const app=express(),PORT=process.env.PORT||3000,SECRET=process.env.JWT_SECRET||"change_me_marina_team";
const DB=path.join(__dirname,"data.json"),upload=multer({dest:path.join(__dirname,"uploads")});
app.use(express.json({limit:"25mb"})); app.use(express.static(path.join(__dirname,"public")));
function pw(p){return bcrypt.hashSync(p,10)}
function initDb(){return {users:[
{id:"rachel",name:"Rachel Betty",passwordHash:pw("2802"),role:"admin",team:"admin",poste:"Responsable F&B",contract:39},
{id:"florian",name:"Florian Garcia",passwordHash:pw("0000"),role:"manager",team:"cuisine",poste:"Chef de cuisine",contract:39},
{id:"wissem",name:"Wissem Kharroubi",passwordHash:pw("0000"),role:"manager",team:"salle",poste:"Responsable de salle",contract:39},
{id:"anthony",name:"Anthony Fossati",passwordHash:pw("0000"),role:"manager",team:"salle",poste:"Maître d'hôtel",contract:39},
{id:"david",name:"David",passwordHash:pw("0000"),role:"employee",team:"cuisine",poste:"Second de cuisine",contract:39},
{id:"joanna",name:"Joanna",passwordHash:pw("0000"),role:"employee",team:"cuisine",poste:"Second de cuisine",contract:39},
{id:"thi_cuisine",name:"Thi",passwordHash:pw("0000"),role:"employee",team:"cuisine",poste:"Second de cuisine",contract:39},
{id:"ali_belagha",name:"Ali Belagha",passwordHash:pw("0000"),role:"employee",team:"cuisine",poste:"Commis cuisine",contract:39},
{id:"ali_plonge",name:"Ali",passwordHash:pw("0000"),role:"employee",team:"cuisine",poste:"Plongeur",contract:25},
{id:"kamel",name:"Kamel",passwordHash:pw("0000"),role:"employee",team:"cuisine",poste:"Plongeur",contract:35},
{id:"sebastien",name:"Sébastien",passwordHash:pw("0000"),role:"employee",team:"salle",poste:"1er chef de rang",contract:39},
{id:"eloise",name:"Éloïse",passwordHash:pw("0000"),role:"employee",team:"salle",poste:"Chef de rang",contract:39},
{id:"yanis",name:"Yanis Pacull",passwordHash:pw("0000"),role:"employee",team:"salle",poste:"Chef de rang",contract:35},
{id:"thierry",name:"Thierry",passwordHash:pw("0000"),role:"employee",team:"salle",poste:"Commis / Runner",contract:39},
{id:"quynh",name:"Quynh",passwordHash:pw("0000"),role:"employee",team:"salle",poste:"Petit-déjeuner",contract:39},
{id:"sevinche",name:"Sevinche",passwordHash:pw("0000"),role:"employee",team:"salle",poste:"Alternante",contract:35}
],plannings:{salle:[],cuisine:[]},hours:[],requests:[],notifications:[],processes:[
{id:1,title:"Ouverture restaurant",category:"Salle",body:"Réservations, mise en place, briefing, allergènes, terrasse."},
{id:2,title:"Réclamation client",category:"Service",body:"Écouter, reformuler, prévenir responsable, noter l'incident, proposer solution adaptée."}
],notices:[{id:1,title:"Bienvenue sur Marina Team",body:"Application interne salle et cuisine.",author:"Rachel Betty"}]}}
function db(){if(!fs.existsSync(DB)) fs.writeFileSync(DB,JSON.stringify(initDb(),null,2)); return JSON.parse(fs.readFileSync(DB,"utf8"))}
function save(d){fs.writeFileSync(DB,JSON.stringify(d,null,2))}
function pub(u){return {id:u.id,name:u.name,role:u.role,team:u.team,poste:u.poste,contract:u.contract}}
function auth(req,res,next){try{req.user=jwt.verify((req.headers.authorization||"").replace("Bearer ",""),SECRET);next()}catch(e){res.status(401).json({error:"Non connecté"})}}
function canSee(u,t){return u.role==="admin"||u.role==="manager"||u.team===t}
function notif(d,title,body,team=null,admin=false){d.notifications.unshift({id:Date.now()+Math.random(),title,body,team,admin,createdAt:new Date().toISOString()})}


function clean(x){return String(x??"").replace(/\s+/g," ").trim()}
function norm(x){return clean(x).normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase()}
function isExcelSerialDate(v){return typeof v==="number" && v>30000 && v<60000}
function isJsDate(v){return Object.prototype.toString.call(v)==='[object Date]' && !isNaN(v)}
function serialToDate(v){return new Date(Math.round((v-25569)*86400*1000))}
function dateLabelFromDate(d){const dd=String(d.getUTCDate()).padStart(2,"0"), mm=String(d.getUTCMonth()+1).padStart(2,"0"); return `${dd}/${mm}`}
function excelDateLabel(v){
 if(isExcelSerialDate(v)) return dateLabelFromDate(serialToDate(v));
 if(isJsDate(v)) return dateLabelFromDate(v);
 const t=clean(v);
 const m=t.match(/(\d{1,2})[\/\-.](\d{1,2})(?:[\/\-.]\d{2,4})?/);
 if(m) return `${m[1].padStart(2,"0")}/${m[2].padStart(2,"0")}`;
 return t || "?";
}
function dateSortKey(v){
 if(isExcelSerialDate(v)) return v;
 if(isJsDate(v)) return Math.round(v.getTime()/86400000)+25569;
 const t=clean(v); const m=t.match(/(\d{1,2})[\/\-.](\d{1,2})(?:[\/\-.](\d{2,4}))?/);
 if(m){let y=m[3]?Number(m[3]):new Date().getFullYear(); if(y<100)y+=2000; return Math.round(Date.UTC(y,Number(m[2])-1,Number(m[1]))/86400000)+25569}
 return 0;
}
function isDateLike(v){
 if(isExcelSerialDate(v)||isJsDate(v)) return true;
 const t=norm(v); if(!t) return false;
 return /\d{1,2}[\/\-.]\d{1,2}/.test(t) || /\b(lundi|mardi|mercredi|jeudi|vendredi|samedi|dimanche)\b/.test(t);
}
function isTimeNumber(v){return typeof v==="number" && v>=0 && v<1}
function excelTime(v){
 if(v===null||v===undefined||v==="") return "";
 if(isTimeNumber(v)){
   let mins=Math.round(v*24*60); if(mins>=1440) mins-=1440;
   const h=String(Math.floor(mins/60)).padStart(2,"0"), m=String(mins%60).padStart(2,"0");
   return `${h}h${m}`;
 }
 const t=clean(v); if(!t) return "";
 const low=norm(t);
 const special=["repos","repo","recup","récup","cp","conges","congés","jf","abs","absence","am","arret","arrêt","maladie","formation","ecole","école"];
 if(special.some(x=>low===norm(x)||low.includes(norm(x)))) return t.toUpperCase();
 const m=t.match(/(\d{1,2})\s*[:hH]\s*(\d{2})?/);
 if(m) return `${String(Number(m[1])).padStart(2,"0")}h${String(m[2]||"00").padStart(2,"0")}`;
 if(/^\d{1,2}$/.test(t)) return `${String(Number(t)).padStart(2,"0")}h00`;
 return t;
}
function cellHasUsefulValue(v){return !(v===null||v===undefined||v===""||clean(v)==="")}
function isHourText(x){return /^\d{2}h\d{2}$/.test(x)}
function formatCells(cells){
 const values=cells.filter(cellHasUsefulValue).map(excelTime).filter(Boolean);
 if(!values.length) return "";
 const specials=values.filter(x=>!isHourText(x));
 if(specials.length) return [...new Set(specials)].join(" / ");
 const parts=[];
 for(let i=0;i<values.length;i+=2){
   if(values[i]&&values[i+1]) parts.push(`${values[i]}-${values[i+1]}`);
   else if(values[i]) parts.push(values[i]);
 }
 return parts.join(" / ");
}
function looksLikeName(name){
 const n=clean(name); if(!n) return false;
 const low=norm(n);
 if(low.includes("legende")||low.includes("groupe")||low.includes("pause")||low.includes("total")||low.includes("attention")) return false;
 if(low.includes(":")) return false;
 if(low.includes("cuisine et salle")) return false;
 if(/^[0-9.,\s]+$/.test(n)) return false;
 if(n.length>60) return false;
 return /[A-Za-zÀ-ÿ]/.test(n);
}
function findDateColumns(row){
 const cols=[];
 for(let i=0;i<row.length;i++) if(isDateLike(row[i])) cols.push(i);
 if(cols.length>=7) return cols.slice(0,7);
 if(cols.length>=1){
   const start=cols[0];
   return Array.from({length:7},(_,d)=>start+d*5);
 }
 return [];
}
function parseSheetRows(rows,team,sheet){
 const out=[];
 for(let dateRowIndex=0; dateRowIndex<rows.length; dateRowIndex++){
   const row=rows[dateRowIndex]||[];
   const dateCount=row.filter(isDateLike).length;
   const dayNameCount=row.filter(v=>/\b(lundi|mardi|mercredi|jeudi|vendredi|samedi|dimanche)\b/.test(norm(v))).length;
   if(dateCount<1 && dayNameCount<3) continue;
   const dayCols=findDateColumns(row);
   if(dayCols.length<7) continue;
   const firstDate=row[dayCols[0]], lastDate=row[dayCols[6]];
   const label=`Planning ${team} du ${excelDateLabel(firstDate)} au ${excelDateLabel(lastDate)}`;
   const planningRows=[];
   const firstDayCol=dayCols[0];
   for(let r=dateRowIndex+1; r<Math.min(rows.length, dateRowIndex+60); r++){
     const rr=rows[r]||[];
     const left=rr.slice(0,Math.max(1,firstDayCol)).map(clean).filter(looksLikeName);
     const name=left.length?left[left.length-1]:clean(rr[1]||rr[0]);
     const low=norm(name);
     if(low.includes("legende")||low.includes("groupe")) break;
     if(!looksLikeName(name)) continue;
     const days=[];
     for(let d=0; d<7; d++){
       const start=dayCols[d];
       const end=d<6?dayCols[d+1]:Math.min(rr.length,start+5);
       days.push(formatCells(rr.slice(start,end)));
     }
     if(days.some(Boolean)) planningRows.push({name,days});
   }
   if(planningRows.length){
     const sortKey=dateSortKey(firstDate)||dateSortKey(lastDate)||dateRowIndex;
     out.push({id:`${team}_${sortKey}_${sheet.replace(/[^a-zA-Z0-9_-]/g,"_")}_${dateRowIndex}`,label,team,status:"draft",rows:planningRows,importedAt:new Date().toISOString(),sortKey});
   }
 }
 return out;
}
function parseAdelphiaPlanning(wb,team){
 const out=[];
 const strict=team==="salle"?/planning\s*salle/i:/planning\s*cuisine/i;
 let sheets=wb.SheetNames.filter(s=>strict.test(s) && !/equipe complet|équipe complet|mod[eè]le|modele|trame|base/i.test(s));
 // Si les onglets ne sont pas nommés exactement, on scanne quand même tout le fichier.
 if(!sheets.length) sheets=wb.SheetNames.filter(s=>!/equipe complet|équipe complet|mod[eè]le|modele|trame|base/i.test(s));
 for(const sheet of sheets){
   const ws=wb.Sheets[sheet];
   const rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,blankrows:false,defval:""});
   if(rows&&rows.length) out.push(...parseSheetRows(rows,team,sheet));
 }
 const seen=new Set();
 return out.sort((a,b)=>(b.sortKey||0)-(a.sortKey||0)).filter(w=>{const k=w.label+"|"+w.rows.map(r=>r.name).join(","); if(seen.has(k)) return false; seen.add(k); return true}).map(({sortKey,...w})=>w);
}
function parseExcel(file,team){
 console.log(`[IMPORT] Lecture fichier ${file} pour équipe ${team}`);
 const wb=XLSX.readFile(file,{cellDates:false,raw:true});
 const weeks=parseAdelphiaPlanning(wb,team);
 console.log(`[IMPORT] ${weeks.length} semaine(s) détectée(s) pour ${team}`);
 if(!weeks.length) throw new Error(`Aucune semaine ${team} détectée. L'import scanne maintenant tous les onglets : vérifie que le fichier contient bien les dates et les prénoms.`);
 return weeks;
}
function mergeWeeks(oldWeeks,newWeeks){
 const map=new Map((oldWeeks||[]).map(w=>[w.id,w]));
 for(const w of newWeeks){
   const prev=map.get(w.id);
   map.set(w.id, prev?{...w,status:prev.status||"draft",publishedAt:prev.publishedAt||null}:w);
 }
 return Array.from(map.values()).sort((a,b)=>dateSortKey((b.label.match(/du ([^ ]+)/)||[])[1])-dateSortKey((a.label.match(/du ([^ ]+)/)||[])[1]));
}
app.post("/api/login",(req,res)=>{const d=db(),u=d.users.find(x=>x.id===req.body.id); if(!u||!bcrypt.compareSync(req.body.password,u.passwordHash)) return res.status(401).json({error:"Identifiant ou mot de passe incorrect"}); const user=pub(u); res.json({token:jwt.sign(user,SECRET,{expiresIn:"7d"}),user})});
app.get("/api/users",auth,(req,res)=>{const d=db(); if(req.user.role==="admin") return res.json(d.users.map(pub)); if(req.user.role==="manager") return res.json(d.users.filter(u=>u.team===req.user.team).map(pub)); res.json([req.user])});
app.post("/api/import-planning",auth,upload.single("file"),(req,res)=>{
 try{
  if(req.user.role!=="admin") return res.status(403).json({error:"Réservé Rachel"});
  const team=req.body.team;
  if(!["salle","cuisine"].includes(team)||!req.file) return res.status(400).json({error:"Fichier ou équipe invalide"});
  const d=db(),weeks=parseExcel(req.file.path,team),oldWeeks=d.plannings[team]||[];
  d.plannings[team]=mergeWeeks(oldWeeks,weeks);
  notif(d,`Planning ${team} importé`,`${weeks.length} nouvelle(s) semaine(s) détectée(s) et ajoutée(s) en brouillon.`,team,true);
  save(d);
  res.json({ok:true,count:weeks.length,labels:weeks.map(w=>w.label)});
 }catch(e){
  console.error("[IMPORT ERREUR]",e);
  res.status(400).json({error:e.message||"Erreur import Excel"});
 }finally{try{if(req.file)fs.unlinkSync(req.file.path)}catch{}}
});
app.get("/api/plannings",auth,(req,res)=>{const d=db(),r={}; for(const t of ["salle","cuisine"]){if(!canSee(req.user,t))continue; r[t]=req.user.role==="admin"?d.plannings[t]:d.plannings[t].filter(w=>w.status==="published")} res.json(r)});
app.post("/api/plannings/:team/:id/publish",auth,(req,res)=>{if(req.user.role!=="admin")return res.status(403).json({error:"Réservé Rachel"}); const d=db(),w=d.plannings[req.params.team]?.find(x=>x.id===req.params.id); if(!w)return res.status(404).json({error:"Semaine introuvable"}); w.status="published";w.publishedAt=new Date().toISOString();notif(d,`Nouveau planning ${req.params.team}`,`${w.label} est disponible.`,req.params.team,false);save(d);res.json({ok:true})});
app.post("/api/plannings/:team/:id/unpublish",auth,(req,res)=>{if(req.user.role!=="admin")return res.status(403).json({error:"Réservé Rachel"}); const d=db(),w=d.plannings[req.params.team]?.find(x=>x.id===req.params.id); if(!w)return res.status(404).json({error:"Semaine introuvable"}); w.status="draft";save(d);res.json({ok:true})});
app.post("/api/hours",auth,(req,res)=>{const d=db(),{week,days,total}=req.body; d.hours=d.hours.filter(h=>!(h.userId===req.user.id&&h.week===week)); d.hours.push({id:Date.now(),userId:req.user.id,userName:req.user.name,team:req.user.team,week,days,total,status:"pending",createdAt:new Date().toISOString()}); notif(d,"Heures saisies",`${req.user.name} a envoyé ${total}h à valider.`,req.user.team,false); save(d); res.json({ok:true})});
app.get("/api/hours",auth,(req,res)=>{const d=db(); if(req.user.role==="admin")return res.json(d.hours); if(req.user.role==="manager")return res.json(d.hours.filter(h=>h.team===req.user.team)); res.json(d.hours.filter(h=>h.userId===req.user.id))});
app.post("/api/hours/:id/validate",auth,(req,res)=>{if(req.user.role!=="admin")return res.status(403).json({error:"Seule Rachel valide"}); const d=db(),h=d.hours.find(x=>String(x.id)===String(req.params.id)); if(!h)return res.status(404).json({error:"Introuvable"}); h.status="validated"; h.validatedAt=new Date().toISOString(); save(d);res.json({ok:true})});
app.post("/api/requests",auth,(req,res)=>{const d=db(),r={id:Date.now(),userId:req.user.id,userName:req.user.name,team:req.user.team,...req.body,status:"pending",createdAt:new Date().toISOString()}; d.requests.unshift(r); notif(d,"Nouvelle demande CP/repos",`${req.user.name} a fait une demande.`,req.user.team,false); save(d); res.json({ok:true,r})});
app.get("/api/requests",auth,(req,res)=>{const d=db(); if(req.user.role==="admin")return res.json(d.requests); if(req.user.role==="manager")return res.json(d.requests.filter(r=>r.team===req.user.team)); res.json(d.requests.filter(r=>r.userId===req.user.id))});
app.get("/api/notices",auth,(req,res)=>res.json(db().notices)); app.get("/api/processes",auth,(req,res)=>res.json(db().processes));
app.get("/api/notifications",auth,(req,res)=>{const d=db(); res.json(d.notifications.filter(n=>n.admin?req.user.role==="admin":(req.user.role==="admin"||req.user.role==="manager"||!n.team||n.team===req.user.team)))});
app.get("*",(req,res)=>res.sendFile(path.join(__dirname,"public","index.html")));
app.listen(PORT,()=>console.log(`Marina Team lancé : http://localhost:${PORT}`));
