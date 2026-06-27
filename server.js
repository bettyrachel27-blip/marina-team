
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
function isExcelDate(v){return typeof v==="number" && v>30000 && v<60000}
function excelDateLabel(v){
 const d=new Date(Math.round((v-25569)*86400*1000));
 const dd=String(d.getUTCDate()).padStart(2,"0"), mm=String(d.getUTCMonth()+1).padStart(2,"0");
 return `${dd}/${mm}`;
}
function excelTime(v){
 if(v===null||v===undefined||v==="") return "";
 if(typeof v==="number"){
   let frac=v%1; if(frac<0) frac+=1;
   let mins=Math.round(frac*24*60);
   if(mins>=1440) mins-=1440;
   const h=String(Math.floor(mins/60)).padStart(2,"0"), m=String(mins%60).padStart(2,"0");
   return `${h}h${m}`;
 }
 const t=clean(v);
 if(!t) return "";
 const low=t.toLowerCase();
 if(["repos","récup","recup","cp","jf","abs","am","formation"].includes(low)) return t.toUpperCase();
 return t;
}
function cellHasUsefulValue(v){
 if(v===null||v===undefined||v==="") return false;
 if(typeof v==="number") return true;
 return clean(v)!=="";
}
function formatDay(row,base){
 // Ton planning Adelphia a 5 colonnes par jour : début 1, fin 1, début 2, fin 2, total.
 const a=row[base], b=row[base+1], c=row[base+2], d=row[base+3];
 const block=[a,b,c,d].filter(cellHasUsefulValue);
 if(!block.length) return "";
 const special=block.map(excelTime).filter(x=>x && !/^\d{2}h\d{2}$/.test(x));
 if(special.length) return [...new Set(special)].join(" / ");
 const p1=excelTime(a)&&excelTime(b)?`${excelTime(a)}-${excelTime(b)}`:"";
 const p2=excelTime(c)&&excelTime(d)?`${excelTime(c)}-${excelTime(d)}`:"";
 return [p1,p2].filter(Boolean).join(" / ");
}
function looksLikeName(name){
 const n=clean(name);
 if(!n) return false;
 const low=n.toLowerCase();
 if(low.includes("légende")||low.includes("legende")||low.includes("groupe")||low.includes("pause")) return false;
 if(low.includes(":")) return false; // ignore les remarques type "Thierry : poussière"
 if(low.includes("cuisine et salle")) return false;
 if(low.includes("attention")) return false;
 if(/^[0-9.,\s]+$/.test(n)) return false;
 return true;
}
function parseAdelphiaPlanning(wb,team){
 const out=[];
 const sheetRegex=team==="salle"?/planning\s*salle/i:/planning\s*cuisine/i;
 for(const sheet of wb.SheetNames){
   if(!sheetRegex.test(sheet)) continue;
   if(/equipe complet|équipe complet|mod[eè]le|modele|trame|base/i.test(sheet)) continue;
   const ws=wb.Sheets[sheet];
   const rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,blankrows:false,defval:""});
   if(!rows||rows.length<3) continue;
   const dateRowIndex=rows.findIndex(r=>r && r.filter(isExcelDate).length>=3);
   if(dateRowIndex<0) continue;
   const dateRow=rows[dateRowIndex];
   const firstDateCol=dateRow.findIndex(isExcelDate);
   if(firstDateCol<0) continue;
   const dayBases=[];
   for(let d=0; d<7; d++) dayBases.push(firstDateCol + d*5);
   const firstDate=dateRow[dayBases[0]], lastDate=dateRow[dayBases[6]];
   const label=`Planning ${team} du ${excelDateLabel(firstDate)} au ${excelDateLabel(lastDate)}`;
   const planningRows=[];
   for(let r=dateRowIndex+1; r<Math.min(rows.length, dateRowIndex+40); r++){
     const row=rows[r]||[];
     const name=clean(row[1]); // Dans ton Excel, les prénoms sont en colonne B.
     const low=name.toLowerCase();
     if(low.includes("légende")||low.includes("legende")||low.includes("groupe")) break;
     if(!looksLikeName(name)) continue;
     const days=dayBases.map(base=>formatDay(row,base));
     if(days.some(Boolean)) planningRows.push({name,days});
   }
   if(planningRows.length){
     const sortKey=isExcelDate(firstDate)?firstDate:0;
     out.push({id:`${team}_${sortKey}_${sheet.replace(/[^a-zA-Z0-9_-]/g,"_")}`,label,team,status:"draft",rows:planningRows,importedAt:new Date().toISOString(),sortKey});
   }
 }
 out.sort((a,b)=>(a.sortKey||0)-(b.sortKey||0));
 return out.map(({sortKey,...w})=>w);
}
function parseExcel(file,team){
 console.log(`[IMPORT] Lecture fichier ${file} pour équipe ${team}`);
 const wb=XLSX.readFile(file,{cellDates:false,raw:true});
 const weeks=parseAdelphiaPlanning(wb,team);
 console.log(`[IMPORT] ${weeks.length} semaine(s) détectée(s) pour ${team}`);
 if(!weeks.length){
   throw new Error(`Aucune semaine ${team} détectée dans ce fichier. Vérifie que les feuilles commencent par Planning ${team}.`);
 }
 return weeks;
}
app.post("/api/login",(req,res)=>{const d=db(),u=d.users.find(x=>x.id===req.body.id); if(!u||!bcrypt.compareSync(req.body.password,u.passwordHash)) return res.status(401).json({error:"Identifiant ou mot de passe incorrect"}); const user=pub(u); res.json({token:jwt.sign(user,SECRET,{expiresIn:"7d"}),user})});
app.get("/api/users",auth,(req,res)=>{const d=db(); if(req.user.role==="admin") return res.json(d.users.map(pub)); if(req.user.role==="manager") return res.json(d.users.filter(u=>u.team===req.user.team).map(pub)); res.json([req.user])});
app.post("/api/import-planning",auth,upload.single("file"),(req,res)=>{if(req.user.role!=="admin") return res.status(403).json({error:"Réservé Rachel"}); const team=req.body.team; if(!["salle","cuisine"].includes(team)||!req.file) return res.status(400).json({error:"Fichier ou équipe invalide"}); const d=db(),weeks=parseExcel(req.file.path,team),old=d.plannings[team]||[]; d.plannings[team]=weeks.map(w=>{const prev=old.find(x=>x.id===w.id); return prev?{...w,status:prev.status||"draft",publishedAt:prev.publishedAt||null}:w}); notif(d,`Planning ${team} importé`,`${weeks.length} semaine(s) en brouillon.`,team,true); save(d); res.json({ok:true,count:weeks.length})});
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
