
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
function clean(x){return String(x||"").trim()}

function excelDateToJS(serial){
 if(typeof serial!=="number"||!isFinite(serial)) return null;
 const utc=Math.round((serial-25569)*86400*1000);
 return new Date(utc);
}
function parseDateFromSheet(rows){
 const r=(rows&&rows[1])||[];
 for(const v of r){
  const d=excelDateToJS(v);
  if(d && d.getFullYear()>2020 && d.getFullYear()<2035) return d;
 }
 return null;
}
function mondayOf(d){const x=new Date(d); const day=(x.getDay()+6)%7; x.setHours(0,0,0,0); x.setDate(x.getDate()-day); return x}
function formatDateFr(d){return d?d.toLocaleDateString("fr-FR",{day:"2-digit",month:"2-digit"}):""}
function formatTime(v){
 if(v===null||v===undefined||v==="") return "";
 if(typeof v==="string"){
  const t=v.trim();
  if(!t) return "";
  if(["REPOS","RÉCUP","RECUP","CP","JF","ABS","AM"].includes(t.toUpperCase())) return t.toUpperCase();
  return t;
 }
 if(typeof v==="number"){
  // Les dates Excel sont > 1000. Les horaires sont des fractions de jour.
  if(v>=1) return "";
  let minutes=Math.round(v*24*60);
  minutes=((minutes%(24*60))+(24*60))%(24*60);
  const h=String(Math.floor(minutes/60)).padStart(2,"0");
  const m=String(minutes%60).padStart(2,"0");
  return `${h}h${m}`;
 }
 return String(v);
}
function formatDuration(v){
 if(typeof v!=="number") return formatTime(v);
 if(v>=1) return "";
 const hours=v*24;
 if(hours<=0) return "";
 const rounded=Math.round(hours*4)/4;
 return `${String(rounded).replace('.',',')}h`;
}
function formatDay(row,base){
 // Dans le modèle Adelphia : 5 colonnes par jour = début 1 / fin 1 / début 2 / fin 2 / total ou statut.
 const a=row[base], b=row[base+1], c=row[base+2], d=row[base+3], e=row[base+4];
 const status=[a,b,c,d,e].find(x=>typeof x==="string" && x.trim() && ["REPOS","RÉCUP","RECUP","CP","JF","ABS","AM"].includes(x.trim().toUpperCase()));
 if(status) return status.trim().toUpperCase();
 const shifts=[];
 if(a!==undefined&&a!==null&&a!==""&&b!==undefined&&b!==null&&b!=="") shifts.push(`${formatTime(a)}-${formatTime(b)}`);
 if(c!==undefined&&c!==null&&c!==""&&d!==undefined&&d!==null&&d!=="") shifts.push(`${formatTime(c)}-${formatTime(d)}`);
 const total=formatDuration(e);
 if(shifts.length && total) return `${shifts.join(" / ")}\n${total}`;
 if(shifts.length) return shifts.join(" / ");
 return "";
}
function isPlanningSheet(name,team){
 const n=clean(name).toLowerCase();
 if(!n.includes("planning")) return false;
 if(team==="salle" && !n.includes("salle")) return false;
 if(team==="cuisine" && !n.includes("cuisine")) return false;
 if(n.includes("equipe complet")||n.includes("équipe complet")||n.includes("modèle")||n.includes("modele")) return false;
 return true;
}
function parseExcel(file,team){
 const wb=XLSX.readFile(file,{cellDates:false});
 const candidates=[];
 for(const sheet of wb.SheetNames){
  if(!isPlanningSheet(sheet,team)) continue;
  const rows=XLSX.utils.sheet_to_json(wb.Sheets[sheet],{header:1,raw:true,blankrows:false});
  if(!rows||rows.length<4) continue;
  const startDate=parseDateFromSheet(rows);
  if(!startDate) continue;
  const endDate=new Date(startDate); endDate.setDate(endDate.getDate()+6);
  const planningRows=[];
  for(let r=2;r<Math.min(rows.length,40);r++){
   const row=rows[r]||[];
   const name=clean(row[1]);
   if(!name) continue;
   const low=name.toLowerCase();
   if(["groupes","légende","legende","pause","total","signature"].some(x=>low.includes(x))) break;
   if(row.slice(2,37).every(x=>x===undefined||x===null||clean(x)==="")) continue;
   planningRows.push({
    name,
    days:[0,1,2,3,4,5,6].map(i=>formatDay(row,2+i*5))
   });
  }
  if(planningRows.length){
   candidates.push({
    id:`${team}_${sheet.replace(/[^a-zA-Z0-9_-]/g,"_")}`,
    label:`${sheet}`,
    displayLabel:`Semaine du ${formatDateFr(startDate)} au ${formatDateFr(endDate)}`,
    team,
    status:"draft",
    startDate:startDate.toISOString(),
    rows:planningRows,
    importedAt:new Date().toISOString()
   });
  }
 }
 // On affiche seulement les semaines utiles autour d'aujourd'hui : semaine en cours + semaines à venir.
 const now=new Date();
 const current=mondayOf(now);
 current.setDate(current.getDate()-7); // garde aussi la semaine précédente au cas où
 const filtered=candidates
   .filter(w=>new Date(w.startDate)>=current)
   .sort((a,b)=>new Date(a.startDate)-new Date(b.startDate))
   .slice(0,8);
 return filtered.length ? filtered : candidates.sort((a,b)=>new Date(b.startDate)-new Date(a.startDate)).slice(0,8);
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
