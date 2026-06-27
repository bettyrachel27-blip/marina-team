
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

function excelDateToFR(v){
 if(typeof v!=="number" || !isFinite(v)) return "";
 const d = new Date(Math.round((v - 25569) * 86400 * 1000));
 return d.toLocaleDateString("fr-FR",{weekday:"short", day:"2-digit", month:"2-digit"}).replace(".","");
}
function sheetStartSerial(sheetName, rows){
 const serials = [];
 if(rows && rows[1]){
   for(const c of [2,7,12,17,22,27,32]){
     const v = rows[1][c];
     if(typeof v === "number" && v > 30000 && v < 60000) serials.push(v);
   }
 }
 const m = String(sheetName).match(/(\d{1,2})[-\/ ](\d{1,2})/);
 if(serials.length) return Math.min(...serials);
 return 0;
}
function valueToTime(v){
 if(v===null || v===undefined || v==="") return "";
 if(typeof v==="string"){
   const s=v.trim();
   if(!s) return "";
   if(/^\d{1,2}[:h]\d{0,2}$/i.test(s)){
     return s.replace("h",":").replace(/:$/,"h00").replace(":", "h");
   }
   return s;
 }
 if(typeof v==="number" && isFinite(v)){
   // Excel stores hours as fraction of a day.
   if(v >= 0 && v < 1){
     let mins = Math.round(v*24*60);
     mins = ((mins % 1440) + 1440) % 1440;
     const h = Math.floor(mins/60);
     const m = mins%60;
     return String(h).padStart(2,"0")+"h"+String(m).padStart(2,"0");
   }
   return String(v);
 }
 return String(v);
}
function totalToHours(v){
 if(typeof v==="number" && isFinite(v)){
   // If total is a fraction of a day, convert to hours.
   if(v >= 0 && v < 1) return (v*24).toFixed(2).replace(".",",")+"h";
   return String(v);
 }
 return "";
}
function parseDayBlock(row, startCol){
 const vals = [0,1,2,3,4].map(i => row[startCol+i]);
 const textVals = vals.map(v => String(v ?? "").trim()).filter(Boolean);
 const status = textVals.find(s => /^[A-Za-zÀ-ÿ]{2,}$/.test(s) && !/^\d/.test(s));
 if(status) return status.toUpperCase();

 const times = vals.slice(0,4).map(valueToTime).filter(Boolean);
 const total = totalToHours(vals[4]);
 if(times.length >= 4) return `${times[0]}-${times[1]} / ${times[2]}-${times[3]}${total ? " · "+total : ""}`;
 if(times.length === 3) return `${times[0]}-${times[1]} / ${times[2]}${total ? " · "+total : ""}`;
 if(times.length === 2) return `${times[0]}-${times[1]}${total ? " · "+total : ""}`;
 if(times.length === 1) return times[0] + (total ? " · "+total : "");
 if(total) return total;
 return "";
}
function isPlanningSheetName(name, team){
 const s = String(name || "").toLowerCase();
 if(!s.includes("planning")) return false;
 if(team === "salle" && !s.includes("salle")) return false;
 if(team === "cuisine" && !s.includes("cuisine")) return false;
 if(s.includes("equipe complet") || s.includes("équipe complet") || s.includes("modele") || s.includes("modèle") || s.includes("test")) return false;
 return true;
}
function parseExcel(file,team){
 const wb=XLSX.readFile(file,{cellDates:false});
 const dayNames=["Lundi","Mardi","Mercredi","Jeudi","Vendredi","Samedi","Dimanche"];
 const starts=[2,7,12,17,22,27,32];
 const candidates=[];
 for(const sheet of wb.SheetNames){
  if(!isPlanningSheetName(sheet, team)) continue;
  const rows=XLSX.utils.sheet_to_json(wb.Sheets[sheet],{header:1,raw:true,blankrows:false});
  if(!rows||rows.length<3) continue;
  const startSerial=sheetStartSerial(sheet, rows);
  candidates.push({sheet, rows, startSerial});
 }
 // Affiche les semaines les plus récentes d'abord, sans lire la feuille "équipe complet".
 candidates.sort((a,b)=>b.startSerial-a.startSerial);
 const out=[];
 for(const item of candidates){
  const rows=item.rows;
  const labels=starts.map((c,i)=>{
    const d = rows[1] ? excelDateToFR(rows[1][c]) : "";
    return d ? `${dayNames[i]} ${d}` : dayNames[i];
  });
  const planningRows=[];
  for(let r=2;r<rows.length;r++){
    const row=rows[r]||[];
    const name=clean(row[1]);
    if(!name) continue;
    const low=name.toLowerCase();
    if(low.includes("groupe") || low.includes("légende") || low.includes("legende") || low.includes("pause") || low.includes("total")) break;
    if(name.length<2) continue;
    const days=starts.map(c=>parseDayBlock(row,c));
    if(!days.some(Boolean)) continue;
    planningRows.push({name,days});
  }
  if(planningRows.length){
    out.push({
      id:`${team}_${item.sheet.replace(/[^a-zA-Z0-9_-]/g,"_")}`,
      label:item.sheet,
      team,
      status:"draft",
      daysLabels:labels,
      rows:planningRows,
      importedAt:new Date().toISOString(),
      startSerial:item.startSerial
    });
  }
 }
 return out;
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
