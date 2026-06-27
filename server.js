
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

function clean(x){return String(x??"").replace(/\s+/g," ").trim()}
function excelTimeToText(v){
  if(typeof v!=="number" || !isFinite(v)) return clean(v);
  let frac=((v%1)+1)%1;
  let total=Math.round(frac*24*60);
  let h=Math.floor(total/60)%24, m=total%60;
  return `${String(h).padStart(2,"0")}h${String(m).padStart(2,"0")}`;
}
function excelHoursToText(v){
  if(typeof v!=="number" || !isFinite(v)) return "";
  let mins=Math.round(v*24*60);
  if(mins<=0) return "";
  let h=Math.floor(mins/60), m=mins%60;
  return m?`${h}h${String(m).padStart(2,"0")}`:`${h}h00`;
}
function isStatusText(t){
  const x=clean(t).toUpperCase();
  return /^(CP|REPOS|RECUP|RÉCUP|ABS|ABSENCE|AM|FORMATION|MALADIE|OFF)$/.test(x) || x.includes("REPOS") || x.includes("CP");
}
function isBadName(name){
  const n=clean(name).toLowerCase();
  if(!n) return true;
  if(/^[-–—]+$/.test(n)) return true;
  if(/^(légende|legende|groupe|total|totaux|signature|pause)$/.test(n)) return true;
  if(n.includes("attention") || n.includes("brunch") || n==="cuisine et salle" || n.includes("soirée") || n.includes("soiree")) return true;
  if(/^\d+(\.\d+)?$/.test(n)) return true;
  return false;
}
function parsePlanningCell(block){
  block=block||[];
  const textVals=block.map(clean).filter(Boolean).filter(v=>isNaN(Number(v)));
  const statuses=textVals.filter(isStatusText).map(v=>clean(v).toUpperCase());
  if(statuses.length) return [...new Set(statuses)].join(" / ");

  // Ton fichier Adelphia utilise 5 colonnes par jour : début1, fin1, début2, fin2, total.
  // On ne prend donc que les 4 premières valeurs comme horaires et la 5e comme total.
  const timeVals=block.slice(0,4).filter(v=>typeof v==="number" && isFinite(v) && v>0 && v<1);
  const total=typeof block[4]==="number" && block[4]>0 ? excelHoursToText(block[4]) : "";
  const shifts=[];
  for(let i=0;i+1<timeVals.length;i+=2){
    shifts.push(`${excelTimeToText(timeVals[i])}-${excelTimeToText(timeVals[i+1])}`);
  }
  let out=shifts.join("<br>");
  const notes=textVals.filter(t=>!isStatusText(t));
  if(notes.length) out += (out?"<br>":"") + notes.join("<br>");
  if(total && out) out += `<br><small>${total}</small>`;
  return out;
}
function extractDateStarts(dateRow){
  const starts=[];
  (dateRow||[]).forEach((v,i)=>{
    if(typeof v==="number" && isFinite(v) && v>40000 && v<60000) starts.push(i);
  });
  // structure habituelle si certaines dates sont écrites en texte ou absentes
  if(starts.length<5) return [2,7,12,17,22,27,32];
  return starts.slice(0,7);
}
function parseExcel(file,team){
 const wb=XLSX.readFile(file,{cellDates:false,raw:true});
 const out=[];
 const reTeam=team==="salle"?/planning\s*salle/i:/planning\s*cuisine/i;
 for(const sheet of wb.SheetNames){
  const sheetLow=sheet.toLowerCase();
  if(!reTeam.test(sheet)) continue;
  if(sheetLow.includes("equipe complet") || sheetLow.includes("équipe complet") || sheetLow.includes("modèle") || sheetLow.includes("modele")) continue;
  const rows=XLSX.utils.sheet_to_json(wb.Sheets[sheet],{header:1,raw:true,blankrows:false,defval:null});
  if(!rows||rows.length<3) continue;

  let dateRowIdx=rows.findIndex(r => (r||[]).filter(v=>typeof v==="number" && v>40000 && v<60000).length>=3);
  if(dateRowIdx<0) dateRowIdx=1;
  const starts=extractDateStarts(rows[dateRowIdx]);
  const planningRows=[];

  for(let r=dateRowIdx+1;r<rows.length;r++){
    const row=rows[r]||[];
    const name=clean(row[1]); // dans ton modèle, les prénoms sont en colonne B
    if(isBadName(name)) {
      const n=name.toLowerCase();
      if(n==="légende" || n==="legende" || n==="groupe" || n==="total" || n==="totaux") break;
      continue;
    }
    const days=starts.map(start=>parsePlanningCell(row.slice(start,start+5)));
    if(days.some(Boolean)) planningRows.push({name,days});
  }
  if(planningRows.length){
    out.push({
      id:`${team}_${sheet.replace(/[^a-zA-Z0-9_-]/g,"_")}`,
      label:sheet,
      team,
      status:"draft",
      rows:planningRows,
      importedAt:new Date().toISOString(),
      order:out.length
    });
  }
 }
 // Affiche les dernières feuilles du classeur en premier, car ton fichier est chronologique.
 return out.reverse();
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
