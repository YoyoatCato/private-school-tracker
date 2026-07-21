// Cloudflare Worker — on-type mini-sweep backend for the Private School Tracker.
// Given a school name (+ optional state/town) it returns location, street address,
// school type, and the 2026 FFIEC MSA/MD median family income by querying the NCES
// Private School Survey and the open US Census geocoder. CORS-enabled for the dashboard.
// Deploy: see DEPLOY_BACKEND.md. No API keys required.

const FFIEC_MFI_2026 = {"10180": 92100, "10380": 33600, "10420": 97100, "10500": 71200, "10540": 97300, "10580": 123100, "10740": 100400, "10780": 83300, "10900": 107500, "11020": 92400, "11100": 90800, "11180": 119100, "11200": 124400, "11244": 138600, "11260": 130400, "11460": 142300, "11500": 72300, "11540": 111200, "11640": 32700, "11694": 173100, "11700": 103200, "12020": 100100, "12054": 113500, "12100": 103700, "12220": 105500, "12260": 90300, "12420": 134400, "12540": 81900, "12580": 134000, "12620": 98800, "12700": 123600, "12940": 93700, "12980": 80700, "13020": 89900, "13140": 87700, "13220": 69700, "13380": 123300, "13460": 115100, "13740": 112800, "13780": 88600, "13820": 100300, "13900": 113000, "13980": 107100, "14010": 117000, "14020": 108700, "14260": 109900, "14454": 146500, "14500": 150000, "14540": 81100, "14580": 130200, "14740": 129600, "14860": 156800, "15180": 64600, "15260": 92100, "15380": 101500, "15500": 88600, "15540": 123400, "15764": 163800, "15804": 127000, "15940": 90400, "15980": 105700, "16020": 90200, "16180": 90000, "16220": 99500, "16300": 103100, "16540": 99800, "16580": 113000, "16620": 80200, "16700": 117500, "16740": 111400, "16820": 139800, "16860": 97400, "16940": 104400, "16984": 118900, "17020": 89400, "17140": 109000, "17300": 93300, "17410": 103900, "17420": 91700, "17660": 108300, "17780": 104600, "17820": 115700, "17860": 113700, "17900": 92800, "17980": 85200, "18020": 94200, "18140": 111300, "18580": 84800, "18700": 125000, "18880": 102000, "19124": 121100, "19140": 82100, "19300": 102400, "19340": 97500, "19430": 103000, "19460": 88600, "19500": 95100, "19660": 94400, "19740": 144000, "19780": 114800, "19804": 82900, "20020": 79500, "20100": 112100, "20220": 106300, "20260": 101100, "20500": 126700, "20580": 58600, "20740": 101300, "20940": 75500, "20994": 127500, "21060": 92100, "21140": 89900, "21300": 89800, "21340": 73400, "21420": 86800, "21500": 88600, "21660": 96900, "21780": 93000, "21794": 140200, "21820": 116300, "22020": 115200, "22140": 80300, "22180": 81800, "22220": 106900, "22380": 106000, "22420": 82100, "22500": 74000, "22520": 95900, "22540": 105800, "22660": 130400, "22744": 102500, "22900": 76800, "23060": 95400, "23104": 110500, "23224": 177400, "23420": 89300, "23460": 78600, "23540": 91600, "23580": 102000, "23900": 106800, "24020": 94100, "24140": 70300, "24220": 106800, "24260": 94400, "24300": 100600, "24340": 106100, "24420": 76700, "24500": 89000, "24540": 128000, "24580": 105900, "24660": 89700, "24780": 88800, "24860": 101000, "25020": 26300, "25060": 85600, "25180": 101600, "25220": 65800, "25260": 84900, "25420": 109300, "25500": 97700, "25540": 129200, "25620": 79800, "25740": 124000, "25860": 85200, "25940": 110400, "25980": 75000, "26140": 81400, "26300": 81600, "26380": 86600, "26420": 105100, "26580": 85200, "26620": 115100, "26820": 100800, "26900": 107700, "26980": 117100, "27060": 118000, "27100": 101200, "27140": 89400, "27180": 79100, "27260": 108400, "27340": 85400, "27500": 94500, "27620": 98600, "27740": 76800, "27780": 84500, "27860": 74000, "27900": 79400, "27980": 121400, "28020": 108100, "28100": 102200, "28140": 113200, "28420": 105400, "28450": 117300, "28660": 85600, "28700": 81800, "28740": 117300, "28880": 125900, "28940": 98800, "29020": 96200, "29100": 104100, "29180": 87400, "29200": 97300, "29340": 88100, "29404": 142100, "29414": 100300, "29420": 79900, "29460": 83900, "29484": 145000, "29540": 109200, "29620": 102600, "29700": 74000, "29740": 79900, "29820": 98200, "29940": 110400, "30020": 79300, "30140": 98200, "30300": 102900, "30340": 96600, "30460": 102100, "30500": 153200, "30620": 85700, "30700": 106100, "30780": 92800, "30860": 99100, "30980": 80900, "31020": 106700, "31084": 108100, "31140": 98900, "31180": 89500, "31340": 88300, "31420": 78900, "31540": 129000, "31700": 137600, "31740": 100700, "31860": 106500, "31900": 92500, "31924": 129100, "32420": 30600, "32580": 64000, "32780": 98100, "32820": 91500, "32900": 75800, "33124": 89800, "33140": 89300, "33220": 107200, "33260": 109700, "33340": 108400, "33460": 131100, "33500": 95900, "33540": 109200, "33660": 83300, "33700": 94600, "33740": 77400, "33780": 95200, "33860": 88800, "33874": 154300, "34060": 103000, "34100": 80100, "34580": 120200, "34620": 79600, "34740": 77500, "34820": 86000, "34900": 165400, "34940": 121000, "34980": 114300, "35004": 164300, "35084": 141000, "35300": 123200, "35380": 88700, "35614": 108300, "35660": 92200, "35840": 109700, "35980": 111900, "36084": 162800, "36100": 84000, "36220": 91100, "36260": 117900, "36420": 97100, "36500": 122800, "36540": 114200, "36740": 97600, "36780": 100500, "36980": 92100, "37100": 135600, "37140": 98500, "37340": 97000, "37460": 98300, "37620": 74100, "37860": 92800, "37900": 106100, "37964": 91900, "38060": 112400, "38240": 106300, "38300": 108900, "38340": 117600, "38540": 100900, "38660": 30700, "38860": 125000, "38900": 128300, "38940": 102000, "39150": 92200, "39300": 113400, "39340": 119200, "39380": 87700, "39460": 97500, "39540": 104300, "39580": 132300, "39660": 100800, "39740": 102100, "39820": 97100, "39900": 110900, "40060": 113100, "40140": 106500, "40220": 96300, "40340": 127800, "40380": 107000, "40420": 89400, "40484": 145000, "40580": 80200, "40660": 86500, "40900": 124400, "40980": 83200, "41060": 107100, "41100": 105900, "41140": 91800, "41180": 113200, "41304": 107200, "41420": 103400, "41500": 110500, "41540": 104600, "41620": 126100, "41660": 90900, "41700": 101600, "41740": 130900, "41780": 102300, "41884": 197300, "41940": 200900, "41980": 40600, "42020": 129600, "42034": 221900, "42100": 137200, "42140": 118600, "42200": 118600, "42220": 133400, "42340": 107300, "42540": 92400, "42644": 175700, "42680": 105200, "42700": 75600, "43100": 103200, "43300": 95300, "43340": 79200, "43420": 78500, "43580": 97900, "43620": 118100, "43640": 107700, "43780": 85900, "43900": 88600, "44060": 107700, "44100": 114900, "44140": 96700, "44180": 90700, "44220": 83700, "44300": 118300, "44420": 98500, "44700": 108100, "44940": 74200, "45060": 106800, "45104": 127300, "45220": 99300, "45294": 103500, "45460": 95400, "45500": 76700, "45780": 95400, "45820": 98800, "45900": 107800, "45940": 139800, "46060": 99800, "46140": 93500, "46220": 90300, "46300": 94000, "46340": 97300, "46520": 133400, "46540": 98500, "46660": 82600, "46700": 120300, "47020": 81300, "47220": 89100, "47260": 107700, "47300": 76400, "47380": 89600, "47460": 108700, "47580": 98100, "47664": 121300, "47764": 136100, "47930": 114000, "47940": 92100, "48060": 86600, "48140": 103600, "48260": 80600, "48300": 99100, "48424": 107600, "48540": 78900, "48620": 96500, "48660": 91200, "48680": 104100, "48700": 85200, "48864": 115300, "48900": 106100, "49020": 107000, "49180": 93900, "49340": 127200, "49420": 88600, "49620": 105900, "49660": 81900, "49700": 101900, "49740": 79300};
const NONMETRO_MFI_2026 = {"ALABAMA": 74200, "ALASKA": 110100, "ARIZONA": 66300, "ARKANSAS": 71100, "CALIFORNIA": 97100, "COLORADO": 97400, "CONNECTICUT": 124500, "DELAWARE": 103200, "FLORIDA": 80800, "GEORGIA": 77000, "HAWAII": 104400, "IDAHO": 88000, "ILLINOIS": 89700, "INDIANA": 84800, "IOWA": 94100, "KANSAS": 85700, "KENTUCKY": 71900, "LOUISIANA": 67700, "MAINE": 91900, "MARYLAND": 95000, "MASSACHUSETTS": 131700, "MICHIGAN": 83600, "MINNESOTA": 98000, "MISSISSIPPI": 72000, "MISSOURI": 77500, "MONTANA": 91000, "NEBRASKA": 94100, "NEVADA": 105500, "NEW HAMPSHIRE": 117600, "NEW MEXICO": 72900, "NEW YORK": 89900, "NORTH CAROLINA": 78300, "NORTH DAKOTA": 107100, "OHIO": 87700, "OKLAHOMA": 75900, "OREGON": 83600, "PENNSYLVANIA": 86200, "SOUTH CAROLINA": 74600, "SOUTH DAKOTA": 94500, "TENNESSEE": 78000, "TEXAS": 83700, "UTAH": 106300, "VERMONT": 104700, "VIRGINIA": 79900, "WASHINGTON": 97000, "WEST VIRGINIA": 75100, "WISCONSIN": 96000, "WYOMING": 99400};
const STATE_FIPS = {"AL": "01", "AK": "02", "AZ": "04", "AR": "05", "CA": "06", "CO": "08", "CT": "09", "DE": "10", "DC": "11", "FL": "12", "GA": "13", "HI": "15", "ID": "16", "IL": "17", "IN": "18", "IA": "19", "KS": "20", "KY": "21", "LA": "22", "ME": "23", "MD": "24", "MA": "25", "MI": "26", "MN": "27", "MS": "28", "MO": "29", "MT": "30", "NE": "31", "NV": "32", "NH": "33", "NJ": "34", "NM": "35", "NY": "36", "NC": "37", "ND": "38", "OH": "39", "OK": "40", "OR": "41", "PA": "42", "RI": "44", "SC": "45", "SD": "46", "TN": "47", "TX": "48", "UT": "49", "VT": "50", "VA": "51", "WA": "53", "WV": "54", "WI": "55", "WY": "56"};
const STATE_DIV = {"alabama": ["AL", "East South Central"], "alaska": ["AK", "Pacific"], "arizona": ["AZ", "Mountain"], "arkansas": ["AR", "West South Central"], "california": ["CA", "Pacific"], "colorado": ["CO", "Mountain"], "connecticut": ["CT", "New England"], "delaware": ["DE", "South Atlantic"], "florida": ["FL", "South Atlantic"], "georgia": ["GA", "South Atlantic"], "hawaii": ["HI", "Pacific"], "idaho": ["ID", "Mountain"], "illinois": ["IL", "East North Central"], "indiana": ["IN", "East North Central"], "iowa": ["IA", "West North Central"], "kansas": ["KS", "West North Central"], "kentucky": ["KY", "East South Central"], "louisiana": ["LA", "West South Central"], "maine": ["ME", "New England"], "maryland": ["MD", "South Atlantic"], "massachusetts": ["MA", "New England"], "michigan": ["MI", "East North Central"], "minnesota": ["MN", "West North Central"], "mississippi": ["MS", "East South Central"], "missouri": ["MO", "West North Central"], "montana": ["MT", "Mountain"], "nebraska": ["NE", "West North Central"], "nevada": ["NV", "Mountain"], "new hampshire": ["NH", "New England"], "new jersey": ["NJ", "Middle Atlantic"], "new mexico": ["NM", "Mountain"], "new york": ["NY", "Middle Atlantic"], "north carolina": ["NC", "South Atlantic"], "north dakota": ["ND", "West North Central"], "ohio": ["OH", "East North Central"], "oklahoma": ["OK", "West South Central"], "oregon": ["OR", "Pacific"], "pennsylvania": ["PA", "Middle Atlantic"], "rhode island": ["RI", "New England"], "south carolina": ["SC", "South Atlantic"], "south dakota": ["SD", "West North Central"], "tennessee": ["TN", "East South Central"], "texas": ["TX", "West South Central"], "utah": ["UT", "Mountain"], "vermont": ["VT", "New England"], "virginia": ["VA", "South Atlantic"], "washington": ["WA", "Pacific"], "west virginia": ["WV", "South Atlantic"], "wisconsin": ["WI", "East North Central"], "wyoming": ["WY", "Mountain"], "district of columbia": ["DC", "South Atlantic"]};

const UA = { "User-Agent": "Mozilla/5.0 (school-tracker mini-sweep)" };
const GENERIC = new Set(["the","st","st.","saint","school","schools","academy","catholic","christian","high","middle","elementary","preparatory","prep","learning","center","montessori","of","and","college","holy"]);
const TYPE_RULES = [
  ["Roman Catholic", ["roman catholic","catholic","archdiocese","diocese","diocesan","parochial","jesuit","franciscan","salesian","our lady","sacred heart","notre dame"]],
  ["Episcopal", ["episcopal","episcopalian"]],
  ["Lutheran", ["lutheran","missouri synod","wisconsin synod"]],
  ["Jewish", ["jewish","hebrew academy","hebrew day","yeshiva","torah","judaic","chabad","solomon schechter"]],
  ["Muslim", ["islamic","muslim","quran","qur'an","madrasa","madrassa"]],
  ["Special Needs", ["special needs","special education","autism","autistic","dyslexia","learning disabilities","learning differences"]],
  ["Christian", ["christian","baptist","presbyterian","methodist","evangelical","pentecostal","calvary","gospel","adventist","nazarene","assembly of god"]],
  ["Independent", ["montessori","waldorf","independent school","college prep","college preparatory","preparatory academy","prep school","day school","microschool","micro-school","micro school"]],
];

function inferType(text){ const h=(text||"").toLowerCase(); for(const [label,keys] of TYPE_RULES){ if(keys.some(k=>h.includes(k))) return label; } return ""; }
function ncesCore(name){ return name.replace(/[^\w\s.]/g," ").split(/\s+/).filter(w=>w && !GENERIC.has(w.toLowerCase().replace(/\.$/,""))).join(" ") || name; }
function regionForState(abbr){ for(const k in STATE_DIV){ if(STATE_DIV[k][0]===abbr) return STATE_DIV[k][1]; } return ""; }

async function ncesLookup(name, stateAbbr){
  const fips = STATE_FIPS[(stateAbbr||"").toUpperCase()]; if(!fips) return null;
  const core = ncesCore(name); if(!core) return null;
  const body = new URLSearchParams({Search:"1", SchoolName:core, State:fips}).toString();
  let page;
  try {
    const r = await fetch("https://nces.ed.gov/surveys/pss/privateschoolsearch/school_list.asp",
      {method:"POST", headers:{...UA,"Content-Type":"application/x-www-form-urlencoded"}, body});
    page = await r.text();
  } catch(e){ return null; }
  const rx = /school_detail\.asp\?ID=(\w+)"?>\s*([^<]+?)\s*<\/a><br \/><span>([^<]+)<\/span>/g;
  const rows=[]; let m;
  while((m=rx.exec(page))){ rows.push({name:m[2].replace(/&amp;/g,"&").trim(), addr:m[3].replace(/&nbsp;/g," ").replace(/\s+/g," ").trim()}); }
  if(!rows.length) return null;
  const key = new Set(core.toLowerCase().split(/\s+/));
  let cand = rows.filter(r=>[...key].some(w=>r.name.toLowerCase().split(/\s+/).includes(w)));
  if(cand.length!==1) cand = (cand.length? cand : rows);
  if(cand.length!==1) return null;               // ambiguous -> never guess
  const parts = cand[0].addr.split(",").map(s=>s.trim());
  return { full_address: cand[0].addr, town: parts.length>=2 ? parts[1] : "", nces_name: cand[0].name };
}

async function ffiecIncome(address){
  if(!address) return {};
  let g;
  try {
    const u = "https://geocoding.geo.census.gov/geocoder/geographies/onelineaddress?address="
      + encodeURIComponent(address) + "&benchmark=Public_AR_Current&vintage=Current_Current&layers=all&format=json";
    const r = await fetch(u, {headers:UA}); const d = await r.json();
    const mm = d.result && d.result.addressMatches; if(!mm || !mm.length) return {};
    g = mm[0].geographies;
  } catch(e){ return {}; }
  for(const layer of ["Metropolitan Divisions","Metropolitan Statistical Areas"]){
    for(const x of (g[layer]||[])){ const code=(x.GEOID||"").slice(-5); if(FFIEC_MFI_2026[code]) return {tract_mfi:String(FFIEC_MFI_2026[code]), data_year:"2026", mfi_area:x.NAME}; }
  }
  const st=g["States"]||[]; if(st.length){ const nm=(st[0].NAME||"").toUpperCase(); if(NONMETRO_MFI_2026[nm]) return {tract_mfi:String(NONMETRO_MFI_2026[nm]), data_year:"2026", mfi_area:"nonmetro "+st[0].NAME}; }
  return {};
}

// Best-effort location finder when the card has no state: scrape DuckDuckGo HTML
// for the school name + "school" and read the first US state name that appears.
async function findState(name){
  try {
    const u = "https://html.duckduckgo.com/html/?q=" + encodeURIComponent('"'+name+'" school');
    const r = await fetch(u, {headers:UA}); const t = (await r.text()).toLowerCase();
    const names = Object.keys(STATE_DIV).sort((a,b)=>b.length-a.length);
    for(const n of names){ if(t.includes(n)) return STATE_DIV[n][0]; }
  } catch(e){}
  return "";
}

export default {
  async fetch(request){
    const cors = { "Access-Control-Allow-Origin":"*", "Access-Control-Allow-Methods":"POST, OPTIONS", "Access-Control-Allow-Headers":"Content-Type" };
    if(request.method==="OPTIONS") return new Response(null,{headers:cors});
    if(request.method!=="POST") return new Response("POST {name,state?,town?}",{status:405,headers:cors});
    let inp; try { inp = await request.json(); } catch(e){ inp={}; }
    const name=(inp.name||"").trim();
    if(!name) return new Response(JSON.stringify({error:"name required"}),{status:400,headers:{...cors,"Content-Type":"application/json"}});
    let state=(inp.state||"").trim().toUpperCase();
    const out={ type: inferType(name) };
    if(!state) state = await findState(name);         // web search for location if missing
    if(state){ out.state=state; out.region=regionForState(state); }
    const nces = state ? await ncesLookup(name, state) : null;   // NCES address (single confident match)
    if(nces){ out.full_address=nces.full_address; if(nces.town) out.town=nces.town; out.address_unverified=true; }
    const inc = await ffiecIncome(out.full_address);             // FFIEC 2026 MSA/MD income
    Object.assign(out, inc);
    return new Response(JSON.stringify(out),{headers:{...cors,"Content-Type":"application/json"}});
  }
};
