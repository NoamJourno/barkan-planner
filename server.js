const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const uploadDir = '/tmp/uploads';
if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir, { recursive: true });
const upload = multer({ dest: uploadDir, limits: { fileSize: 10 * 1024 * 1024 } });

app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname)));

// ============================================================
// SETUP MATRIX
// ============================================================
const BOTTLES = [
  'בקבוק אירופה זית 750 מ"ל',
  'בקבוק בורגונדי ירוק זית אוולושיין',
  'בקבוק בורגונדי שקוף אוולושיון סגל 750',
  'בקבוק בורדו טכניקה י.ענתיק 750 מ"ל',
  'בקבוק הברגה 6828 סוטרן זית 750 פיניציה',
  'בקבוק הוק הברגה שקוף 750 מ"ל',
  'בקבוק הוק שעם ירוק ענתיק 750 מ"ל',
  'בקבוק הוק שעם שקוף 750 מ"ל',
  'בקבוק ירוק 738 750 מ"ל',
  'בקבוק ירוק BORD.RISERVA 76.8 FA 8007779',
  'בקבוק ירוק UF FALANGHINA',
  'בקבוק פלטינום קברנה BORD. BAROS AG',
  'בקבוק קוניקה פניציה  750 S-2588',
  'בקבוק שקוף בורדולז אירופה 6718 750 מ"ל',
  'בקבוק שקוף רוזה 750מ"ל Vetreria Etrusca',
  'בקבוק תירוש שקוף 1 ליטר',
  'בקבוק תירוש שקוף 1 ליטר נמוך- פיטיסיאי',
  'בקבוק הוק אלור 750 מ"ל שקוף',
  'בקבוק בורגונדי דוקאס 750 מ"ל',
];

const MTX = [
  [0,90,90,60,30,60,60,60,45,60,150,180,180,45,350,180,180,60,90],
  [90,0,30,90,180,180,150,150,150,150,150,180,150,150,350,180,180,60,90],
  [90,30,0,90,90,90,90,90,90,90,90,90,90,90,350,180,180,60,90],
  [60,90,90,0,90,90,90,90,90,90,90,90,60,90,350,180,180,60,90],
  [30,180,90,90,0,80,90,90,90,90,90,90,90,90,350,180,180,60,90],
  [60,180,90,90,80,0,70,70,80,90,90,90,90,90,350,180,180,60,90],
  [60,150,90,90,90,70,0,40,90,90,90,90,90,90,350,180,180,60,90],
  [60,150,90,90,90,70,40,0,90,90,90,90,90,90,350,180,180,40,90],
  [45,150,90,90,90,80,90,90,0,60,90,90,90,90,350,180,180,60,90],
  [60,150,90,90,90,90,90,90,60,0,90,90,90,90,350,180,180,60,90],
  [150,150,90,90,90,90,90,90,90,90,0,90,90,90,350,180,180,60,90],
  [180,180,90,90,90,90,90,90,90,90,90,0,90,90,350,180,180,120,120],
  [180,150,90,60,90,90,90,90,90,90,90,90,0,60,350,180,180,90,90],
  [45,150,90,90,90,90,90,90,90,90,90,90,60,0,350,180,180,90,90],
  [350,350,350,350,350,350,350,350,350,350,350,350,350,350,0,350,350,350,350],
  [180,180,180,180,180,180,180,180,180,180,180,180,180,180,350,0,0,999,999],
  [180,180,180,180,180,180,180,180,180,180,180,180,180,180,350,0,0,999,999],
  [60,60,60,60,60,60,60,40,60,60,60,120,90,90,350,999,999,0,180],
  [90,90,90,90,90,90,90,90,90,90,90,120,90,90,350,999,999,180,0],
];

const BOTTLE_SPEEDS = {
  'בקבוק אירופה זית 750 מ"ל': 4500,
  'בקבוק בורגונדי ירוק זית אוולושיין': 4500,
  'בקבוק בורגונדי שקוף אוולושיון סגל 750': 4500,
  'בקבוק שקוף בורדולז אירופה 6718 750 מ"ל': 4500,
  'בקבוק ירוק 738 750 מ"ל': 4000,
  'בקבוק קוניקה פניציה  750 S-2588': 3500,
};

const COLOR_ORDER = { 'לבן': 0, 'כתום': 0, 'רוזה': 1, 'אדום': 2, 'מיץ': 3 };

const HOLIDAYS = new Set([
  '2026-04-01','2026-04-02','2026-04-08','2026-04-09',
  '2026-09-21','2026-09-22','2026-09-30','2026-10-01',
  '2026-10-05','2026-10-06','2026-10-12','2026-10-13',
  '2026-10-19','2026-10-20'
]);

const PRODUCT_MAP = {
  '7001153': { name:'סגל לבן יבש', bottle:'בקבוק בורגונדי שקוף אוולושיון סגל 750', color:'לבן', wineBase:'סגל לבן', tank:'בסיס סגל לבן 2025', tankTotal:311100 },
  '7001447': { name:'סגל לבן חצי יבש', bottle:'בקבוק בורגונדי שקוף אוולושיון סגל 750', color:'לבן', wineBase:'סגל לבן', tank:'בסיס סגל לבן 2025', tankTotal:311100 },
  '7003746': { name:'סגל אדום שביעית', bottle:'בקבוק בורגונדי ירוק זית אוולושיין', color:'אדום', wineBase:'סגל אדום', tank:'בסיס סגל אדום 2025', tankTotal:201878 },
  '7001958': { name:'קלאסיק שרדונה', bottle:'בקבוק בורגונדי ירוק זית אוולושיין', color:'לבן', wineBase:'קלאסיק שרדונה', tank:'בסיס קלאסיק שרדונה 2025', tankTotal:71500 },
  '7004849': { name:'קלאסיק שרדונה ארה"ב', bottle:'בקבוק בורגונדי ירוק זית אוולושיין', color:'לבן', wineBase:'קלאסיק שרדונה', tank:'בסיס קלאסיק שרדונה 2025', tankTotal:71500 },
  '7004847': { name:'בן עמי שרדונה', bottle:'בקבוק בורגונדי ירוק זית אוולושיין', color:'לבן', wineBase:'קלאסיק שרדונה', tank:'בסיס קלאסיק שרדונה 2025', tankTotal:71500 },
  '7005137': { name:'גבעון שרדונה', bottle:'בקבוק בורגונדי ירוק זית אוולושיין', color:'לבן', wineBase:'קלאסיק שרדונה', tank:'בסיס קלאסיק שרדונה 2025', tankTotal:71500 },
  '7008662': { name:'רכסים כרם בן זמרה שרדונה', bottle:'בקבוק בורגונדי ירוק זית אוולושיין', color:'לבן', wineBase:'קלאסיק שרדונה', tank:'בסיס קלאסיק שרדונה 2025', tankTotal:71500 },
  '7003334': { name:'קלאסיק פינו נואר', bottle:'בקבוק בורגונדי ירוק זית אוולושיין', color:'אדום', wineBase:'פינו נואר קלאסיק', tank:'בסיס פינו נואר קלאסיק 2025', tankTotal:50000 },
  '7003732': { name:'רזרב שרדונה שביעית', bottle:'בקבוק בורגונדי ירוק זית אוולושיין', color:'לבן', wineBase:'שרדונה רזרב', tank:'בסיס ספיישל רזרב שרדונה 2025', tankTotal:4003 },
  '7014377': { name:'גולד שרדונה', bottle:'בקבוק בורגונדי ירוק זית אוולושיין', color:'לבן', wineBase:'גולד שרדונה', tank:'בסיס שרדונה גולד 2025', tankTotal:59136 },
  '7001991': { name:'קלאסיק אמרלד ריזלינג', bottle:'בקבוק שקוף בורדולז אירופה 6718 750 מ"ל', color:'לבן', wineBase:'אמרלד ריזלינג קלאסיק', tank:'בסיס קלאסיק אמרלד ריזלינג 2025', tankTotal:70000 },
  '7002646': { name:"פיוז'ן שרדונה קולומברד", bottle:'בקבוק שקוף בורדולז אירופה 6718 750 מ"ל', color:'לבן', wineBase:"פיוז'ן לבן", tank:'בסיס סגל לבן 2025', tankTotal:311100 },
  '7003556': { name:'בן עמי זמורה סמיליון', bottle:'בקבוק שקוף בורדולז אירופה 6718 750 מ"ל', color:'לבן', wineBase:'זמורה לבן', tank:'בסיס סגל לבן 2025', tankTotal:311100 },
  '7013271': { name:'כתום יקב מדבר', bottle:'בקבוק שקוף בורדולז אירופה 6718 750 מ"ל', color:'כתום', wineBase:'מדבר כתום', tank:'בסיס מדבר כתום 2025', tankTotal:13000 },
  '7002643': { name:'בן עמי קברנה', bottle:'בקבוק אירופה זית 750 מ"ל', color:'אדום', wineBase:'קברנה סוביניון קלאסיק', tank:'בסיס קלאסיק קברנה 2025', tankTotal:866700 },
  '7005669': { name:'קלאסיק קברנה צרפת', bottle:'בקבוק אירופה זית 750 מ"ל', color:'אדום', wineBase:'קברנה סוביניון קלאסיק', tank:'בסיס קלאסיק קברנה 2025', tankTotal:866700 },
  '7004856': { name:'קלאסיק קברנה ארה"ב', bottle:'בקבוק אירופה זית 750 מ"ל', color:'אדום', wineBase:'קברנה סוביניון קלאסיק', tank:'בסיס קלאסיק קברנה 2025', tankTotal:866700 },
  '7004854': { name:'קלאסיק קברנה רוסיה', bottle:'בקבוק אירופה זית 750 מ"ל', color:'אדום', wineBase:'קברנה סוביניון קלאסיק', tank:'בסיס קלאסיק קברנה 2025', tankTotal:866700 },
  '7005704': { name:'קלאסיק קברנה רויאל', bottle:'בקבוק אירופה זית 750 מ"ל', color:'אדום', wineBase:'קברנה סוביניון קלאסיק', tank:'בסיס קלאסיק קברנה 2025', tankTotal:866700 },
  '7003697': { name:'קלאסיק קברנה שביעית', bottle:'בקבוק אירופה זית 750 מ"ל', color:'אדום', wineBase:'קברנה סוביניון קלאסיק', tank:'בסיס קלאסיק קברנה 2025', tankTotal:866700 },
  '7005130': { name:'גבעון קברנה', bottle:'בקבוק אירופה זית 750 מ"ל', color:'אדום', wineBase:'קברנה סוביניון קלאסיק', tank:'בסיס קלאסיק קברנה 2025', tankTotal:866700 },
  '7005662': { name:'בן עמי זמורה קברנה', bottle:'בקבוק אירופה זית 750 מ"ל', color:'אדום', wineBase:'זמורה אדום', tank:'בסיס קלאסיק קברנה 2025', tankTotal:866700 },
  '7002644': { name:'בן עמי מרלו', bottle:'בקבוק אירופה זית 750 מ"ל', color:'אדום', wineBase:'מרלו קלאסיק', tank:'בסיס קלאסיק מרלו 2025', tankTotal:251500 },
  '7002490': { name:'מרום גליל מרלו', bottle:'בקבוק אירופה זית 750 מ"ל', color:'אדום', wineBase:'מרלו קלאסיק', tank:'בסיס קלאסיק מרלו 2025', tankTotal:251500 },
  '7003344': { name:'קלאסיק מרלו ארגמן', bottle:'בקבוק אירופה זית 750 מ"ל', color:'אדום', wineBase:'מרלו/ארגמן קלאסיק', tank:'בסיס קלאסיק מרלו 2025', tankTotal:251500 },
  '7005683': { name:'קלאסיק מרלו ארגמן פולין', bottle:'בקבוק אירופה זית 750 מ"ל', color:'אדום', wineBase:'מרלו/ארגמן קלאסיק', tank:'בסיס קלאסיק מרלו 2025', tankTotal:251500 },
  '7002491': { name:'מרום גליל קברנה סינגל', bottle:'בקבוק אירופה זית 750 מ"ל', color:'אדום', wineBase:'קברנה סוביניון קלאסיק', tank:'בסיס קלאסיק קברנה 2025', tankTotal:866700 },
  '7003331': { name:'קלאסיק שיראז', bottle:'בקבוק אירופה זית 750 מ"ל', color:'אדום', wineBase:'שיראז קלאסיק', tank:'בסיס קלאסיק קברנה 2025', tankTotal:866700 },
  '7005680': { name:'קלאסיק שיראז קנדה', bottle:'בקבוק אירופה זית 750 מ"ל', color:'אדום', wineBase:'שיראז קלאסיק', tank:'בסיס קלאסיק קברנה 2025', tankTotal:866700 },
  '7003335': { name:'קלאסיק מלבק', bottle:'בקבוק אירופה זית 750 מ"ל', color:'אדום', wineBase:'מלבק קלאסיק', tank:'בסיס מלבק קלאסיק 2025', tankTotal:30000 },
  '7003333': { name:"קלאסיק פינוטאז'", bottle:'בקבוק אירופה זית 750 מ"ל', color:'אדום', wineBase:'קברנה סוביניון קלאסיק', tank:'בסיס קלאסיק קברנה 2025', tankTotal:866700 },
  '7003012': { name:'פרמייר קברנה מלבק', bottle:'בקבוק אירופה זית 750 מ"ל', color:'אדום', wineBase:'סגל אדום', tank:'בסיס סגל אדום 2025', tankTotal:201878 },
  '7004857': { name:'מהדורה מיוחדת קברנה', bottle:'בקבוק אירופה זית 750 מ"ל', color:'אדום', wineBase:'קברנה סוביניון קלאסיק', tank:'בסיס קלאסיק קברנה 2025', tankTotal:866700 },
  '7005880': { name:'קלאסיק פטיט סירה', bottle:'בקבוק אירופה זית 750 מ"ל', color:'אדום', wineBase:'פטיט סירה קלאסיק', tank:'בסיס פטיט סירה קלאסיק 2025', tankTotal:10000 },
  '7003332': { name:'קלאסיק פטיט סירה ארה"ב', bottle:'בקבוק אירופה זית 750 מ"ל', color:'אדום', wineBase:'פטיט סירה קלאסיק', tank:'בסיס פטיט סירה קלאסיק 2025', tankTotal:10000 },
  '7003710': { name:'פרמייר אדום שביעית', bottle:'בקבוק אירופה זית 750 מ"ל', color:'אדום', wineBase:'מונפורט אדום', tank:'בסיס קלאסיק קברנה 2025', tankTotal:866700 },
  '7005230': { name:"וולדורף אסטוריה פיוז'ן", bottle:'בקבוק אירופה זית 750 מ"ל', color:'אדום', wineBase:'קברנה סוביניון רזרב', tank:'בסיס רזרב קברנה 2025', tankTotal:100000 },
  '7002650': { name:'קברנה רכסים דישון', bottle:'בקבוק אירופה זית 750 מ"ל', color:'אדום', wineBase:'רכסים דישון קברנה', tank:'בסיס רזרב קברנה 2025', tankTotal:100000 },
  '7003360': { name:'יין מתוק כינור', bottle:'בקבוק הברגה 6828 סוטרן זית 750 פיניציה', color:'אדום', wineBase:'סגל אדום מתוק', tank:'בסיס סגל אדום מתוק 2025', tankTotal:446000 },
  '7003314': { name:'רזרב קברנה 12בק', bottle:'בקבוק ירוק 738 750 מ"ל', color:'אדום', wineBase:'קברנה סוביניון רזרב', tank:'בסיס רזרב קברנה 2025', tankTotal:100000 },
  '7002881': { name:"רזרב קברנה 750 12יח'", bottle:'בקבוק ירוק 738 750 מ"ל', color:'אדום', wineBase:'קברנה סוביניון רזרב', tank:'בסיס רזרב קברנה 2025', tankTotal:100000 },
  '7003315': { name:'רזרב מרלו 12בק', bottle:'בקבוק ירוק 738 750 מ"ל', color:'אדום', wineBase:'מרלו רזרב', tank:'בסיס רזרב מרלו 2025', tankTotal:52000 },
  '7003730': { name:'רזרב מרלו שביעית', bottle:'בקבוק ירוק 738 750 מ"ל', color:'אדום', wineBase:'מרלו רזרב', tank:'בסיס רזרב מרלו 2025', tankTotal:52000 },
  '7005661': { name:'מרום גליל קברנה סינגל', bottle:'בקבוק ירוק 738 750 מ"ל', color:'אדום', wineBase:'קברנה סוביניון רזרב', tank:'בסיס רזרב קברנה 2025', tankTotal:100000 },
  '7009310': { name:'גולד מרלו שביעית', bottle:'בקבוק קוניקה פניציה  750 S-2588', color:'אדום', wineBase:'גולד מרלו', tank:'גולד מרלו יין למילוי', tankTotal:52000 },
  '7009300': { name:'גולד קברנה שביעית', bottle:'בקבוק קוניקה פניציה  750 S-2588', color:'אדום', wineBase:'מהדורת גולד קברנה', tank:'מהדורת גולד קברנה יין למילוי', tankTotal:16500 },
  '7003531': { name:'סופרייר קברנה', bottle:'בקבוק בורדו טכניקה י.ענתיק 750 מ"ל', color:'אדום', wineBase:'קברנה סופרייר', tank:'בסיס סופר פרמיום אדום 2025', tankTotal:38400 },
  '7003284': { name:'אלטיטיוד קברנה', bottle:'בקבוק ירוק BORD.RISERVA 76.8 FA 8007779', color:'אדום', wineBase:'קברנה אלטיטוד', tank:'קברנה אלטיטוד 585 יין למילוי', tankTotal:4999 },
  '7013610': { name:'רזרב מוסקטו גולד', bottle:'בקבוק הוק שעם ירוק ענתיק 750 מ"ל', color:'לבן', wineBase:'גולד מוסקטו', tank:'בסיס מוסקטו 2025', tankTotal:75000 },
  '7008668': { name:'רזרב גולד גוורצטרמינר', bottle:'בקבוק הוק שעם ירוק ענתיק 750 מ"ל', color:'לבן', wineBase:'גוורצטרמינר רזרב', tank:'בסיס גולד גוורצטרמינר 2025', tankTotal:15460 },
  '7009783': { name:'ברקן מוסקטו', bottle:'בקבוק הוק הברגה שקוף 750 מ"ל', color:'לבן', wineBase:'גולד מוסקטו', tank:'בסיס מוסקטו 2025', tankTotal:75000 },
  '7009550': { name:'ברקן בטא כתום', bottle:'בקבוק בורגונדי שקוף אוולושיון סגל 750', color:'כתום', wineBase:'יין כתום בטא', tank:'בסיס מדבר כתום 2025', tankTotal:13000 },
  '7003713': { name:'תירוש סגל 1 ליטר', bottle:'בקבוק תירוש שקוף 1 ליטר', color:'מיץ', wineBase:'מיץ ענבים', tank:'בסיס סגל אדום מתוק 2025', tankTotal:446000 },
  '7011730': { name:'תירוש חסידות גור', bottle:'בקבוק תירוש שקוף 1 ליטר נמוך- פיטיסיאי', color:'מיץ', wineBase:'מיץ ענבים', tank:'בסיס סגל אדום מתוק 2025', tankTotal:446000 },
};

function getBottleIdx(b) {
  if (!b) return -1;
  for (let i = 0; i < BOTTLES.length; i++)
    if (BOTTLES[i] === b || b.includes(BOTTLES[i]) || BOTTLES[i].includes(b)) return i;
  return -1;
}
function getSetupTime(b1, b2, c1, c2) {
  const wash = (c1 && c2 && (COLOR_ORDER[c2] ?? 1) < (COLOR_ORDER[c1] ?? 1)) ? 30 : 0;
  if (!b1 || !b2 || b1 === b2) return wash;
  const i = getBottleIdx(b1), j = getBottleIdx(b2);
  return Math.max((i >= 0 && j >= 0) ? MTX[i][j] : 90, wash);
}
function getSpeed(b) { return BOTTLE_SPEEDS[b] || 3000; }
function getLitr(b) {
  if (!b) return 0.75;
  if (b.includes('מגנום') || b.includes('1.5')) return 1.5;
  if (b.includes('1 ליטר')) return 1.0;
  return 0.75;
}
function getWorkDays(start, end) {
  const days = [], d = new Date(start), e = new Date(end);
  while (d <= e) {
    const ds = d.toISOString().slice(0, 10);
    if (d.getDay() !== 5 && d.getDay() !== 6 && !HOLIDAYS.has(ds)) days.push(ds);
    d.setDate(d.getDate() + 1);
  }
  return days;
}
function buildDetail(b1, b2, c1, c2) {
  const parts = [];
  if (b1 && b2 && b1 !== b2) {
    const i = getBottleIdx(b1), j = getBottleIdx(b2);
    parts.push(`החלפת בקבוק: ${(i >= 0 && j >= 0) ? MTX[i][j] : 90} דק'`);
  }
  if (c1 && c2 && (COLOR_ORDER[c2] ?? 1) < (COLOR_ORDER[c1] ?? 1))
    parts.push(`שטיפה (${c1}→${c2}): 30 דק'`);
  return parts.join(' + ');
}

function parseMD16(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const prods = {};
  for (const row of data) {
    if (!row[0] || !row[1]) continue;
    const code = String(row[0]).trim();
    if (!/^\d+$/.test(code)) continue;
    const name = String(row[1]).trim();
    const dateNum = row[2];
    const qty = parseInt(row[3]) || 0;
    if (!qty) continue;
    let deadline = null;
    if (dateNum && typeof dateNum === 'number') {
      deadline = new Date((dateNum - 25569) * 86400 * 1000).toISOString().slice(0, 10);
    } else if (dateNum) {
      deadline = String(dateNum).slice(0, 10);
    }
    if (!prods[code]) prods[code] = { code, name, totalBtl: 0, deadline, earliestDeadline: deadline };
    prods[code].totalBtl += qty;
    if (deadline && (!prods[code].earliestDeadline || deadline < prods[code].earliestDeadline))
      prods[code].earliestDeadline = deadline;
  }
  return Object.values(prods).map(p => {
    const m = PRODUCT_MAP[p.code];
    const bottle = m?.bottle || 'בקבוק אירופה זית 750 מ"ל';
    return {
      ...p, bottle, color: m?.color || 'אדום',
      wineBase: m?.wineBase || p.name,
      tank: m?.tank || 'בסיס קלאסיק קברנה 2025',
      tankTotal: m?.tankTotal || 866700,
      tankQty: Math.round(p.totalBtl * getLitr(bottle) / 5000) * 5000 || 25000,
      unknown: !m,
    };
  });
}

function sortByBottle(products) {
  const byB = {};
  products.forEach(p => { byB[p.bottle] = byB[p.bottle] || []; byB[p.bottle].push(p); });
  const bClr = {};
  Object.entries(byB).forEach(([b, items]) => {
    bClr[b] = Math.min(...items.map(p => COLOR_ORDER[p.color] ?? 1));
  });
  const result = [];
  Object.keys(byB).sort((a, b) => bClr[a] - bClr[b]).forEach(b => {
    byB[b].sort((a, c) => (COLOR_ORDER[a.color] ?? 1) - (COLOR_ORDER[c.color] ?? 1));
    result.push(...byB[b]);
  });
  return result;
}

function scheduleProducts(products, shift = 8) {
  const today = new Date();
  const end = new Date(today); end.setMonth(end.getMonth() + 5);
  const days = getWorkDays(today.toISOString().slice(0, 10), end.toISOString().slice(0, 10));
  const avail = {}; days.forEach(d => { avail[d] = shift; });
  const result = []; let pB = null, pC = null;
  for (const p of products) {
    const sm = pB ? getSetupTime(pB, p.bottle, pC, p.color) : 0;
    const totalH = p.totalBtl / getSpeed(p.bottle);
    let remH = totalH, isFirst = true;
    for (let di = 0; di < days.length && remH > 0.001; di++) {
      const day = days[di], av = avail[day];
      if (av <= 0.001) continue;
      const setupH = isFirst ? sm / 60 : 0;
      const canP = av - setupH;
      if (canP <= 0.001) continue;
      const forDay = Math.min(remH, canP);
      const frac = forDay / totalH;
      const st = remH <= forDay + 0.001 ? (isFirst ? 'full' : 'end') : (isFirst ? 'start' : 'cont');
      result.push({
        product: p, day,
        setup_min: isFirst ? sm : 0,
        prodH: parseFloat(forDay.toFixed(2)),
        btlDay: Math.round(p.totalBtl * frac),
        litDay: Math.round(p.totalBtl * getLitr(p.bottle) * frac),
        tankQtyDay: Math.round((p.tankQty || 25000) * frac / 5000) * 5000,
        spanType: st, isFirst,
        setupDetail: isFirst && sm > 0 ? buildDetail(pB, p.bottle, pC, p.color) : ''
      });
      avail[day] -= (forDay + setupH); remH -= forDay; isFirst = false;
    }
    pB = p.bottle; pC = p.color;
  }
  return result;
}

// API routes
app.post('/api/upload-md16', upload.single('file'), (req, res) => {
  try {
    const buffer = fs.readFileSync(req.file.path);
    const products = parseMD16(buffer);
    fs.unlinkSync(req.file.path);
    res.json({ success: true, products, unknownCount: products.filter(p => p.unknown).length });
  } catch (e) {
    res.status(400).json({ success: false, error: e.message });
  }
});

app.post('/api/optimize', (req, res) => {
  try {
    const { products, shift = 8 } = req.body;
    const sorted = sortByBottle(products);
    const assignments = scheduleProducts(sorted, shift);
    const totalSetup = assignments.reduce((s, a) => s + a.setup_min, 0);
    const totalBtl = assignments.reduce((s, a) => s + a.btlDay, 0);
    const dayH = {};
    assignments.forEach(a => { dayH[a.day] = (dayH[a.day] || 0) + a.prodH; });
    const util = Math.round(
      Object.values(dayH).reduce((s, h) => s + Math.min(h, shift) / shift, 0) /
      Object.keys(dayH).length * 100
    );
    res.json({
      success: true, assignments,
      stats: {
        totalSetupMin: totalSetup,
        totalBottles: totalBtl,
        daysUsed: Object.keys(dayH).length,
        utilization: util,
        setupCount: assignments.filter(a => a.setup_min > 0).length,
      }
    });
  } catch (e) {
    res.status(400).json({ success: false, error: e.message });
  }
});

app.post('/api/update-mapping', (req, res) => {
  const { code, bottle, color, wineBase, tank } = req.body;
  PRODUCT_MAP[code] = { ...(PRODUCT_MAP[code] || {}), bottle, color, wineBase, tank, tankTotal: 50000 };
  res.json({ success: true });
});

// Serve barkan_planner.html for all routes
app.get('*', (req, res) => {
  const candidates = [
    path.join(__dirname, 'public', 'index.html'),
    path.join(__dirname, 'index.html'),
    path.join(__dirname, 'barkan_planner.html'),
  ];
  for (const f of candidates) {
    if (fs.existsSync(f)) return res.sendFile(f);
  }
  res.status(404).send('Page not found');
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Barkan Planner running on port ${PORT}`));
