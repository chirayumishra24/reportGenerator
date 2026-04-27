

function findClass9Column(headers) {
  return headers.find(h => {
    const lower = String(h || '').toLowerCase();
    const isBaseline = (
      lower.includes('% in ix') || 
      lower.includes('ix %') || 
      lower.includes('class 9') || 
      lower.includes('9th class') || 
      lower.includes('ix percent') || 
      lower.includes('ix marks') || 
      lower.includes('baseline') ||
      (lower.includes('9th') && (lower.includes('%') || lower.includes('percentage')))
    );
    return isBaseline && !lower.includes('+30') && !lower.includes('target');
  }) || null;
}

function toSafeNumber(value) {
  const num = parseFloat(value);
  return Number.isFinite(num) ? num : null;
}

const headers = ["s.no","Enrollment No.","Name","Father Name","Mother Name","DOB","Gender","English       80","Hindi   80","Sanskrit 80","French   80","Maths  80","Science   80","Social Science  80","IT Th  50","Grand Total","% in IX","% in IX+30","ENG 100 IX","English      +30","Hindi +  30","Sanskrit + 30","French +  30","Maths + 30","Science  + 30","Social Science  +  30"];

const class9Col = findClass9Column(headers);
console.log('Found Class 9 Column:', class9Col);

const rowData = [1,"2013-2014/0705","AADISHWAR PAREEK","MR. ASHISH PAREEK","MS. CHANDA PAREEK",39692,"MALE",32,37,null,null,21,34,37,41,161,40.25,70.25,40,62,67,null,null,51,64,67];

const row = {};
headers.forEach((h, i) => { row[h] = rowData[i]; });

const val = class9Col ? row[class9Col] : null;
console.log('Raw value from row:', val);
console.log('toSafeNumber(val):', toSafeNumber(val));
