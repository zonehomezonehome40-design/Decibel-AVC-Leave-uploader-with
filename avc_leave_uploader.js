
(async()=>{
if(location.href!=="https://hrms1.mydecibel.com/TMS/AVCD_BulkLeaves.aspx"){
let p=document.createElement("div");
p.textContent="Not Allowed";
p.style="position:fixed;top:20px;left:50%;transform:translateX(-50%);background:red;color:#fff;padding:12px 25px;border-radius:8px;font-family:sans-serif;font-size:16px;z-index:999999;";
document.body.appendChild(p);
setTimeout(()=>p.remove(),2000);
return;
}

if(!window.XLSX){
let s=document.createElement("script");
s.src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
document.body.appendChild(s);
await new Promise(r=>s.onload=r);
}

let c=document.createElement("div");
c.style="position:fixed;top:5%;left:50%;transform:translateX(-50%);background:#fff;padding:40px 30px;z-index:999999;box-shadow:0 0 40px 10px rgba(0,0,0,0.8);border-radius:20px;text-align:center;font-family:sans-serif;min-width:500px;";
c.innerHTML=`<div style="margin-bottom:15px;"><span style="font-weight:bold;font-size:26px;color:maroon;">AVC Leave Uploader</span></div>
<button id="dl">Download Template</button>
<input type="file" id="up" accept=".xlsx">
<button id="ex">Export to Excel</button>
<button id="closeBtn">Close</button>
<div id="footerText" style="visibility:hidden;margin-top:15px;font-size:13px;">Prepared by Abdullah Shah and Mirza Laiq Ahmed</div>`;

document.body.appendChild(c);
closeBtn.onclick=()=>c.remove();
})();
