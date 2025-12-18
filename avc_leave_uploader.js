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
        let s=document.createElement('script');
        s.src='https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
        document.body.appendChild(s);
        await new Promise(r=>s.onload=r);
    }

    let c=document.createElement('div');
    c.style="position:fixed;top:5%;left:50%;transform:translateX(-50%);background:#fff;padding:40px 30px 30px 30px;z-index:999999;box-shadow:0 0 40px 10px rgba(0,0,0,0.8);border-radius:20px;text-align:center;font-family:sans-serif;min-width:500px;";
    c.innerHTML=`
        <div style='margin-bottom:15px;'>
            <span style='font-weight:bold;font-size:26px;color:maroon;'>AVC Leave Uploader</span>
        </div>
        <button id='dl' style='background:#000;color:#fff;padding:10px 14px;border:none;border-radius:8px;margin:8px 0;cursor:pointer;font-size:14px;'>Download Template</button>
        <input type='file' id='up' accept='.xlsx' style='margin:8px;'>
        <button id='ex' style='background:#217346;color:#fff;padding:10px 14px;border:none;border-radius:8px;margin:8px;cursor:pointer;font-size:14px;'>Export to Excel</button>
        <div style='text-align:left;position:relative;margin-top:20px;'>
            <button id='closeBtn' style='background:maroon;color:#fff;padding:7px 14px;border:none;border-radius:8px;cursor:pointer;position:absolute;top:0;left:0;'>Close</button>
        </div>
        <div id="footerText" style="visibility:hidden;margin-top:20px;font-size:13px;color:#000;text-align:center;height:16px;">
            Prepared by Abdullah Shah and Mirza Laiq Ahmed
        </div>
    `;
    document.body.appendChild(c);

    let footer=document.getElementById("footerText"), timer=null, hoverZone=document.createElement("div");
    hoverZone.style="position:absolute;bottom:10px;right:10px;width:30px;height:30px;z-index:1;background:transparent;";
    c.appendChild(hoverZone);
    hoverZone.onmouseenter=()=>{timer=setTimeout(()=>{footer.style.visibility="visible";footer.style.opacity="1";footer.style.transition="opacity 0.3s";},3000)};
    hoverZone.onmouseleave=()=>{clearTimeout(timer);footer.style.opacity="0";footer.style.visibility="hidden"};

    document.getElementById('closeBtn').onclick=()=>c.remove();

    document.getElementById('dl').onclick=()=>{
        let wb=XLSX.utils.book_new();
        let ws=XLSX.utils.aoa_to_sheet([["Emp ID","Attendance Date","Leave Type","Remarks"]]);
        XLSX.utils.book_append_sheet(wb,ws,"Template");
        XLSX.writeFile(wb,"Leave_Template.xlsx");
    };

    document.getElementById('ex').onclick=()=>{
        let t=document.querySelector("table");
        if(!t){alert("❌ Table not found on this page.");return;}
        let csv=Array.from(t.rows).map(r=>Array.from(r.cells).map(c=>`"${c.innerText.trim().replace(/"/g,'""')}"`).join("\t")).join("\n");
        let a=document.createElement("a");
        a.href="data:application/vnd.ms-excel,"+encodeURIComponent(csv);
        a.download="Crystal_Report_Export.xls";
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    };

    document.getElementById("up").onchange=async e=>{
        let f=e.target.files[0];
        if(!f)return;
        let data=await f.arrayBuffer();
        let wb=XLSX.read(data,{type:"array"});
        let ws=wb.Sheets[wb.SheetNames[0]];
        let rows=XLSX.utils.sheet_to_json(ws,{header:1}).slice(1).filter(r=>r[0]&&r[1]&&r[2]&&r[3]);
        let rejected=[],i=0;
        let search=document.querySelector("input.form-control.form-control-sm[type='search']");
        if(!search){alert("❌ Search box not found.");return;}
        function next(){
            if(i>=rows.length){
                if(rejected.length){
                    let csv="EmployeeID,Date,LeaveCode,Remark\n"+rejected.map(r=>r.map(v=>`"${v}"`).join(",")).join("\n");
                    let a=document.createElement("a");
                    a.href=URL.createObjectURL(new Blob([csv],{type:"text/csv"}));
                    a.download="Rejected_Leave_Entries.csv";
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                }
                c.remove();
                return;
            }
            let[id,date,leaveCode,remark]=rows[i++];
            search.value=id;
            search.dispatchEvent(new Event("input"));
            setTimeout(()=>{
                let trList=document.querySelectorAll("table tbody tr"),matched=false;
                trList.forEach(row=>{
                    let txt=row.innerText.replace(/\s+/g," ").trim();
                    if(txt.includes(id)&&txt.includes(date)){
                        let cb=row.querySelector("input[type='checkbox']");
                        if(cb&&!cb.checked)cb.click();
                        let sel=row.querySelector("select");
                        if(sel){sel.value=leaveCode;sel.dispatchEvent(new Event("change",{bubbles:true}));}
                        let inputs=row.querySelectorAll("input[type='text']");
                        if(inputs.length){inputs[inputs.length-1].value=remark;inputs[inputs.length-1].dispatchEvent(new Event("input",{bubbles:true}));}
                        matched=true;
                    }
                });
                if(!matched)rejected.push([id,date,leaveCode,"Emp not found"]);
                next();
            },500);
        }
        next();
    };
})();
