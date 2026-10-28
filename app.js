const form=document.getElementById("convert-form");
const fileInput=document.getElementById("pdf-file");
const dropzone=document.getElementById("dropzone");
const fileTitle=document.getElementById("file-title");
const fileSubtitle=document.getElementById("file-subtitle");
const fileInfo=document.getElementById("file-info");
const button=document.getElementById("convert-button");
const statusBox=document.getElementById("status");
document.getElementById("year").textContent=new Date().getFullYear();
function setStatus(message,kind=""){statusBox.textContent=message;statusBox.className="status"+(kind?" "+kind:"");}
function useFile(file){
 if(!file)return;
 if(!/\.pdf$/i.test(file.name)&&file.type!=="application/pdf"){fileInput.value="";fileInfo.hidden=true;setStatus("Please choose a PDF file.","error");return;}
 fileTitle.textContent=file.name;fileSubtitle.textContent="Ready to convert";fileInfo.hidden=false;fileInfo.textContent=(file.size/1024/1024).toFixed(2)+" MB";setStatus("");
}
fileInput.addEventListener("change",()=>useFile(fileInput.files[0]));
["dragenter","dragover"].forEach(name=>dropzone.addEventListener(name,e=>{e.preventDefault();dropzone.classList.add("dragover");}));
["dragleave","drop"].forEach(name=>dropzone.addEventListener(name,e=>{e.preventDefault();dropzone.classList.remove("dragover");}));
dropzone.addEventListener("drop",e=>{const file=e.dataTransfer.files[0];if(!file)return;const dt=new DataTransfer();dt.items.add(file);fileInput.files=dt.files;useFile(file);});
form.addEventListener("submit",async e=>{
 e.preventDefault();const file=fileInput.files[0];if(!file)return setStatus("Please select a PDF file first.","error");
 if(file.size>20*1024*1024)return setStatus("File is too large. Maximum size is 20 MB.","error");
 const body=new FormData();body.append("pdf",file);button.disabled=true;button.textContent="Converting…";setStatus("Please wait while NM PDF prepares your Excel file.");
 try{
  const response=await fetch("/api/pdf-to-excel",{method:"POST",body});const type=response.headers.get("content-type")||"";
  if(!response.ok||type.includes("application/json")){const data=await response.json().catch(()=>({}));throw new Error(data.error||"Conversion failed. Please try again.");}
  const blob=await response.blob();const disposition=response.headers.get("content-disposition")||"";const match=disposition.match(/filename="?([^"]+)"?/i);const filename=match?match[1]:"NM-PDF-converted.xlsx";
  const url=URL.createObjectURL(blob);const a=document.createElement("a");a.href=url;a.download=filename;document.body.appendChild(a);a.click();a.remove();URL.revokeObjectURL(url);setStatus("Done! Your Excel file has been downloaded.","success");
 }catch(error){setStatus(error.message||"Could not connect to the server. Is NM PDF running?","error");}
 finally{button.disabled=false;button.innerHTML='Convert to Excel <span>→</span>';}
});
