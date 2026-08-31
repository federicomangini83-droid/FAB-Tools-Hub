
function setStatus(message,type=''){const el=document.getElementById('status');el.textContent=message;el.className='status show '+type}
function bindDrop(inputId,dropId){const i=document.getElementById(inputId),d=document.getElementById(dropId);['dragenter','dragover'].forEach(e=>d.addEventListener(e,x=>{x.preventDefault();d.classList.add('drag')}));['dragleave','drop'].forEach(e=>d.addEventListener(e,x=>{x.preventDefault();d.classList.remove('drag')}));d.addEventListener('drop',e=>{i.files=e.dataTransfer.files;i.dispatchEvent(new Event('change'))})}
function safeName(s){return String(s||'output').replace(/[^a-zA-Z0-9._-]+/g,'_').replace(/^_+|_+$/g,'')||'output'}
