(function(){
 const links=[...document.querySelectorAll('#sideNav a')],panels=[...document.querySelectorAll('.panel')],main=document.getElementById('mainContent');
 function activate(id,hash=true){panels.forEach(p=>p.classList.toggle('active',p.id===id));links.forEach(a=>a.classList.toggle('active',a.dataset.target===id));if(main)main.scrollTop=0;if(hash)history.replaceState(null,'','#'+id)}
 links.forEach(a=>a.addEventListener('click',e=>{e.preventDefault();activate(a.dataset.target)}));
 document.querySelectorAll('[data-go]').forEach(b=>b.addEventListener('click',()=>activate(b.dataset.go)));
 const initial=location.hash.slice(1);if(initial&&document.getElementById(initial))activate(initial,false);
 const slides=[...document.querySelectorAll('.tutorial-slide')],dots=document.getElementById('tutorialDots'),prev=document.getElementById('tutorialPrev'),next=document.getElementById('tutorialNext'),counter=document.getElementById('tutorialCounter');let i=0;
 if(slides.length){slides.forEach((_,n)=>{const d=document.createElement('button');d.className='tutorial-dot';d.setAttribute('aria-label','Tutorial step '+(n+1));d.onclick=()=>show(n);dots.appendChild(d)});function show(n){i=Math.max(0,Math.min(n,slides.length-1));slides.forEach((s,x)=>s.classList.toggle('active',x===i));[...dots.children].forEach((d,x)=>d.classList.toggle('active',x===i));prev.disabled=i===0;next.disabled=i===slides.length-1;counter.textContent=`Step ${i+1} of ${slides.length}`}prev.onclick=()=>show(i-1);next.onclick=()=>show(i+1);show(0)}
})();
