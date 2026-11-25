const themeButton=document.getElementById('theme-toggle');
function paintTheme(){const dark=document.documentElement.dataset.theme==='dark';themeButton.textContent=dark?'Modo claro':'Modo oscuro';themeButton.setAttribute('aria-pressed',String(dark));}
themeButton.addEventListener('click',()=>{document.documentElement.dataset.theme=document.documentElement.dataset.theme==='dark'?'light':'dark';try{localStorage.setItem('pax-theme',document.documentElement.dataset.theme)}catch{}paintTheme()});paintTheme();
