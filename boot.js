
(function(){
  const app = document.getElementById('app');
  const log = (msg) => {
    const div = document.createElement('div');
    div.style.cssText='max-width:800px;margin:20px auto;background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.08);padding:12px;border-radius:12px;color:#9db1ff;font:14px system-ui';
    div.textContent = msg;
    app.appendChild(div);
  };
  log('正在加载所需库… 如久无反应，可能是CDN被网络限制，可切换网络或稍后重试。');

  const libs = [
    {global:'React', urls:[
      'https://cdn.jsdelivr.net/npm/react@18/umd/react.production.min.js',
      'https://unpkg.com/react@18/umd/react.production.min.js'
    ]},
    {global:'ReactDOM', urls:[
      'https://cdn.jsdelivr.net/npm/react-dom@18/umd/react-dom.production.min.js',
      'https://unpkg.com/react-dom@18/umd/react-dom.production.min.js'
    ]},
    {global:'XLSX', urls:[
      'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
      'https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js'
    ]},
    {global:'html2canvas', urls:[
      'https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js',
      'https://unpkg.com/html2canvas@1.4.1/dist/html2canvas.min.js'
    ]}
  ];

  function loadOne(url){
    return new Promise((resolve,reject)=>{
      const s = document.createElement('script');
      s.src = url; s.async = false;
      s.onload = ()=> resolve(url);
      s.onerror = ()=> reject(new Error('加载失败：'+url));
      document.head.appendChild(s);
    });
  }

  async function ensureGlobal(globalName, urls){
    for(const u of urls){
      try{
        await loadOne(u);
        if(window[globalName]) return;
      }catch(e){
        console.warn(e);
      }
    }
    throw new Error('无法加载依赖：'+globalName);
  }

  (async()=>{
    try{
      for(const lib of libs){
        await ensureGlobal(lib.global, lib.urls);
      }
      // After libs, load the app main.js
      await loadOne('./main.js');
    }catch(e){
      log('❌ 依赖加载失败：'+e.message);
    }
  })();
})();
