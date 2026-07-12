(function(){
  const cfg = window.STORM_CONFIG || {};
  function normalize(raw){
    if(!raw) return window.STORM_FALLBACK;
    if(raw.ok === false) throw new Error(raw.error || 'Could not load Storm data.');
    return raw.data || raw;
  }
  window.StormAPI = {
    loadPublic(){
      const url=(cfg.apiUrl||'').trim();
      if(!url) return Promise.resolve(window.STORM_FALLBACK);
      return new Promise((resolve)=>{
        const cb='stormPublic_'+Date.now()+'_'+Math.random().toString(36).slice(2);
        const script=document.createElement('script');
        const timer=setTimeout(()=>finish(window.STORM_FALLBACK),9000);
        function finish(data){clearTimeout(timer); try{delete window[cb]}catch(e){}; script.remove(); resolve(data)}
        window[cb]=(payload)=>{try{finish(normalize(payload))}catch(e){console.error(e);finish(window.STORM_FALLBACK)}};
        script.src=url+(url.includes('?')?'&':'?')+'action=public&callback='+encodeURIComponent(cb)+'&_='+Date.now();
        script.onerror=()=>finish(window.STORM_FALLBACK);
        document.head.appendChild(script);
      });
    },
    appLink(page){
      const url=(cfg.apiUrl||'').trim();
      return url ? url+(url.includes('?')?'&':'?')+'page='+encodeURIComponent(page) : '';
    }
  };
})();
