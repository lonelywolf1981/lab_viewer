(function(){
  // Loader stub: real logic is split into /static/js/*.js for easier maintenance.
  // Keeps compatibility with existing index.html that loads /static/app.js.
  var v = (window.ASSET_V != null) ? String(window.ASSET_V) : "";
  function withV(url){
    if(!v) return url;
    return url + (url.indexOf("?")>=0 ? "&" : "?") + "v=" + encodeURIComponent(v);
  }
  var files = [
    "/static/js/core.js",
    "/static/js/channels.js",
    "/static/js/plot_export.js",
    "/static/js/main.js"
  ];
  function loadNext(i){
    if(i >= files.length) return;
    var s = document.createElement("script");
    s.src = withV(files[i]);
    s.async = false;
    s.onload = function(){ loadNext(i+1); };
    s.onerror = function(){
      try{ console.error("Failed to load", s.src); }catch(e){}
      loadNext(i+1);
    };
    (document.head || document.documentElement).appendChild(s);
  }
  loadNext(0);
})();
