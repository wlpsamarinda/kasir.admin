/* Shim google.script.run -> fetch ke Apps Script Web App (doPost). */
(function () {
  if (window.google && window.google.script && window.google.script.run) return; // berjalan di dalam Apps Script
  function call(fn, args, ok, fail) {
    var url = window.GAS_URL;
    if (!url || url.indexOf('PASTE_') === 0) {
      var e = new Error('GAS_URL belum diatur di config.js');
      if (fail) fail(e); else console.error(e);
      return;
    }
    fetch(url, {
      method: 'POST',
      redirect: 'follow',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ fn: fn, args: args })
    }).then(function (r) { return r.text(); })
      .then(function (t) {
        var j;
        try { j = JSON.parse(t); } catch (x) { throw new Error('Respons server tidak valid'); }
        if (!j.ok) throw new Error(j.error || 'Terjadi kesalahan server');
        if (ok) ok(j.result);
      })
      .catch(function (err) {
        if (!navigator.onLine) err = new Error('Tidak ada koneksi internet');
        if (fail) fail(err); else console.error(err);
      });
  }
  function make(ok, fail, userObj) {
    return new Proxy({}, {
      get: function (_, name) {
        if (name === 'withSuccessHandler') return function (f) { return make(function (r) { f(r, userObj); }, fail, userObj); };
        if (name === 'withFailureHandler') return function (f) { return make(ok, function (e) { f(e, userObj); }, userObj); };
        if (name === 'withUserObject') return function (u) { return make(ok, fail, u); };
        if (typeof name !== 'string' || name === 'then') return undefined;
        return function () { call(name, Array.prototype.slice.call(arguments), ok, fail); };
      }
    });
  }
  window.google = window.google || {};
  window.google.script = { run: make(null, null) };
})();
