(function () {
  const $ = (id) => document.getElementById(id);
  let state = null;

  const PW_PLACEHOLDER = "••••••••";
  const STATUS_TEXT = { idle: "Idle", connecting: "Connecting", connected: "Connected" };

  function fmt(s) {
    s = Math.max(0, Math.floor(s || 0));
    const h = Math.floor(s / 3600), m = Math.floor((s % 3600) / 60), x = s % 60;
    const p = (n) => String(n).padStart(2, "0");
    return h ? `${h}:${p(m)}:${p(x)}` : `${m}:${p(x)}`;
  }
  function api() { return (window.pywebview && window.pywebview.api) || null; }

  function renderState(s) {
    state = s;
    const np = !!(s.nowplaying && s.title);

    $("statuspill").className = "status-pill " + (s.status || "idle");
    $("statustext").textContent = STATUS_TEXT[s.status] || "Idle";

    const cb = $("connectbtn");
    cb.textContent = s.connected ? "Disconnect" : "Connect";
    cb.classList.toggle("connected", !!s.connected);

    $("nptitle").textContent = np ? s.title : "Nothing playing";
    $("npartist").textContent = np ? s.artist : "";
    $("nplabel").textContent = s.paused ? "Paused" : "Now playing";

    const img = $("artimg"), ph = document.querySelector(".art .ph");
    if (np && s.cover) {
      if (img.src !== s.cover) img.src = s.cover;
      img.hidden = false; ph.style.display = "none";
    } else {
      img.hidden = true; img.removeAttribute("src"); ph.style.display = "";
    }
    $("pausebadge").hidden = !s.paused;

    const qb = $("qbadge");
    if (np && s.quality) { $("qtext").textContent = s.quality; qb.hidden = false; }
    else qb.hidden = true;

    $("progress").hidden = !(np && s.dur);
    $("s-songs").textContent = s.songs || 0;
    $("s-listen").textContent = fmt(s.listened || 0);
    $("msg").textContent = s.msg || "";
  }

  function renderSettings(st) {
    $("f-discord").value = st.discord_app_id || "";
    $("f-email").value = st.qobuz_email || "";
    $("f-pw").value = st.has_password ? PW_PLACEHOLDER : "";
    setQuality(st.quality_label || "Hi-Res 24-Bit / 96 kHz");
    $("f-interval").value = st.update_interval || 3;
    setToggle("t-badge", st.show_quality_badge);
    setToggle("t-auto", st.auto_connect);
    setToggle("t-start", st.start_with_windows);
  }
  const setToggle = (id, on) => $(id).setAttribute("aria-checked", on ? "true" : "false");
  const getToggle = (id) => $(id).getAttribute("aria-checked") === "true";

  function setQuality(v) {
    $("f-quality").value = v;
    $("qlabel").textContent = v;
    document.querySelectorAll("#qmenu li").forEach((li) =>
      li.setAttribute("aria-selected", li.dataset.v === v ? "true" : "false"));
  }

  // Repaint the live progress + session clocks on a coarse interval (every 250ms)
  // instead of every animation frame. A 60fps rAF loop forced the glass cards
  // (backdrop-filter) to recomposite constantly, which is what made scrolling
  // chug. 4 updates/sec looks identical and leaves the GPU alone; a CSS
  // transition on the fill keeps the bar motion smooth between ticks.
  function paintProgress() {
    if (!state) return;
    if (state.nowplaying && state.dur) {
      let el = state.paused ? (state.pos || 0) : (Date.now() / 1000 - state.tstart);
      el = Math.max(0, Math.min(el, state.dur));
      $("fill").style.width = (el / state.dur * 100) + "%";
      $("t-pos").textContent = fmt(el);
      $("t-dur").textContent = fmt(state.dur);
    }
    if (state.connected && state.session_start) {
      $("s-session").textContent = fmt(Date.now() / 1000 - state.session_start);
    }
  }

  function wire() {
    $("btn-min").onclick = () => api() && api().minimize();
    $("btn-close").onclick = () => api() && api().close();
    $("connectbtn").onclick = () => api() && api().toggle_connect();
    document.querySelectorAll(".toggle").forEach((t) => {
      t.onclick = () => t.setAttribute("aria-checked", getToggle(t.id) ? "false" : "true");
    });

    // custom fallback-quality dropdown (native <select> popups can't be themed)
    const qsel = $("qsel"), qbtn = $("qbtn"), qmenu = $("qmenu");
    const closeQ = () => { qsel.classList.remove("open"); qmenu.hidden = true; qbtn.setAttribute("aria-expanded", "false"); };
    qbtn.onclick = (e) => {
      e.stopPropagation();
      const open = !qsel.classList.contains("open");
      qsel.classList.toggle("open", open); qmenu.hidden = !open; qbtn.setAttribute("aria-expanded", String(open));
    };
    qmenu.querySelectorAll("li").forEach((li) => { li.onclick = () => { setQuality(li.dataset.v); closeQ(); }; });
    document.addEventListener("click", (e) => { if (!qsel.contains(e.target)) closeQ(); });
    document.addEventListener("keydown", (e) => { if (e.key === "Escape") closeQ(); });

    // number stepper for the update interval
    const step = (d) => { const el = $("f-interval"); let v = parseInt(el.value, 10); if (isNaN(v)) v = 3; el.value = Math.max(1, v + d); };
    document.querySelector(".step-up").onclick = () => step(1);
    document.querySelector(".step-down").onclick = () => step(-1);

    $("savebtn").onclick = () => {
      const s = {
        discord_app_id: $("f-discord").value,
        qobuz_email: $("f-email").value,
        qobuz_password: $("f-pw").value,
        quality_label: $("f-quality").value,
        update_interval: $("f-interval").value,
        show_quality_badge: getToggle("t-badge"),
        auto_connect: getToggle("t-auto"),
        start_with_windows: getToggle("t-start"),
      };
      const a = api();
      if (a) a.save_settings(s).then((r) => {
        $("msg").textContent = "Settings saved";
        if (r && r.settings) renderSettings(r.settings);
      });
    };
  }

  window.qSetState = (s) => renderState(s);

  function loadFromBackend() {
    const a = api();
    if (!a) return;
    a.get_state().then((d) => { renderSettings(d.settings); renderState(d.state); });
  }

  function init() { wire(); setInterval(paintProgress, 250); }
  if (document.readyState !== "loading") init();
  else document.addEventListener("DOMContentLoaded", init);

  window.addEventListener("pywebviewready", loadFromBackend);
  if (window.pywebview && window.pywebview.api) loadFromBackend();
})();
