/** Vista previa local: datos ficticios, mismo motor HomeLedger, sin conexión a Google. */
const http = require("node:http");
const fs = require("node:fs");
const path = require("node:path");
const { createHarness } = require("./harness.cjs");
const root = path.resolve(__dirname, "../dist");
const h = createHarness();
const [ana, luis, maria] = h.seed();
const participants = (people) => people.map((p) => ({ personId: p.id }));
const energy = h.execute("TEMPLATE", {
  name: "Energía eléctrica",
  provider: "ENEE",
  mode: "equal",
  participants: participants([ana, luis]),
}).state.templates[0];
const claro = h.execute("TEMPLATE", {
  name: "Cable e internet",
  provider: "Claro · Plan del hogar",
  mode: "equal",
  participants: participants([ana, luis, maria]),
}).state.templates[1];
h.execute("CHARGE", {
  templateId: energy.id,
  period: "2026-11",
  amount: "1248.50",
  date: "2026-11-01",
  due: "2026-11-20",
});
h.execute("CHARGE", {
  templateId: claro.id,
  period: "2026-11",
  amount: "900",
  date: "2026-11-01",
  due: "2026-11-15",
  payerId: ana.id,
});
h.execute("DEBT", {
  name: "Préstamo para compra del hogar",
  debtorId: ana.id,
  creditorId: luis.id,
  amount: "450",
  date: "2026-11-05",
  due: "2026-11-25",
});
h.execute("DEBT", {
  name: "Adelanto de transporte",
  debtorId: luis.id,
  creditorId: maria.id,
  amount: "200",
  date: "2026-11-03",
  due: "2026-11-12",
});
const obligation = h.context
  .homeGetState("token-admin")
  .obligations.find((o) => o.debtorId === maria.id && o.creditorId === ana.id);
h.execute(
  "PAYMENT",
  {
    obligationId: obligation.id,
    amount: "100",
    date: "2026-11-10",
    note: "Transferencia de ejemplo",
  },
  "maria",
);
function html() {
  let source = fs.readFileSync(path.join(root, "Frontend/Index.html"), "utf8");
  source = source.replace(/<\?!= include\('([^']+)'\); \?>/g, (_, file) =>
    fs.readFileSync(path.join(root, "Frontend", file + ".html"), "utf8"),
  );
  const bridge = `<style>.home-demo-banner{position:fixed;top:0;left:0;right:0;z-index:3000;background:#173c36;color:#fff;padding:8px 16px;font:12px system-ui;display:flex;justify-content:space-between;align-items:center;gap:8px}.home-demo-banner select{border-radius:6px;padding:4px}body{padding-top:44px}.wbs-sidebar{top:44px!important}dialog{z-index:4000}</style>
  <div class="home-demo-banner"><span>VISTA PREVIA LOCAL · datos ficticios · sin conexión a Google</span><label>Ver como <select id="demo-user"><option value="admin">Administrador</option><option value="ana">Ana</option><option value="luis">Luis</option><option value="maria">María</option></select></label></div>
  <script>
  const demoUsers=${JSON.stringify(h.users)};
  let selectedDemo=localStorage.getItem('HOME_DEMO_USER')||'admin';
  function chooseDemo(id){localStorage.setItem('HOME_DEMO_USER',id);localStorage.setItem('APP_TOKEN','token-'+id);localStorage.setItem('APP_USER',JSON.stringify(demoUsers.find(u=>u.id===id)));}
  chooseDemo(selectedDemo);document.getElementById('demo-user').value=selectedDemo;
  document.getElementById('demo-user').onchange=e=>{chooseDemo(e.target.value);location.reload();};
  window.google={script:{get run(){let success=()=>{},failure=()=>{};const chain=new Proxy({}, {get(_,name){if(name==='withSuccessHandler')return fn=>{success=fn;return chain;};if(name==='withFailureHandler')return fn=>{failure=fn;return chain;};return (...args)=>fetch('/rpc',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({name,args})}).then(r=>r.json()).then(r=>{if(r.error)failure({message:r.error});else success(r.result);}).catch(e=>failure({message:e.message}));}});return chain;}}};
  </script>`;
  return source.replace(
    '<div id="root"></div>',
    bridge + '<div id="root"></div>',
  );
}
const server = http.createServer(async (req, res) => {
  res.setHeader("Cache-Control", "no-store");
  if (req.method === "GET" && req.url === "/") {
    res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
    res.end(html());
    return;
  }
  if (req.method === "POST" && req.url === "/rpc") {
    let body = "";
    try {
      for await (const chunk of req) {
        body += chunk;
        if (body.length > 100000) throw Error("Solicitud demasiado grande");
      }
      const { name, args } = JSON.parse(body);
      let result;
      if (
        [
          "homeGetState",
          "homeExecute",
          "homeMigrationPreview",
          "validateToken",
        ].includes(name)
      )
        result = h.context[name](...args);
      else if (name === "listarUsuarios") {
        h.context.homeLegacyAuthorize_(args);
        result = h.users.map((u) => ({
          id: u.id,
          nombre_usuario: u.nombre_usuario,
          nombre_persona: u.nombreCompleto,
        }));
      } else if (name === "logout") result = { ok: true };
      else
        throw Error(
          "Este módulo histórico no está habilitado en la demostración local.",
        );
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify({ result }));
    } catch (e) {
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify({ error: e.message }));
    }
    return;
  }
  res.writeHead(404);
  res.end("No encontrado");
});
server.listen(4173, "127.0.0.1", () =>
  console.log("Home: http://127.0.0.1:4173 · datos ficticios en memoria"),
);
