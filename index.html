<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Validador de Reclassifica&ccedil;&otilde;es &middot; Valida&ccedil;&atilde;o de Reclassifica&ccedil;&otilde;es de PEP</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;600&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js"></script>
<style>
  :root{
    --blue:#0072CE; --blue-lt:#3aa0ff; --blue-dk:#005AAB;
    --green:#00E08A; --green-dk:#00A859;
    --r:18px; --r-sm:12px;
    --ink:#e8f1fb; --ink-soft:#a9bdd6; --ink-mute:#6f86a3;
    --glass:rgba(20,34,56,.55); --glass-2:rgba(13,24,42,.62);
    --stroke:rgba(120,170,230,.18); --stroke-2:rgba(120,170,230,.10);
    --ok:#00E08A; --warn:#ffc44d; --err:#ff6b81;
    --glow-b:0 0 0 1px rgba(0,114,206,.22),0 14px 40px rgba(0,80,160,.30);
    --glow-g:0 0 0 1px rgba(0,224,138,.25),0 14px 40px rgba(0,168,89,.26);
    --page:#060b16; --page-2:#040810;
    --au1:rgba(0,114,206,.28); --au2:rgba(0,224,138,.18); --au3:rgba(0,90,171,.30);
    --orb-op:.55; --grid-op:.05; --tbl-bg:rgba(8,17,32,.5); --tr-even:rgba(120,170,230,.03);
    --th-grad:linear-gradient(180deg,rgba(0,90,171,.55),rgba(0,60,120,.5)); --th-ink:#cfe2f7;
    --code-bg:rgba(0,114,206,.14); --logo-grad:linear-gradient(135deg,#fff,#eaf4ff);
    --h1-grad:linear-gradient(100deg,#fff 10%,var(--blue-lt) 45%,var(--green) 90%);
  }
  body.light{
    --ink:#13314e; --ink-soft:#42607f; --ink-mute:#7791ad;
    --glass:rgba(255,255,255,.78); --glass-2:rgba(255,255,255,.66);
    --stroke:rgba(0,90,171,.16); --stroke-2:rgba(0,90,171,.09);
    --ok:#00A859; --warn:#d98a16; --err:#e2566a;
    --glow-b:0 0 0 1px rgba(0,114,206,.18),0 14px 40px rgba(0,80,160,.16);
    --glow-g:0 0 0 1px rgba(0,168,89,.22),0 14px 40px rgba(0,168,89,.16);
    --page:#eaf2fb; --page-2:#dbe8f6;
    --au1:rgba(0,114,206,.20); --au2:rgba(0,224,138,.16); --au3:rgba(0,114,206,.16);
    --orb-op:.40; --grid-op:.05; --tbl-bg:rgba(255,255,255,.7); --tr-even:rgba(0,90,171,.035);
    --th-grad:linear-gradient(180deg,var(--blue),var(--blue-dk)); --th-ink:#fff;
    --code-bg:rgba(0,114,206,.10); --logo-grad:linear-gradient(135deg,var(--blue),var(--blue-dk));
    --h1-grad:linear-gradient(100deg,var(--blue-dk) 5%,var(--blue) 45%,var(--green-dk) 95%);
  }
  *{box-sizing:border-box}
  html{scroll-behavior:smooth}
  body{margin:0;font-family:"Manrope",system-ui,Segoe UI,Roboto,Arial,sans-serif;color:var(--ink);font-size:14px;
    background:var(--page);overflow-x:hidden;min-height:100vh;transition:color .4s,background .4s}
  code,.mono{font-family:"JetBrains Mono",ui-monospace,monospace}

  .aurora{position:fixed;inset:0;z-index:-2;overflow:hidden;transition:background .4s;background:
    radial-gradient(1200px 700px at 12% -10%,var(--au1),transparent 60%),
    radial-gradient(1000px 680px at 100% 0%,var(--au2),transparent 55%),
    radial-gradient(900px 700px at 50% 120%,var(--au3),transparent 60%),
    linear-gradient(180deg,var(--page) 0,var(--page-2) 100%)}
  .aurora:before,.aurora:after{content:"";position:absolute;width:60vmax;height:60vmax;border-radius:50%;
    filter:blur(80px);opacity:var(--orb-op);mix-blend-mode:screen;animation:float 22s ease-in-out infinite}
  .aurora:before{background:radial-gradient(circle,#0072CE,transparent 65%);top:-15vmax;left:-10vmax}
  .aurora:after{background:radial-gradient(circle,#00E08A,transparent 65%);bottom:-18vmax;right:-12vmax;animation-delay:-8s}
  @keyframes float{0%,100%{transform:translate(0,0) scale(1)}50%{transform:translate(6vmax,4vmax) scale(1.12)}}
  .grid-bg{position:fixed;inset:0;z-index:-1;pointer-events:none;
    background-image:linear-gradient(rgba(120,170,230,var(--grid-op)) 1px,transparent 1px),linear-gradient(90deg,rgba(120,170,230,var(--grid-op)) 1px,transparent 1px);
    background-size:46px 46px;mask-image:radial-gradient(circle at 50% 30%,#000 0,transparent 78%)}

  header{position:relative;padding:30px 28px 22px;border-bottom:1px solid var(--stroke)}
  .head-inner{max-width:1280px;margin:0 auto}
  .brand{display:flex;align-items:center;gap:18px;flex-wrap:wrap}
  .logo{height:56px;min-width:150px;padding:0 22px;border-radius:16px;display:flex;align-items:center;justify-content:center;
    background:var(--logo-grad);box-shadow:0 10px 30px rgba(0,0,0,.35),inset 0 1px 0 rgba(255,255,255,.6);position:relative;transition:background .4s}
  .logo span{font-weight:800;font-size:27px;letter-spacing:.5px;line-height:1;transition:.4s;
    background:linear-gradient(120deg,var(--blue-dk),var(--green-dk));-webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent}
  body.light .logo span{-webkit-text-fill-color:#fff;background:none;color:#fff}
  .titlewrap{flex:1}
  .titlewrap h1{margin:0;font-size:30px;font-weight:800;letter-spacing:-.3px;line-height:1.05;
    background:var(--h1-grad);-webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent;
    background-size:220% auto;animation:shine 7s linear infinite}
  @keyframes shine{to{background-position:220% center}}
  .titlewrap .sub{margin-top:4px;color:var(--ink-soft);font-size:13.5px;font-weight:500;display:flex;align-items:center;gap:8px}

  .theme-toggle{margin-left:auto;display:inline-flex;align-items:center;gap:0;background:var(--glass);border:1px solid var(--stroke);
    border-radius:30px;padding:4px;cursor:pointer;-webkit-backdrop-filter:blur(10px);backdrop-filter:blur(10px);position:relative;transition:.3s;flex-shrink:0}
  .theme-toggle:hover{border-color:var(--blue-lt)}
  .theme-toggle .opt{width:34px;height:30px;border-radius:24px;display:flex;align-items:center;justify-content:center;font-size:15px;
    position:relative;z-index:1;color:var(--ink-mute);transition:.3s}
  .theme-toggle .opt.on{color:#fff}
  .theme-toggle .knob{position:absolute;top:4px;left:4px;width:34px;height:30px;border-radius:24px;
    background:linear-gradient(135deg,var(--blue),var(--green-dk));box-shadow:0 4px 12px rgba(0,114,206,.5);transition:transform .35s cubic-bezier(.4,1.4,.5,1)}
  body.light .theme-toggle .knob{transform:translateX(34px);background:linear-gradient(135deg,var(--warn),#ffd98a)}

  .wrap{max-width:1280px;margin:0 auto;padding:26px 28px 90px}

  .glass{background:linear-gradient(180deg,var(--glass),var(--glass-2));border:1px solid var(--stroke);
    border-radius:var(--r);box-shadow:0 20px 50px rgba(0,0,0,.35),inset 0 1px 0 rgba(255,255,255,.05);
    -webkit-backdrop-filter:blur(14px);backdrop-filter:blur(14px);transition:background .4s,border-color .4s}
  body.light .glass{box-shadow:0 16px 42px rgba(60,90,130,.14),inset 0 1px 0 rgba(255,255,255,.7)}

  .instr{padding:20px 22px;margin-bottom:22px;position:relative;overflow:hidden}
  .instr:before{content:"";position:absolute;left:0;top:0;bottom:0;width:4px;background:linear-gradient(var(--blue),var(--green))}
  .instr h2{margin:0 0 9px;font-size:16px;font-weight:800;color:var(--ink);display:flex;align-items:center;gap:10px}
  .instr h2 .ico{width:26px;height:26px;border-radius:8px;display:inline-flex;align-items:center;justify-content:center;font-size:14px;font-weight:800;
    background:linear-gradient(135deg,var(--blue),var(--green));color:#04101f;box-shadow:0 6px 18px rgba(0,224,138,.35)}
  .instr p{margin:0 0 7px;color:var(--ink-soft);font-size:13px;line-height:1.65}
  .instr p:last-child{margin-bottom:0}

  .card{padding:20px}
  .card h3{margin:0 0 4px;font-size:14px;font-weight:700;color:var(--ink);display:flex;align-items:center;gap:9px}
  .card h3 .n{width:24px;height:24px;border-radius:7px;display:inline-flex;align-items:center;justify-content:center;font-size:12px;font-weight:800;
    background:rgba(0,114,206,.18);color:var(--blue);border:1px solid rgba(0,114,206,.35)}
  .card .hint{color:var(--ink-mute);font-size:12px;margin:0 0 14px;line-height:1.5}

  .drop{position:relative;border:1.5px dashed rgba(120,170,230,.3);border-radius:var(--r-sm);padding:30px 18px;text-align:center;cursor:pointer;
    transition:.22s cubic-bezier(.2,.8,.2,1);background:var(--tbl-bg);overflow:hidden}
  .drop:before{content:"";position:absolute;inset:0;background:radial-gradient(400px 120px at var(--mx,50%) var(--my,0%),rgba(0,114,206,.18),transparent 70%);opacity:0;transition:.25s}
  .drop:hover{border-color:var(--blue-lt);transform:translateY(-2px);box-shadow:var(--glow-b)}
  .drop:hover:before,.drop.drag:before{opacity:1}
  .drop.drag{border-color:var(--blue-lt);background:rgba(0,114,206,.10)}
  .drop .ic{width:46px;height:46px;margin:0 auto 10px;border-radius:14px;display:flex;align-items:center;justify-content:center;
    background:linear-gradient(135deg,rgba(0,114,206,.25),rgba(0,224,138,.18));border:1px solid var(--stroke);font-size:22px}
  .drop .big{font-size:13.5px;font-weight:700;color:var(--ink);position:relative}
  .drop .sub{color:var(--ink-mute);font-size:12px;margin-top:4px;position:relative}
  .drop.loaded{border-color:var(--green);border-style:solid;background:rgba(0,224,138,.08);box-shadow:var(--glow-g)}
  .drop.loaded .ic{background:linear-gradient(135deg,var(--green),var(--green-dk));color:#04101f}
  .filemeta{font-size:12px;color:var(--green);margin-top:10px;word-break:break-all;font-weight:600;position:relative}
  input[type=file]{display:none}

  details.cfg{margin:20px 0 4px}
  details.cfg summary{cursor:pointer;color:var(--ink-soft);font-size:13px;font-weight:700;padding:9px 0;list-style:none;display:flex;align-items:center;gap:8px}
  details.cfg summary::-webkit-details-marker{display:none}
  details.cfg summary:before{content:"\25B8";transition:.2s;color:var(--blue-lt)}
  details.cfg[open] summary:before{transform:rotate(90deg)}
  .cfg-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:16px;padding:14px 0 4px}
  @media(max-width:860px){.cfg-grid{grid-template-columns:1fr}}
  .cfg-box{border:1px solid var(--stroke);border-radius:var(--r-sm);padding:14px;background:var(--tbl-bg)}
  .cfg-box h4{margin:0 0 10px;font-size:12px;font-weight:800;text-transform:uppercase;letter-spacing:.6px;color:var(--blue-lt)}
  .fld{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:8px}
  .fld label{font-size:12.5px;color:var(--ink-soft);font-weight:600}
  .fld input{font-family:"JetBrains Mono",monospace;width:74px;text-align:center;background:var(--glass-2);color:var(--ink);
    border:1px solid var(--stroke);border-radius:8px;padding:6px 8px;font-size:13px;font-weight:600;text-transform:uppercase}
  .fld input:focus{outline:none;border-color:var(--blue-lt);box-shadow:0 0 0 3px rgba(0,114,206,.2)}

  .actions{display:flex;gap:13px;margin:24px 0 6px;align-items:center;flex-wrap:wrap}
  button{font-family:inherit;font-size:14px;font-weight:700;border:0;border-radius:12px;padding:13px 26px;cursor:pointer;
    transition:.2s cubic-bezier(.2,.8,.2,1);position:relative;overflow:hidden}
  button:active{transform:scale(.97)}
  .btn{color:#fff;background:linear-gradient(135deg,var(--blue),var(--blue-dk));box-shadow:0 10px 28px rgba(0,114,206,.4)}
  .btn:before{content:"";position:absolute;top:0;left:-120%;width:60%;height:100%;background:linear-gradient(110deg,transparent,rgba(255,255,255,.4),transparent);transition:.5s}
  .btn:hover:not(:disabled){transform:translateY(-2px);box-shadow:0 16px 38px rgba(0,114,206,.5)}
  .btn:hover:not(:disabled):before{left:130%}
  .btn:disabled{background:rgba(80,100,130,.35);color:var(--ink-mute);cursor:not-allowed;box-shadow:none}
  .btn.ghost{background:rgba(120,170,230,.08);border:1px solid var(--stroke);color:var(--ink);box-shadow:none}
  .btn.ghost:hover{background:rgba(120,170,230,.16);border-color:var(--blue-lt)}
  .btn.ok{background:linear-gradient(135deg,var(--green),var(--green-dk));color:#04101f;box-shadow:0 10px 28px rgba(0,224,138,.42)}
  .btn.ok:before{content:"";position:absolute;top:0;left:-120%;width:60%;height:100%;background:linear-gradient(110deg,transparent,rgba(255,255,255,.5),transparent);transition:.5s}
  .btn.ok:hover{transform:translateY(-2px);box-shadow:0 16px 40px rgba(0,224,138,.55)}
  .btn.ok:hover:before{left:130%}

  .msg{padding:12px 16px;border-radius:12px;margin:16px 0;font-size:13px;display:none;font-weight:500}
  .msg.show{display:block;animation:slideIn .35s ease}
  .msg.error{background:rgba(255,107,129,.12);color:#ffb3bf;border:1px solid rgba(255,107,129,.35)}
  @keyframes slideIn{from{opacity:0;transform:translateY(-6px)}to{opacity:1;transform:none}}

  #results{margin-top:8px}
  #results.reveal{animation:reveal .5s ease}
  @keyframes reveal{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:none}}

  .verdict{padding:20px 22px;margin:10px 0 20px;border-left:5px solid;border-radius:var(--r)}
  .verdict.ok{border-color:var(--green);background:rgba(0,224,138,.08)}
  .verdict.err{border-color:var(--err);background:rgba(255,107,129,.08)}
  .verdict h2{margin:0 0 6px;font-size:19px;font-weight:800}
  .verdict.ok h2{color:var(--green)} .verdict.err h2{color:var(--err)}
  .verdict p{margin:0;color:var(--ink-soft);font-size:13.5px;line-height:1.6}

  .stat-row{display:grid;grid-template-columns:repeat(6,1fr);gap:14px;margin:10px 0 20px}
  @media(max-width:1100px){.stat-row{grid-template-columns:repeat(3,1fr)}}
  @media(max-width:640px){.stat-row{grid-template-columns:repeat(2,1fr)}}
  .stat{padding:18px 18px 16px;position:relative;overflow:hidden}
  .stat:before{content:"";position:absolute;top:0;left:0;right:0;height:3px;background:linear-gradient(90deg,var(--blue),var(--blue-lt))}
  .stat .v{font-size:25px;font-weight:800;letter-spacing:-.5px;color:var(--ink);line-height:1.1}
  .stat .v.sm{font-size:18px;padding-right:40px;word-break:break-all}
  .stat .l{font-size:11.5px;color:var(--ink-mute);margin-top:4px;text-transform:uppercase;letter-spacing:.6px;font-weight:600}
  .stat .spark{position:absolute;right:14px;top:14px;width:34px;height:34px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:16px;opacity:.85}
  .stat.ok:before{background:linear-gradient(90deg,var(--green),#7cffc4)} .stat.ok .v{color:var(--green)} .stat.ok .spark{background:rgba(0,224,138,.14)}
  .stat.warn:before{background:linear-gradient(90deg,var(--warn),#ffe09a)} .stat.warn .v{color:var(--warn)} .stat.warn .spark{background:rgba(255,196,77,.14)}
  .stat.err:before{background:linear-gradient(90deg,var(--err),#ffa7b4)} .stat.err .v{color:var(--err)} .stat.err .spark{background:rgba(255,107,129,.14)}
  .stat .spark.b{background:rgba(0,114,206,.16)}

  .tabs{display:flex;gap:6px;margin:20px 0 0;flex-wrap:wrap}
  .tab{padding:10px 16px;cursor:pointer;color:var(--ink-soft);font-size:13px;font-weight:600;border-radius:11px;border:1px solid transparent;transition:.18s}
  .tab:hover{background:rgba(120,170,230,.08);color:var(--ink)}
  .tab.active{color:var(--ink);background:linear-gradient(135deg,rgba(0,114,206,.3),rgba(0,224,138,.18));border-color:var(--stroke);box-shadow:inset 0 1px 0 rgba(255,255,255,.08)}
  body.light .tab.active{background:linear-gradient(135deg,rgba(0,114,206,.16),rgba(0,168,89,.12))}
  .tab .cnt{display:inline-block;margin-left:7px;padding:1px 7px;border-radius:20px;font-size:10.5px;font-weight:800;background:rgba(120,170,230,.18)}
  .tab .cnt.err{background:rgba(255,107,129,.25);color:var(--err)}
  .tab .cnt.ok{background:rgba(0,224,138,.2);color:var(--green)}
  .panel-body{display:none;padding-top:16px;animation:fadeUp .35s ease}
  .panel-body.active{display:block}
  @keyframes fadeUp{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:none}}

  .filterbar{display:flex;gap:10px;align-items:center;margin-bottom:12px;flex-wrap:wrap}
  .filterbar input[type=search]{flex:1;min-width:200px;font-family:inherit;background:var(--tbl-bg);color:var(--ink);
    border:1px solid var(--stroke);border-radius:10px;padding:9px 13px;font-size:13px}
  .filterbar input[type=search]:focus{outline:none;border-color:var(--blue-lt);box-shadow:0 0 0 3px rgba(0,114,206,.2)}
  .filterbar label{display:flex;gap:8px;align-items:center;color:var(--ink-soft);font-size:12.5px;font-weight:600;cursor:pointer}

  .tbl-scroll{max-height:520px;overflow:auto;border:1px solid var(--stroke);border-radius:var(--r-sm);background:var(--tbl-bg)}
  table{width:100%;border-collapse:collapse;font-size:12.5px}
  th,td{padding:9px 12px;text-align:left;border-bottom:1px solid var(--stroke-2);white-space:nowrap}
  th{color:var(--th-ink);font-weight:700;position:sticky;top:0;z-index:2;font-size:11px;text-transform:uppercase;letter-spacing:.5px;
    background:var(--th-grad);-webkit-backdrop-filter:blur(6px);backdrop-filter:blur(6px)}
  tbody tr{transition:.12s}
  tbody tr:nth-child(even){background:var(--tr-even)}
  tbody tr:hover{background:rgba(0,114,206,.12)}
  tbody tr.row-err{background:rgba(255,107,129,.09)}
  tbody tr.row-warn{background:rgba(255,196,77,.07)}
  td.num,th.num{text-align:right;font-variant-numeric:tabular-nums;font-family:"JetBrains Mono",monospace}
  td.mono{font-family:"JetBrains Mono",monospace;color:var(--ink-soft)}
  .badge{display:inline-flex;align-items:center;gap:5px;padding:3px 10px;border-radius:30px;font-size:11px;font-weight:700;white-space:nowrap}
  .b-ok{background:rgba(0,224,138,.14);color:var(--green)}
  .b-warn{background:rgba(255,196,77,.14);color:var(--warn)}
  .b-err{background:rgba(255,107,129,.14);color:var(--err)}
  .bar{position:relative;height:6px;border-radius:4px;background:rgba(120,170,230,.16);min-width:70px;overflow:hidden}
  .bar i{position:absolute;left:0;top:0;bottom:0;border-radius:4px;background:linear-gradient(90deg,var(--blue),var(--green))}
  .bar i.full{background:linear-gradient(90deg,var(--warn),#ffe09a)}
  .bar i.over{background:linear-gradient(90deg,var(--err),#ffa7b4)}
  .alert{padding:12px 15px;border-radius:12px;margin-bottom:9px;font-size:13px;border-left:4px solid;font-weight:500;line-height:1.5}
  .alert.warn{background:rgba(255,196,77,.08);border-color:var(--warn);color:var(--warn)}
  .alert.err{background:rgba(255,107,129,.08);border-color:var(--err);color:var(--err)}
  .alert.ok{background:rgba(0,224,138,.08);border-color:var(--green);color:var(--green)}
  .alert b{color:var(--ink)}
  .alert .sub{display:block;margin-top:5px;color:var(--ink-soft);font-size:12.5px;font-weight:500}

  /* causa raiz */
  .b-casc{background:rgba(255,107,129,.16);color:var(--err)}
  .b-pc{background:rgba(255,107,129,.16);color:var(--err);border:1px dashed rgba(255,107,129,.5)}
  .b-cls{background:rgba(255,150,60,.16);color:#ff9a3c}
  body.light .b-cls{color:#c9661a}
  .b-pep{background:rgba(255,196,77,.14);color:var(--warn)}
  .b-inv{background:rgba(160,120,255,.16);color:#b9a0ff}
  body.light .b-inv{color:#6d4fd1}
  tbody tr.row-inv{background:rgba(160,120,255,.08)}

  /* diagnostico por PEP */
  .diag{margin-bottom:14px;padding:0;overflow:hidden;border-left:5px solid var(--err)}
  .diag.warn{border-left-color:var(--warn)} .diag.cls{border-left-color:#ff9a3c}
  .diag-h{display:flex;align-items:center;gap:14px;padding:14px 18px;flex-wrap:wrap;cursor:pointer}
  .diag-h:hover{background:rgba(120,170,230,.06)}
  .diag-h .pep{font-family:"JetBrains Mono",monospace;font-size:15px;font-weight:700;color:var(--ink)}
  .diag-h .kv{display:flex;flex-direction:column;min-width:120px}
  .diag-h .kv .k{font-size:10.5px;text-transform:uppercase;letter-spacing:.5px;color:var(--ink-mute);font-weight:700}
  .diag-h .kv .v{font-family:"JetBrains Mono",monospace;font-size:13px;font-weight:600;color:var(--ink)}
  .diag-h .kv .v.neg{color:var(--err)}
  .diag-h .caret{margin-left:auto;color:var(--ink-mute);transition:.2s;font-size:13px}
  .diag.open .diag-h .caret{transform:rotate(90deg)}
  .diag-b{display:none;border-top:1px solid var(--stroke-2);padding:12px 18px 16px}
  .diag.open .diag-b{display:block}
  .diag-b .why{font-size:13px;color:var(--ink-soft);margin:0 0 10px;line-height:1.55}
  .diag-b .why b{color:var(--ink)}
  .diag-b table{font-size:12px}
  .diag-b th{position:static}
  .diag-b .meta{margin-top:10px;font-size:12px;color:var(--ink-mute)}
  .diag-b .notas{margin:12px 0 0;padding-left:18px;font-size:12.5px;color:var(--ink-soft);line-height:1.6}
  .diag-b .notas li{margin-bottom:6px}
  .diag-b .notas .mono{color:var(--ink);font-weight:600}
  .legend{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-bottom:14px;font-size:12px;color:var(--ink-soft)}
  .legend .badge{font-size:10.5px}
  .filterbar select{font-family:inherit;background:var(--tbl-bg);color:var(--ink);border:1px solid var(--stroke);border-radius:10px;padding:9px 12px;font-size:13px}

  .hidden{display:none}
  code{background:var(--code-bg);color:var(--blue);padding:1.5px 7px;border-radius:6px;font-size:12px;border:1px solid rgba(0,114,206,.18)}
  body.light code{color:var(--blue-dk)}
  .neg{color:var(--err)}
  ::-webkit-scrollbar{width:10px;height:10px}
  ::-webkit-scrollbar-track{background:transparent}
  ::-webkit-scrollbar-thumb{background:rgba(120,170,230,.25);border-radius:10px;border:2px solid transparent;background-clip:content-box}
  ::-webkit-scrollbar-thumb:hover{background:rgba(120,170,230,.4);background-clip:content-box}
  .foot{text-align:center;color:var(--ink-mute);font-size:11.5px;margin-top:30px;letter-spacing:.3px}
</style>
</head>
<body>
<div class="aurora"></div>
<div class="grid-bg"></div>

<header>
  <div class="head-inner">
    <div class="brand">
      <div class="logo"><span>SABESP</span></div>
      <div class="titlewrap">
        <h1>Validador de Reclassifica&ccedil;&otilde;es</h1>
        <div class="sub">Valida&ccedil;&atilde;o de saldos nas reclassifica&ccedil;&otilde;es de PEP</div>
      </div>
      <div class="theme-toggle" id="theme-toggle" title="Alternar tema claro/escuro" role="button" aria-label="Alternar tema">
        <div class="knob"></div>
        <div class="opt on" data-th="dark">&#127769;</div>
        <div class="opt" data-th="light">&#9728;</div>
      </div>
    </div>
  </div>
</header>

<div class="wrap">
  <div class="instr glass">
    <h2><span class="ico">i</span> Instru&ccedil;&otilde;es</h2>
    <p>Carregue a planilha de an&aacute;lise (ex.: <code>AN&Aacute;LISE RECLASSIFICA&Ccedil;&Otilde;ES - SETEMBRO.xlsx</code>) contendo as guias <code>Reclassifica&ccedil;&otilde;es</code>, <code>SALDO DO PEP</code> e <code>CJI3</code>. A ferramenta confere duas premissas:</p>
    <p><b>Premissa 1 &mdash; Saldo do PEP:</b> a soma reclassificada de cada <b>PEP de origem</b> (coluna B) n&atilde;o pode superar o saldo cont&aacute;bil do PEP na guia <code>SALDO DO PEP</code> (PEP na coluna H, saldo na coluna K).</p>
    <p><b>Premissa 2 &mdash; Saldo da classe de custo:</b> a soma reclassificada de cada par <b>PEP de origem + classe</b> (coluna E) n&atilde;o pode superar o saldo daquela classe na guia <code>CJI3</code> (PEP na coluna C, classe na coluna O, valor na coluna U).</p>
    <p><b>Causa raiz:</b> quando o PEP estoura, as classes dele tamb&eacute;m estouram em cascata. A aba <b>Diagn&oacute;stico por PEP</b> separa o que &eacute; <b>CASCATA</b> (resolver o PEP resolve a classe), <b>PEP + CLASSE</b> (a classe estoura al&eacute;m do PEP), <b>S&Oacute; CLASSE</b> (o PEP tem saldo, mas n&atilde;o nesta classe) e <b>S&Oacute; PEP</b>.</p>
    <p><b>Dados inv&aacute;lidos:</b> linhas com valor em branco, zero, n&atilde;o num&eacute;rico ou <b>negativo</b> s&atilde;o bloqueadas e <u>exclu&iacute;das das somas</u>. PEP fora da m&aacute;scara gera aviso. As compara&ccedil;&otilde;es s&atilde;o arredondadas a centavos antes de aplicar a toler&acirc;ncia.</p>
    <p>As letras de coluna e a toler&acirc;ncia s&atilde;o ajust&aacute;veis em <b>Configura&ccedil;&atilde;o de colunas</b>. Nada &eacute; deduzido automaticamente. O bot&atilde;o de download gera um Excel de confer&ecirc;ncia com todas as abas do resultado.</p>
  </div>

  <div class="card glass">
    <h3><span class="n">1</span> Planilha de an&aacute;lise</h3>
    <p class="hint">Arquivo &uacute;nico com as tr&ecirc;s guias. O processamento &eacute; 100% local no navegador &mdash; nenhum dado sai da sua m&aacute;quina.</p>
    <div class="drop" id="drop-x">
      <div class="ic">&#128202;</div>
      <div class="big">Clique ou arraste a planilha aqui</div>
      <div class="sub">.xlsx / .xlsm / .xls</div>
      <div class="filemeta hidden" id="meta-x"></div>
    </div>
    <input type="file" id="file-x" accept=".xlsx,.xlsm,.xls">

    <details class="cfg">
      <summary>Configura&ccedil;&atilde;o de colunas e toler&acirc;ncia</summary>
      <div class="cfg-grid">
        <div class="cfg-box">
          <h4>Guia Reclassifica&ccedil;&otilde;es</h4>
          <div class="fld"><label>PEP origem (DE)</label><input id="c-origem" value="B" maxlength="3"></div>
          <div class="fld"><label>PEP destino (PARA)</label><input id="c-destino" value="C" maxlength="3"></div>
          <div class="fld"><label>Classe de custo</label><input id="c-classe" value="E" maxlength="3"></div>
          <div class="fld"><label>Valor</label><input id="c-valor" value="H" maxlength="3"></div>
          <div class="fld"><label>Unidade</label><input id="c-unid" value="A" maxlength="3"></div>
          <div class="fld"><label>Aprovador</label><input id="c-aprov" value="N" maxlength="3"></div>
          <div class="fld"><label>Linha do cabe&ccedil;alho</label><input id="c-hdr-recl" value="1" maxlength="3"></div>
          <div class="fld"><label title="A = letra, 9 = digito; demais caracteres literais. Vazio = nao validar.">M&aacute;scara do PEP</label><input id="c-mask" value="AA9999A9.9999/A-999" maxlength="30" style="width:170px;text-transform:none"></div>
        </div>
        <div class="cfg-box">
          <h4>Guia SALDO DO PEP</h4>
          <div class="fld"><label>Elemento PEP</label><input id="s-pep" value="H" maxlength="3"></div>
          <div class="fld"><label>Valor cont&aacute;bil</label><input id="s-valor" value="K" maxlength="3"></div>
          <div class="fld"><label>Linha do cabe&ccedil;alho</label><input id="c-hdr-saldo" value="1" maxlength="3"></div>
        </div>
        <div class="cfg-box">
          <h4>Guia CJI3 / Toler&acirc;ncia</h4>
          <div class="fld"><label>Elemento PEP</label><input id="j-pep" value="C" maxlength="3"></div>
          <div class="fld"><label>Classe de custo</label><input id="j-classe" value="O" maxlength="3"></div>
          <div class="fld"><label>Valor/moeda objeto</label><input id="j-valor" value="U" maxlength="3"></div>
          <div class="fld"><label>Linha do cabe&ccedil;alho</label><input id="c-hdr-cji3" value="1" maxlength="3"></div>
          <div class="fld"><label>Toler&acirc;ncia (R$)</label><input id="c-tol" value="0,01" maxlength="8"></div>
        </div>
      </div>
    </details>
  </div>

  <div class="actions">
    <button class="btn" id="btn-run" disabled>&#9889; Validar</button>
    <button class="btn ghost" id="btn-reset">Limpar</button>
    <button class="btn ok hidden" id="btn-download">&#11015; Baixar confer&ecirc;ncia</button>
  </div>

  <div class="msg" id="msg"></div>

  <div id="results" class="hidden">
    <div class="verdict glass" id="verdict">
      <h2 id="v-title"></h2>
      <p id="v-text"></p>
    </div>

    <div class="stat-row">
      <div class="stat"><div class="spark b">&#128203;</div><div class="v" id="s-linhas">0</div><div class="l">Linhas</div></div>
      <div class="stat"><div class="spark b">&#128176;</div><div class="v sm" id="s-total">0</div><div class="l">Total reclassificado (R$)</div></div>
      <div class="stat"><div class="spark b">&#128273;</div><div class="v" id="s-peps">0</div><div class="l">PEPs de origem</div></div>
      <div class="stat" id="st-estpep"><div class="spark">&#128165;</div><div class="v" id="s-estpep">0</div><div class="l">PEPs que estouram saldo</div></div>
      <div class="stat" id="st-estcls"><div class="spark">&#127991;</div><div class="v" id="s-estcls">0</div><div class="l">Classes &mdash; estouro isolado</div></div>
      <div class="stat" id="st-estinv"><div class="spark">&#9888;</div><div class="v" id="s-estinv">0</div><div class="l">Linhas com dados inv&aacute;lidos</div></div>
    </div>

    <div class="tabs">
      <div class="tab active" data-t="dg">Diagn&oacute;stico por PEP<span class="cnt" id="n-dg">0</span></div>
      <div class="tab" data-t="ap">Apontamentos<span class="cnt" id="n-ap">0</span></div>
      <div class="tab" data-t="p1">Premissa 1 &middot; Saldo do PEP<span class="cnt" id="n-p1">0</span></div>
      <div class="tab" data-t="p2">Premissa 2 &middot; Saldo da classe<span class="cnt" id="n-p2">0</span></div>
      <div class="tab" data-t="lin">Detalhe por linha<span class="cnt" id="n-lin">0</span></div>
      <div class="tab" data-t="qa">Qualidade dos dados<span class="cnt" id="n-qa">0</span></div>
    </div>

    <div class="panel-body active" id="pb-dg">
      <div class="legend">
        <span>Legenda:</span>
        <span class="badge b-casc">CASCATA</span><span>PEP estoura e arrasta a classe &mdash; resolver o PEP resolve a classe</span>
        <span class="badge b-pc">PEP + CLASSE</span><span>a classe estoura al&eacute;m do excedente do PEP</span>
        <span class="badge b-cls">S&Oacute; CLASSE</span><span>PEP tem saldo, mas n&atilde;o nesta classe</span>
        <span class="badge b-pep">S&Oacute; PEP</span><span>PEP estoura, mas esta classe cabe</span>
      </div>
      <div class="filterbar">
        <input type="search" id="f-dg" placeholder="Filtrar por PEP, aprovador ou unidade...">
        <select id="o-dg">
          <option value="">Todas as causas</option>
          <option value="PEP">Estouro do PEP (cascata / PEP + classe / s&oacute; PEP)</option>
          <option value="CLS">S&oacute; classe</option>
        </select>
        <button class="btn ghost" id="dg-toggle" style="padding:8px 14px;font-size:12.5px">Expandir todos</button>
      </div>
      <div id="dg-list"></div>
    </div>

    <div class="panel-body" id="pb-ap"><div id="ap-list"></div></div>

    <div class="panel-body" id="pb-p1">
      <div class="filterbar">
        <input type="search" id="f-p1" placeholder="Filtrar por PEP...">
        <label><input type="checkbox" id="o-p1"> Mostrar s&oacute; problemas</label>
      </div>
      <div class="tbl-scroll"><table id="t-p1"></table></div>
    </div>

    <div class="panel-body" id="pb-p2">
      <div class="filterbar">
        <input type="search" id="f-p2" placeholder="Filtrar por PEP ou classe...">
        <label><input type="checkbox" id="o-p2"> Mostrar s&oacute; problemas</label>
      </div>
      <div class="tbl-scroll"><table id="t-p2"></table></div>
    </div>

    <div class="panel-body" id="pb-lin">
      <div class="filterbar">
        <input type="search" id="f-lin" placeholder="Filtrar por PEP, classe ou linha...">
        <label><input type="checkbox" id="o-lin"> Mostrar s&oacute; problemas</label>
      </div>
      <div class="tbl-scroll"><table id="t-lin"></table></div>
    </div>

    <div class="panel-body" id="pb-qa"><div id="qa-list"></div></div>
  </div>

  <div class="foot">SIGRA Lite &middot; Ger&ecirc;ncia de Conformidade Regulat&oacute;ria</div>
</div>

<script>
var state = { wb:null, name:null, result:null };

/* ---------- utilidades ---------- */
function norm(s){
  if(s==null) return "";
  var t = String(s).normalize("NFD").replace(/[\u0300-\u036f]/g,"");
  return t.replace(/[^a-z0-9]/gi,"").toUpperCase();
}
function esc(s){
  if(s==null) return "";
  return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
}
function fmtBRL(n){ return (n||0).toLocaleString("pt-BR",{minimumFractionDigits:2,maximumFractionDigits:2}); }

// Numero BR ou US: "30.732,52" -> 30732.52 ; 5896.6690 -> 5896.6690
function toNum(v){
  if(v==null||v==="") return 0;
  if(typeof v==="number") return v;
  var s=String(v).trim().replace(/\s/g,"").replace(/R\$/i,"");
  if(s==="") return 0;
  var neg=false;
  if(/-$/.test(s)){ neg=true; s=s.replace(/-$/,""); }          // sinal ao final (padrao SAP)
  if(s.indexOf(",")>-1 && s.indexOf(".")>-1) s=s.replace(/\./g,"").replace(",",".");
  else if(s.indexOf(",")>-1) s=s.replace(",",".");
  var n=parseFloat(s);
  if(isNaN(n)) return NaN;
  return neg?-n:n;
}
// PEP e classe entram como texto normalizado (sem espacos, sem ".0" do Excel)
function txt(v){
  if(v==null) return "";
  var s=String(v).trim();
  if(/^\d+\.0$/.test(s)) s=s.slice(0,-2);
  return s;
}
function keyPep(v){ return txt(v).toUpperCase(); }
// arredonda a centavos antes de comparar: evita falso "EXCEDE" por acumulo de ponto flutuante
function r2(n){ return Math.round((n+Number.EPSILON)*100)/100; }
// mascara "AA9999A9.9999/A-999": A = letra, 9 = digito, demais literais
function maskRegex(mask){
  var m=String(mask||"").trim(); if(!m) return null;
  var re="";
  for(var i=0;i<m.length;i++){
    var ch=m[i];
    if(ch==="A") re+="[A-Z]"; else if(ch==="9") re+="[0-9]"; else re+="\\"+ch;
  }
  try{ return new RegExp("^"+re+"$"); }catch(e){ return null; }
}

function colIdx(letter){
  var s=String(letter||"").trim().toUpperCase().replace(/[^A-Z]/g,"");
  if(!s) return -1;
  var n=0;
  for(var i=0;i<s.length;i++) n = n*26 + (s.charCodeAt(i)-64);
  return n-1;
}
function cfg(id){ return document.getElementById(id).value; }
function cfgInt(id,dflt){ var n=parseInt(cfg(id),10); return isNaN(n)?dflt:n; }

function findSheet(wb, candidates){
  for(var j=0;j<candidates.length;j++){
    var target=norm(candidates[j]);
    for(var i=0;i<wb.SheetNames.length;i++){ if(norm(wb.SheetNames[i])===target) return wb.SheetNames[i]; }
  }
  for(var j2=0;j2<candidates.length;j2++){
    var t2=norm(candidates[j2]);
    for(var i2=0;i2<wb.SheetNames.length;i2++){ if(norm(wb.SheetNames[i2]).indexOf(t2)>-1) return wb.SheetNames[i2]; }
  }
  return null;
}
function aoa(wb,name){
  return XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,raw:true,defval:null,blankrows:false});
}

/* ---------- nucleo da validacao ---------- */
function processar(){
  var wb=state.wb;
  var TOL=Math.abs(toNum(cfg("c-tol"))) || 0;

  var nRecl  = findSheet(wb,["Reclassificacoes","Reclassificacao","Reclass"]);
  var nSaldo = findSheet(wb,["SALDO DO PEP","Saldo PEP","Saldo"]);
  var nCji3  = findSheet(wb,["CJI3","CJI 3"]);
  if(!nRecl)  throw new Error("Nao encontrei a guia 'Reclassificacoes'. Guias do arquivo: "+wb.SheetNames.join(", "));
  if(!nSaldo) throw new Error("Nao encontrei a guia 'SALDO DO PEP'. Guias do arquivo: "+wb.SheetNames.join(", "));
  if(!nCji3)  throw new Error("Nao encontrei a guia 'CJI3'. Guias do arquivo: "+wb.SheetNames.join(", "));

  var cOri=colIdx(cfg("c-origem")), cDes=colIdx(cfg("c-destino")), cCls=colIdx(cfg("c-classe")), cVal=colIdx(cfg("c-valor"));
  var cUni=colIdx(cfg("c-unid")), cApr=colIdx(cfg("c-aprov"));      // opcionais: -1 desliga
  var sPep=colIdx(cfg("s-pep")),    sVal=colIdx(cfg("s-valor"));
  var jPep=colIdx(cfg("j-pep")),    jCls=colIdx(cfg("j-classe")), jVal=colIdx(cfg("j-valor"));
  if([cOri,cDes,cCls,cVal,sPep,sVal,jPep,jCls,jVal].some(function(x){return x<0;}))
    throw new Error("Configuracao de colunas invalida: use letras (ex.: B, C, E, H).");
  var MASK=maskRegex(cfg("c-mask"));

  var hR=cfgInt("c-hdr-recl",1), hS=cfgInt("c-hdr-saldo",1), hJ=cfgInt("c-hdr-cji3",1);

  /* --- 1) guia Reclassificacoes --- */
  var rowsR=aoa(wb,nRecl);
  var recl=[], qa=[], totalGeral=0, nTexto=0, nEspaco=0, nInval=0, nMask=0, totalInval=0;
  for(var i=hR;i<rowsR.length;i++){
    var r=rowsR[i]||[];
    var rawOri=r[cOri], rawDes=r[cDes], rawVal=r[cVal];
    var origem=keyPep(rawOri), destino=keyPep(rawDes), classe=txt(r[cCls]);
    var unidade=cUni>=0?txt(r[cUni]):"", aprovador=cApr>=0?txt(r[cApr]):"";
    var valBranco=(rawVal==null||String(rawVal).trim()==="");
    if(!origem && !destino && !classe && valBranco) continue;         // linha em branco
    var linha=i+1;                                                      // numero da linha no Excel
    var valor=valBranco?NaN:toNum(rawVal);
    var motivo=[];                                                      // motivos de bloqueio da linha
    if(valBranco){ motivo.push("VALOR EM BRANCO"); qa.push({lvl:"err",linha:linha,msg:"Valor em branco na coluna "+cfg("c-valor").toUpperCase()+". Linha bloqueada e excluida das somas."}); }
    else if(isNaN(valor)){ motivo.push("VALOR NAO NUMERICO"); qa.push({lvl:"err",linha:linha,msg:"Valor nao numerico na coluna "+cfg("c-valor").toUpperCase()+": "+JSON.stringify(rawVal)+". Linha bloqueada e excluida das somas."}); }
    else {
      if(typeof rawVal==="string") nTexto++;
      if(valor<0){ motivo.push("VALOR NEGATIVO"); qa.push({lvl:"err",linha:linha,msg:"Valor negativo ("+fmtBRL(valor)+"). Reclassificacao nao pode ser negativa; se a intencao e o sentido inverso, troque DE/PARA. Linha bloqueada e excluida das somas."}); }
      else if(r2(valor)===0){ motivo.push("VALOR ZERO"); qa.push({lvl:"err",linha:linha,msg:"Valor zero. Linha bloqueada e excluida das somas."}); }
    }
    if(rawOri!=null && String(rawOri)!==String(rawOri).trim()) nEspaco++;
    if(!origem){ motivo.push("PEP ORIGEM VAZIO"); qa.push({lvl:"err",linha:linha,msg:"PEP de origem vazio. Linha bloqueada."}); }
    if(!classe){ motivo.push("CLASSE VAZIA"); qa.push({lvl:"err",linha:linha,msg:"Classe de custo vazia. Linha bloqueada."}); }
    if(!destino) qa.push({lvl:"warn",linha:linha,msg:"PEP de destino vazio."});
    if(origem && destino && origem===destino){ motivo.push("DE = PARA"); qa.push({lvl:"err",linha:linha,msg:"PEP de origem igual ao de destino ("+origem+"). Linha bloqueada."}); }
    var foraMask=false;
    if(MASK){
      if(origem && !MASK.test(origem)){ foraMask=true; nMask++; qa.push({lvl:"warn",linha:linha,msg:"PEP de origem fora da mascara "+cfg("c-mask")+": \""+origem+"\" (verifique letra O x zero, espacos, caracteres invisiveis)."}); }
      if(destino && !MASK.test(destino)){ nMask++; qa.push({lvl:"warn",linha:linha,msg:"PEP de destino fora da mascara "+cfg("c-mask")+": \""+destino+"\"."}); }
    }
    var inval=motivo.length>0;
    if(inval){ nInval++; if(!isNaN(valor)) totalInval+=valor; valor=isNaN(valor)?0:valor; }
    else totalGeral+=valor;
    recl.push({linha:linha,origem:origem,destino:destino,classe:classe,valor:valor,unidade:unidade,aprovador:aprovador,
               inval:inval,motivo:motivo.join(" + "),foraMask:foraMask});
  }
  if(!recl.length) throw new Error("Nenhuma linha de reclassificacao encontrada na guia '"+nRecl+"'.");
  if(nTexto) qa.push({lvl:"warn",linha:null,msg:nTexto+" valor(es) da coluna "+cfg("c-valor").toUpperCase()+" estao gravados como TEXTO (nao somam no Excel). A validacao converteu corretamente."});
  if(nEspaco) qa.push({lvl:"warn",linha:null,msg:nEspaco+" PEP(s) de origem com espaco sobrando no inicio/fim. A validacao aplicou trim."});

  /* --- 2) guia SALDO DO PEP --- */
  var rowsS=aoa(wb,nSaldo);
  var saldoPep={}, nSaldoLin=0, dupPep={};
  for(var s=hS;s<rowsS.length;s++){
    var rs=rowsS[s]||[];
    var p=keyPep(rs[sPep]); if(!p) continue;
    var v=toNum(rs[sVal]); if(isNaN(v)) v=0;
    if(Object.prototype.hasOwnProperty.call(saldoPep,p)) dupPep[p]=(dupPep[p]||1)+1;
    saldoPep[p]=(saldoPep[p]||0)+v;
    nSaldoLin++;
  }
  var dupList=Object.keys(dupPep);
  if(dupList.length) qa.push({lvl:"warn",linha:null,msg:"Guia SALDO DO PEP: "+dupList.length+" PEP(s) em mais de uma linha (saldos somados): "+dupList.slice(0,10).join(", ")+(dupList.length>10?"...":"")+"."});

  /* --- 3) guia CJI3 --- */
  var rowsJ=aoa(wb,nCji3);
  var saldoCls={}, clsPorPep={}, pepCji3={}, nCjiLin=0;
  for(var j=hJ;j<rowsJ.length;j++){
    var rj=rowsJ[j]||[];
    var p2=keyPep(rj[jPep]); if(!p2) continue;
    var cl=txt(rj[jCls]);
    var v2=toNum(rj[jVal]); if(isNaN(v2)) v2=0;
    var k=p2+"\u0000"+cl;
    saldoCls[k]=(saldoCls[k]||0)+v2;
    pepCji3[p2]=true;
    if(!clsPorPep[p2]) clsPorPep[p2]={};
    clsPorPep[p2][cl]=(clsPorPep[p2][cl]||0)+v2;
    nCjiLin++;
  }

  /* --- 4) agregacoes --- */
  var totPep={}, totPC={}, linhasPC={}, aprPep={}, uniPep={}, maskPep={};
  for(var a=0;a<recl.length;a++){
    var x=recl[a];
    if(x.foraMask) maskPep[x.origem]=true;
    if(x.inval) continue;                                               // linha bloqueada nao consome saldo
    totPep[x.origem]=(totPep[x.origem]||0)+x.valor;
    var kk=x.origem+"\u0000"+x.classe;
    totPC[kk]=(totPC[kk]||0)+x.valor;
    if(!linhasPC[kk]) linhasPC[kk]=[];
    linhasPC[kk].push(x.linha);
    if(x.aprovador){ if(!aprPep[x.origem]) aprPep[x.origem]={}; aprPep[x.origem][x.aprovador]=true; }
    if(x.unidade){ if(!uniPep[x.origem]) uniPep[x.origem]={}; uniPep[x.origem][x.unidade]=true; }
  }

  /* --- 5) Premissa 1 --- */
  var p1=[], p1Falhas=0, mapP1={};
  for(var p in totPep){
    var usado=r2(totPep[p]);
    var temSaldo=Object.prototype.hasOwnProperty.call(saldoPep,p);
    var disp=temSaldo?r2(saldoPep[p]):0;
    var exc=r2(usado-disp);                                             // excedente arredondado a centavos
    var st="OK", lvl="ok";
    if(!temSaldo){ st="PEP SEM SALDO"; lvl="err"; }
    else if(exc>TOL){ st="EXCEDE SALDO"; lvl="err"; }
    else if(disp>0 && Math.abs(exc)<=TOL){ st="ZERA O PEP"; lvl="warn"; }
    if(lvl==="err") p1Falhas++;
    var row1={pep:p,usado:usado,disp:disp,folga:r2(disp-usado),exc:lvl==="err"?Math.max(exc,0):0,
              pct:disp>0?(usado/disp*100):(usado>0?Infinity:0),status:st,lvl:lvl,
              aprovadores:Object.keys(aprPep[p]||{}).sort(),unidades:Object.keys(uniPep[p]||{}).sort(),
              foraMask:!!maskPep[p],nCls:0,nClsErr:0};
    p1.push(row1); mapP1[p]=row1;
  }
  p1.sort(function(a,b){ if(a.lvl!==b.lvl) return a.lvl==="err"?-1:(b.lvl==="err"?1:(a.lvl==="warn"?-1:1)); return b.exc-a.exc||b.usado-a.usado; });

  /* --- 6) Premissa 2 + causa raiz --- */
  var CAUSA={CASC:"CASCATA",PC:"PEP + CLASSE",CLS:"SO CLASSE",PEP:"SO PEP"};
  var p2=[], p2Falhas=0, nCasc=0, nPC=0, nSoCls=0, nSoPep=0;
  for(var k2 in totPC){
    var parts=k2.split("\u0000"), pep=parts[0], cls=parts[1];
    var usado2=r2(totPC[k2]);
    var temCls=Object.prototype.hasOwnProperty.call(saldoCls,k2);
    var disp2=temCls?r2(saldoCls[k2]):0;
    var exc2=r2(usado2-disp2);
    var st2="OK", lvl2="ok", obs="";
    if(!temCls){
      st2="CLASSE INEXISTENTE"; lvl2="err";
      obs = pepCji3[pep] ? "O PEP existe na CJI3, mas nao ha lancamento nesta classe."
                         : "O PEP nao possui nenhum lancamento na CJI3.";
    } else if(exc2>TOL){
      st2="EXCEDE CLASSE"; lvl2="err";
      obs="Excedente de R$ "+fmtBRL(exc2)+"."+(disp2<=0?" Saldo da classe na CJI3 e zero/negativo.":"");
    } else if(disp2>0 && Math.abs(exc2)<=TOL){
      st2="ZERA A CLASSE"; lvl2="warn";
    }
    var a1=mapP1[pep]; a1.nCls++;
    var causa="", causaKey="", extra=0;
    if(lvl2==="err"){
      p2Falhas++; a1.nClsErr++;
      if(a1.lvl==="err"){
        // PEP tambem estoura: a classe foi arrastada ou estoura alem do PEP?
        if(a1.status==="PEP SEM SALDO" || exc2<=a1.exc+TOL){ causaKey="CASC"; nCasc++; }
        else { causaKey="PC"; nPC++; extra=r2(exc2-a1.exc); }
      } else { causaKey="CLS"; nSoCls++; }
      var alt=[];
      var mapa=clsPorPep[pep]||{};
      for(var cc in mapa){ if(cc!==cls && mapa[cc]>usado2-TOL) alt.push(cc+" (R$ "+fmtBRL(mapa[cc])+")"); }
      if(alt.length) obs+=" Classes do mesmo PEP com saldo suficiente: "+alt.sort().join("; ")+".";
    } else if(a1.lvl==="err"){ causaKey="PEP"; nSoPep++; }
    causa=causaKey?CAUSA[causaKey]:"";
    var causaTxt="";
    if(causaKey==="CASC") causaTxt= a1.status==="PEP SEM SALDO" ? "PEP ausente na guia SALDO DO PEP; a classe falha por consequencia."
                                   : "Excedente da classe (R$ "+fmtBRL(exc2)+") cabe dentro do excedente do PEP (R$ "+fmtBRL(a1.exc)+"). Reduzir o PEP em R$ "+fmtBRL(a1.exc)+" resolve as duas premissas.";
    else if(causaKey==="PC") causaTxt="A classe estoura R$ "+fmtBRL(extra)+" ALEM do excedente do PEP (R$ "+fmtBRL(a1.exc)+"). Mesmo ajustando o PEP, falta saldo nesta classe.";
    else if(causaKey==="CLS") causaTxt="O PEP tem saldo (folga R$ "+fmtBRL(a1.folga)+"), mas esta classe nao comporta R$ "+fmtBRL(usado2)+". Reduzir R$ "+fmtBRL(Math.max(exc2,0))+" nesta classe ou usar outra classe.";
    else if(causaKey==="PEP") causaTxt="Esta classe cabe no saldo CJI3, mas o PEP como um todo estoura; a linha fica bloqueada pelo PEP.";
    p2.push({pep:pep,classe:cls,usado:usado2,disp:disp2,folga:r2(disp2-usado2),exc:lvl2==="err"?Math.max(exc2,0):0,
             status:st2,lvl:lvl2,obs:obs,linhas:linhasPC[k2],causa:causa,causaKey:causaKey,causaTxt:causaTxt,extra:extra});
  }
  p2.sort(function(a,b){ if(a.lvl!==b.lvl) return a.lvl==="err"?-1:(b.lvl==="err"?1:(a.lvl==="warn"?-1:1)); return b.exc-a.exc||b.usado-a.usado; });

  /* --- 7) status por linha --- */
  var mapP2={}; for(var i2=0;i2<p2.length;i2++) mapP2[p2[i2].pep+"\u0000"+p2[i2].classe]=p2[i2];
  var lin=[], linFalhas=0;
  for(var b=0;b<recl.length;b++){
    var y=recl[b];
    var a1l=mapP1[y.origem], a2l=mapP2[y.origem+"\u0000"+y.classe];
    var lvl3="ok", st3="OK", causa3="";
    if(y.inval){ lvl3="inv"; st3=y.motivo; causa3="DADOS INVALIDOS"; }
    else {
      if(a1l && a1l.lvl==="err"){ lvl3="err"; st3=a1l.status; }
      if(a2l && a2l.lvl==="err"){ lvl3="err"; st3=(st3==="OK"?a2l.status:st3+" + "+a2l.status); }
      if(lvl3==="ok" && ((a1l&&a1l.lvl==="warn")||(a2l&&a2l.lvl==="warn"))){ lvl3="warn"; st3="ZERA SALDO"; }
      if(a2l) causa3=a2l.causa;
    }
    if(lvl3==="err"||lvl3==="inv") linFalhas++;
    lin.push({linha:y.linha,origem:y.origem,destino:y.destino,classe:y.classe,valor:y.valor,unidade:y.unidade,aprovador:y.aprovador,
              status:st3,lvl:lvl3,causa:causa3,causaKey:(y.inval?"INV":(a2l?a2l.causaKey:"")),
              saldoPep:a1l?a1l.disp:0, saldoCls:a2l?a2l.disp:0, excPep:a1l?a1l.exc:0, excCls:a2l?a2l.exc:0});
  }

  /* --- 8) diagnostico por PEP (uma entrada por PEP com problema) --- */
  var diag=[];
  for(var d1=0;d1<p1.length;d1++){
    var q1=p1[d1];
    var clsRows=p2.filter(function(z){ return z.pep===q1.pep; });
    var temErrCls=clsRows.some(function(z){ return z.lvl==="err"; });
    if(q1.lvl!=="err" && !temErrCls) continue;
    var tipo = q1.lvl==="err" ? "PEP" : "CLS";
    var nC=0,nP=0,nS=0,nOk=0;
    for(var d2=0;d2<clsRows.length;d2++){
      var ck=clsRows[d2].causaKey;
      if(ck==="CASC") nC++; else if(ck==="PC") nP++; else if(ck==="CLS") nS++; else if(clsRows[d2].lvl!=="err") nOk++;
    }
    var resumo;
    if(q1.status==="PEP SEM SALDO") resumo="PEP nao consta na guia SALDO DO PEP"+(q1.foraMask?" (fora da mascara: verifique digitacao)":"")+". Todas as "+clsRows.length+" classe(s) falham em consequencia.";
    else if(tipo==="PEP"){
      resumo="O PEP estoura o saldo em R$ "+fmtBRL(q1.exc)+". ";
      var pt=[];
      if(nC) pt.push(nC+" classe(s) arrastada(s) em cascata");
      if(nP) pt.push(nP+" classe(s) estoura(m) tambem por conta propria");
      if(nOk) pt.push(nOk+" classe(s) cabe(m), mas fica(m) bloqueada(s) pelo PEP");
      resumo+=pt.join("; ")+".";
      resumo+= nP ? " Acao: reduzir R$ "+fmtBRL(q1.exc)+" no PEP e ainda ajustar as classes marcadas PEP + CLASSE."
                  : " Acao: reduzir R$ "+fmtBRL(q1.exc)+" no PEP resolve tudo.";
    } else {
      resumo="O PEP tem saldo (folga R$ "+fmtBRL(q1.folga)+"), mas "+nS+" classe(s) estoura(m). Problema de classe, nao de saldo do PEP. Acao: redistribuir entre classes ou reduzir o valor nas classes indicadas.";
    }
    var linhasPep=[]; for(var d3=0;d3<clsRows.length;d3++) linhasPep=linhasPep.concat(clsRows[d3].linhas);
    diag.push({pep:q1.pep,tipo:tipo,p1:q1,classes:clsRows,nCasc:nC,nPC:nP,nSo:nS,nOk:nOk,resumo:resumo,
               linhas:linhasPep.sort(function(a,b){return a-b;}),aprovadores:q1.aprovadores,unidades:q1.unidades});
  }
  diag.sort(function(a,b){ if(a.tipo!==b.tipo) return a.tipo==="PEP"?-1:1; return (b.p1.exc||0)-(a.p1.exc||0)||b.p1.usado-a.p1.usado; });

  /* --- 9) apontamentos (agrupados por PEP) --- */
  var aps=[];
  for(var e1=0;e1<diag.length;e1++){
    var g=diag[e1], q=g.p1;
    var titulo = g.tipo==="PEP" ? ("Estouro do PEP &mdash; "+q.status) : "Estouro s&oacute; de classe";
    var msg = g.tipo==="PEP"
      ? "PEP <b>"+esc(q.pep)+"</b>: reclassificado R$ "+fmtBRL(q.usado)+" contra saldo de R$ "+fmtBRL(q.disp)+
        (q.status==="EXCEDE SALDO"?" (excedente de R$ "+fmtBRL(q.exc)+")":" (PEP ausente na guia SALDO DO PEP)")+"."
      : "PEP <b>"+esc(q.pep)+"</b>: saldo OK (folga R$ "+fmtBRL(q.folga)+"), mas "+g.nSo+" classe(s) estoura(m).";
    var subs=[];
    for(var e2=0;e2<g.classes.length;e2++){
      var w=g.classes[e2]; if(w.lvl!=="err") continue;
      subs.push("Classe "+w.classe+" ["+w.causa+"]: R$ "+fmtBRL(w.usado)+" x saldo CJI3 R$ "+fmtBRL(w.disp)+" (excedente R$ "+fmtBRL(w.exc)+"). Linha(s): "+w.linhas.join(", ")+".");
    }
    var quem=[]; if(g.unidades.length) quem.push("Unidade: "+g.unidades.join(", ")); if(g.aprovadores.length) quem.push("Aprovador: "+g.aprovadores.join(", "));
    aps.push({lvl:"err",pep:q.pep,classe:"",tipo:g.tipo,titulo:titulo,msg:msg,sub:subs.join("\n"),quem:quem.join(" | ")});
  }
  if(nInval) aps.push({lvl:"inv",pep:"",classe:"",tipo:"INV",titulo:"Dados inv&aacute;lidos",
    msg:nInval+" linha(s) bloqueada(s) por valor em branco/zero/negativo/nao numerico, PEP ou classe vazios. Veja a aba <b>Qualidade dos dados</b>. Essas linhas nao entraram nas somas.",sub:"",quem:""});

  state.result={
    sheets:{recl:nRecl,saldo:nSaldo,cji3:nCji3},
    counts:{linhas:recl.length,saldo:nSaldoLin,cji3:nCjiLin,peps:Object.keys(totPep).length},
    totalGeral:totalGeral, totalInval:totalInval, tol:TOL,
    p1:p1,p2:p2,lin:lin,aps:aps,qa:qa,diag:diag,
    p1Falhas:p1Falhas,p2Falhas:p2Falhas,linFalhas:linFalhas,
    nCasc:nCasc,nPC:nPC,nSoCls:nSoCls,nSoPep:nSoPep,nInval:nInval,nMask:nMask,
    nDiagPep:diag.filter(function(z){return z.tipo==="PEP";}).length,
    nDiagCls:diag.filter(function(z){return z.tipo==="CLS";}).length
  };
  render();
}

/* ---------- render ---------- */
function pctTxt(p){ return p===Infinity?"&#8734;":p.toFixed(2).replace(".",",")+"%"; }
function badge(lvl,txtv){
  var c = lvl==="err"?"b-err":(lvl==="warn"?"b-warn":(lvl==="inv"?"b-inv":"b-ok"));
  return "<span class='badge "+c+"'>"+esc(txtv)+"</span>";
}
function badgeCausa(key,txtv){
  if(!key) return "";
  var c = key==="CASC"?"b-casc":(key==="PC"?"b-pc":(key==="CLS"?"b-cls":(key==="PEP"?"b-pep":"b-inv")));
  return "<span class='badge "+c+"'>"+esc(txtv||key)+"</span>";
}
function rowCls(lvl){ return lvl==="err"?"row-err":(lvl==="warn"?"row-warn":(lvl==="inv"?"row-inv":"")); }

function render(){
  var r=state.result;
  document.getElementById("results").classList.remove("hidden");
  document.getElementById("results").classList.add("reveal");
  document.getElementById("btn-download").classList.remove("hidden");

  document.getElementById("s-linhas").textContent=r.counts.linhas.toLocaleString("pt-BR");
  document.getElementById("s-total").textContent=fmtBRL(r.totalGeral);
  document.getElementById("s-peps").textContent=r.counts.peps.toLocaleString("pt-BR");
  document.getElementById("s-estpep").textContent=r.nDiagPep;
  document.getElementById("s-estcls").textContent=r.nSoCls;
  document.getElementById("s-estinv").textContent=r.nInval;
  document.getElementById("st-estpep").className="stat "+(r.nDiagPep?"err":"ok");
  document.getElementById("st-estcls").className="stat "+(r.nSoCls?"err":"ok");
  document.getElementById("st-estinv").className="stat "+(r.nInval?"err":"ok");

  var total=r.nDiagPep+r.nSoCls+r.nInval;
  var vd=document.getElementById("verdict");
  vd.className="verdict glass "+(total?"err":"ok");
  document.getElementById("v-title").innerHTML = total
      ? "&#10060; Reprovado &mdash; "+r.nDiagPep+" PEP(s) estouram saldo, "+r.nSoCls+" classe(s) estouram isoladamente"+(r.nInval?", "+r.nInval+" linha(s) inv&aacute;lida(s)":"")
      : "&#9989; Aprovado &mdash; todas as reclassifica&ccedil;&otilde;es respeitam os saldos";
  if(total){
    var partes=[];
    if(r.nDiagPep) partes.push("<b>"+r.nDiagPep+"</b> PEP(s) estouram o saldo cont&aacute;bil, arrastando <b>"+r.nCasc+"</b> classe(s) em cascata"+
                               (r.nPC?" (+<b>"+r.nPC+"</b> classe(s) que estouram tamb&eacute;m al&eacute;m do PEP)":"")+
                               (r.nSoPep?"; "+r.nSoPep+" classe(s) desses PEPs cabem, mas ficam bloqueadas":""));
    if(r.nSoCls) partes.push("<b>"+r.nSoCls+"</b> classe(s) estouram com o PEP dentro do saldo (problema s&oacute; de classe)");
    if(r.nInval) partes.push("<b>"+r.nInval+"</b> linha(s) com dados inv&aacute;lidos (R$ "+fmtBRL(r.totalInval)+" fora das somas)");
    document.getElementById("v-text").innerHTML = partes.join(". ")+". <b>"+r.linFalhas+"</b> de <b>"+r.counts.linhas+
        "</b> linha(s) bloqueadas. Premissa 1: "+r.p1Falhas+" falha(s); Premissa 2: "+r.p2Falhas+" falha(s). Veja a aba <b>Diagn&oacute;stico por PEP</b>.";
  } else {
    document.getElementById("v-text").innerHTML = "As <b>"+r.counts.linhas+"</b> linhas respeitam o saldo dos <b>"+r.counts.peps+
        "</b> PEPs de origem e o saldo por classe de custo na CJI3, com toler&acirc;ncia de R$ "+fmtBRL(r.tol)+".";
  }

  var nQa=r.qa.length;
  setCnt("n-dg",r.diag.length,r.diag.length?"err":"ok");
  setCnt("n-ap",r.aps.length,r.aps.length?"err":"ok");
  setCnt("n-p1",r.p1.length,r.p1Falhas?"err":"ok");
  setCnt("n-p2",r.p2.length,r.p2Falhas?"err":"ok");
  setCnt("n-lin",r.counts.linhas,r.linFalhas?"err":"ok");
  setCnt("n-qa",nQa,nQa?"":"ok");

  /* apontamentos */
  var ap=document.getElementById("ap-list");
  if(!r.aps.length){
    ap.innerHTML="<div class='alert ok'><b>Nenhum apontamento.</b> Todas as reclassifica&ccedil;&otilde;es cabem no saldo do PEP de origem e no saldo da respectiva classe de custo na CJI3.</div>";
  } else {
    var h="";
    for(var i=0;i<r.aps.length;i++){
      var A=r.aps[i];
      h+="<div class='alert "+(A.lvl==="err"?"err":"warn")+"'><b>"+A.titulo+"</b>"+
         (A.quem?" <span style='font-weight:500;opacity:.8'>&middot; "+esc(A.quem)+"</span>":"")+"<br>"+A.msg+
         (A.sub?"<span class='sub'>"+esc(A.sub).replace(/\n/g,"<br>")+"</span>":"")+"</div>";
    }
    ap.innerHTML=h;
  }

  /* qualidade */
  var qa=document.getElementById("qa-list");
  if(!r.qa.length){
    qa.innerHTML="<div class='alert ok'><b>Sem observa&ccedil;&otilde;es.</b> Valores num&eacute;ricos, PEPs e classes preenchidos corretamente.</div>";
  } else {
    var hq="";
    for(var q=0;q<r.qa.length;q++){
      hq+="<div class='alert "+(r.qa[q].lvl==="err"?"err":"warn")+"'>"+
          (r.qa[q].linha?"<b>Linha "+r.qa[q].linha+":</b> ":"")+esc(r.qa[q].msg)+"</div>";
    }
    qa.innerHTML=hq;
  }

  renderDiag(); renderP1(); renderP2(); renderLin();
}
function setCnt(id,n,cls){
  var el=document.getElementById(id);
  el.textContent=n.toLocaleString("pt-BR");
  el.className="cnt"+(cls?" "+cls:"");
}

function renderDiag(){
  var r=state.result;
  var f=document.getElementById("f-dg").value.trim().toUpperCase();
  var tipo=document.getElementById("o-dg").value;
  var out=document.getElementById("dg-list");
  if(!r.diag.length){
    out.innerHTML="<div class='alert ok'><b>Nenhum PEP com problema de saldo.</b>"+(r.nInval?" H&aacute; "+r.nInval+" linha(s) com dados inv&aacute;lidos na aba <b>Qualidade dos dados</b>.":"")+"</div>";
    return;
  }
  var h="";
  for(var i=0;i<r.diag.length;i++){
    var g=r.diag[i], q=g.p1;
    if(tipo && g.tipo!==tipo) continue;
    if(f && (g.pep+" "+g.aprovadores.join(" ")+" "+g.unidades.join(" ")).toUpperCase().indexOf(f)<0) continue;
    var cardCls = g.tipo==="PEP" ? "" : " cls";
    var tags="";
    if(g.tipo==="PEP"){
      tags+=badge("err",q.status);
      if(g.nCasc) tags+=" "+badgeCausa("CASC",g.nCasc+" em cascata");
      if(g.nPC)   tags+=" "+badgeCausa("PC",g.nPC+" PEP + classe");
      if(g.nOk)   tags+=" "+badgeCausa("PEP",g.nOk+" classe(s) ok");
    } else {
      tags+=badgeCausa("CLS",g.nSo+" classe(s) estoura(m)")+" "+badge("ok","PEP com saldo");
    }
    var quem="";
    if(g.unidades.length) quem+="<span class='kv'><span class='k'>Unidade</span><span class='v' style='font-family:inherit'>"+esc(g.unidades.join(", "))+"</span></span>";
    if(g.aprovadores.length) quem+="<span class='kv'><span class='k'>Aprovador</span><span class='v' style='font-family:inherit'>"+esc(g.aprovadores.join(", "))+"</span></span>";
    h+="<div class='diag glass"+cardCls+"' data-i='"+i+"'>"+
       "<div class='diag-h'>"+
         "<span class='pep'>"+esc(g.pep)+"</span>"+
         "<span class='kv'><span class='k'>Reclassificado</span><span class='v'>"+fmtBRL(q.usado)+"</span></span>"+
         "<span class='kv'><span class='k'>Saldo do PEP</span><span class='v'>"+fmtBRL(q.disp)+"</span></span>"+
         "<span class='kv'><span class='k'>"+(q.lvl==="err"?"Excedente":"Folga")+"</span><span class='v "+(q.lvl==="err"?"neg":"")+"'>"+fmtBRL(q.lvl==="err"?q.exc:q.folga)+"</span></span>"+
         quem+
         "<span style='display:flex;gap:6px;flex-wrap:wrap'>"+tags+"</span>"+
         "<span class='caret'>&#9654;</span>"+
       "</div>"+
       "<div class='diag-b'>"+
         "<p class='why'>"+esc(g.resumo)+"</p>"+
         "<div class='tbl-scroll' style='max-height:none'><table><thead><tr><th>Classe</th><th class='num'>Reclassificado</th><th class='num'>Saldo CJI3</th><th class='num'>Excedente</th><th>Status</th><th>Causa</th><th>Linhas</th></tr></thead><tbody>";
    var notas="";
    for(var j=0;j<g.classes.length;j++){
      var w=g.classes[j];
      h+="<tr class='"+rowCls(w.lvl)+"'>"+
         "<td class='mono'>"+esc(w.classe)+"</td>"+
         "<td class='num'>"+fmtBRL(w.usado)+"</td>"+
         "<td class='num'>"+fmtBRL(w.disp)+"</td>"+
         "<td class='num "+(w.exc>0?"neg":"")+"'>"+(w.exc>0?fmtBRL(w.exc):"&mdash;")+"</td>"+
         "<td>"+badge(w.lvl,w.status)+"</td>"+
         "<td>"+badgeCausa(w.causaKey,w.causa)+"</td>"+
         "<td class='mono' title='"+esc(w.linhas.join(", "))+"'>"+(w.linhas.length>8?w.linhas.slice(0,8).join(", ")+" &hellip; (+"+(w.linhas.length-8)+")":w.linhas.join(", "))+"</td></tr>";
      if(w.lvl==="err"){
        var altIdx=(w.obs||"").indexOf("Classes do mesmo PEP");
        notas+="<li><span class='mono'>"+esc(w.classe)+"</span> "+badgeCausa(w.causaKey,w.causa)+" "+esc(w.causaTxt||"")+
               (w.status==="CLASSE INEXISTENTE"?" "+esc(w.obs.split(" Classes do mesmo PEP")[0]):"")+
               (altIdx>-1?" <span style='color:var(--ink-mute)'>"+esc(w.obs.slice(altIdx))+"</span>":"")+"</li>";
      }
    }
    h+="</tbody></table></div>"+
       (notas?"<ul class='notas'>"+notas+"</ul>":"")+
       "<div class='meta'>"+g.classes.length+" classe(s) &middot; "+g.linhas.length+" linha(s) da planilha: "+g.linhas.join(", ")+"</div>"+
       "</div></div>";
  }
  out.innerHTML=h||"<div class='alert ok'>Nenhum PEP corresponde ao filtro.</div>";
  var hs=out.querySelectorAll(".diag-h");
  for(var k=0;k<hs.length;k++) hs[k].addEventListener("click",function(){ this.parentNode.classList.toggle("open"); });
  if(r.diag.length<=3){ var ds=out.querySelectorAll(".diag"); for(var k2=0;k2<ds.length;k2++) ds[k2].classList.add("open"); }
}
function toggleDiag(){
  var ds=document.querySelectorAll("#dg-list .diag"); if(!ds.length) return;
  var anyClosed=false; for(var i=0;i<ds.length;i++) if(!ds[i].classList.contains("open")) anyClosed=true;
  for(var j=0;j<ds.length;j++) ds[j].classList.toggle("open",anyClosed);
  document.getElementById("dg-toggle").textContent=anyClosed?"Recolher todos":"Expandir todos";
}

function renderP1(){
  var r=state.result;
  var f=document.getElementById("f-p1").value.trim().toUpperCase();
  var only=document.getElementById("o-p1").checked;
  var h="<thead><tr><th>PEP origem</th><th class='num'>Reclassificado</th><th class='num'>Saldo do PEP</th><th class='num'>Folga</th><th class='num'>Uso</th><th>Ocupa&ccedil;&atilde;o</th><th>Status</th><th>Classes</th><th>Aprovador</th></tr></thead><tbody>";
  for(var i=0;i<r.p1.length;i++){
    var q=r.p1[i];
    if(only && q.lvl==="ok") continue;
    if(f && (q.pep+" "+q.aprovadores.join(" ")).toUpperCase().indexOf(f)<0) continue;
    var w=q.pct===Infinity?100:Math.min(q.pct,100);
    var cls=q.lvl==="err"?"over":(q.lvl==="warn"?"full":"");
    h+="<tr class='"+rowCls(q.lvl)+"'>"+
       "<td class='mono'>"+esc(q.pep)+"</td>"+
       "<td class='num'>"+fmtBRL(q.usado)+"</td>"+
       "<td class='num'>"+fmtBRL(q.disp)+"</td>"+
       "<td class='num "+(q.folga<0?"neg":"")+"'>"+fmtBRL(q.folga)+"</td>"+
       "<td class='num'>"+pctTxt(q.pct)+"</td>"+
       "<td><div class='bar'><i class='"+cls+"' style='width:"+w.toFixed(1)+"%'></i></div></td>"+
       "<td>"+badge(q.lvl,q.status)+"</td>"+
       "<td>"+(q.nClsErr?"<span class='neg'>"+q.nClsErr+" de "+q.nCls+" estoura(m)</span>":q.nCls+" ok")+"</td>"+
       "<td style='white-space:normal'>"+esc(q.aprovadores.join(", "))+"</td></tr>";
  }
  document.getElementById("t-p1").innerHTML=h+"</tbody>";
}

function renderP2(){
  var r=state.result;
  var f=document.getElementById("f-p2").value.trim().toUpperCase();
  var only=document.getElementById("o-p2").checked;
  var h="<thead><tr><th>PEP origem</th><th>Classe</th><th class='num'>Reclassificado</th><th class='num'>Saldo CJI3</th><th class='num'>Folga</th><th>Linhas</th><th>Status</th><th>Causa</th><th>Observa&ccedil;&atilde;o</th></tr></thead><tbody>";
  for(var i=0;i<r.p2.length;i++){
    var w=r.p2[i];
    if(only && w.lvl==="ok") continue;
    if(f && (w.pep+" "+w.classe+" "+w.causa).toUpperCase().indexOf(f)<0) continue;
    h+="<tr class='"+rowCls(w.lvl)+"'>"+
       "<td class='mono'>"+esc(w.pep)+"</td>"+
       "<td class='mono'>"+esc(w.classe)+"</td>"+
       "<td class='num'>"+fmtBRL(w.usado)+"</td>"+
       "<td class='num'>"+fmtBRL(w.disp)+"</td>"+
       "<td class='num "+(w.folga<0?"neg":"")+"'>"+fmtBRL(w.folga)+"</td>"+
       "<td class='mono' style='white-space:normal;max-width:200px'>"+w.linhas.join(", ")+"</td>"+
       "<td>"+badge(w.lvl,w.status)+"</td>"+
       "<td>"+badgeCausa(w.causaKey,w.causa)+"</td>"+
       "<td style='white-space:normal;max-width:420px'>"+esc((w.causaTxt?w.causaTxt+" ":"")+(w.obs||""))+"</td></tr>";
  }
  document.getElementById("t-p2").innerHTML=h+"</tbody>";
}

function renderLin(){
  var r=state.result;
  var f=document.getElementById("f-lin").value.trim().toUpperCase();
  var only=document.getElementById("o-lin").checked;
  var h="<thead><tr><th class='num'>Linha</th><th>Unidade</th><th>PEP origem (DE)</th><th>PEP destino (PARA)</th><th>Classe</th><th class='num'>Valor</th><th class='num'>Saldo do PEP</th><th class='num'>Saldo da classe</th><th>Status</th><th>Causa</th><th>Aprovador</th></tr></thead><tbody>";
  for(var i=0;i<r.lin.length;i++){
    var y=r.lin[i];
    if(only && y.lvl==="ok") continue;
    if(f && (y.origem+" "+y.destino+" "+y.classe+" "+y.linha+" "+y.aprovador+" "+y.unidade+" "+y.causa).toUpperCase().indexOf(f)<0) continue;
    h+="<tr class='"+rowCls(y.lvl)+"'>"+
       "<td class='num'>"+y.linha+"</td>"+
       "<td>"+esc(y.unidade)+"</td>"+
       "<td class='mono'>"+esc(y.origem)+"</td>"+
       "<td class='mono'>"+esc(y.destino)+"</td>"+
       "<td class='mono'>"+esc(y.classe)+"</td>"+
       "<td class='num "+(y.valor<0?"neg":"")+"'>"+fmtBRL(y.valor)+"</td>"+
       "<td class='num'>"+fmtBRL(y.saldoPep)+"</td>"+
       "<td class='num'>"+fmtBRL(y.saldoCls)+"</td>"+
       "<td>"+badge(y.lvl,y.status)+"</td>"+
       "<td>"+badgeCausa(y.causaKey,y.causa)+"</td>"+
       "<td style='white-space:normal'>"+esc(y.aprovador)+"</td></tr>";
  }
  document.getElementById("t-lin").innerHTML=h+"</tbody>";
}

/* ---------- download da conferencia ---------- */
var HDR={fill:{patternType:"solid",fgColor:{rgb:"FF005AAB"}},font:{bold:true,color:{rgb:"FFFFFFFF"}}};
var TIT={font:{bold:true,sz:12}};
function estilizaCabecalho(ws,rowIdx,nCols){
  for(var c=0;c<nCols;c++){
    var addr=XLSX.utils.encode_cell({r:rowIdx,c:c});
    if(ws[addr]) ws[addr].s=HDR;
  }
}
function baixar(){
  var r=state.result; if(!r) return;
  var wb=XLSX.utils.book_new();

  // Resumo
  var total=r.nDiagPep+r.nSoCls+r.nInval;
  var res=[
    ["VALIDACAO DE RECLASSIFICACOES DE PEP"],
    ["Arquivo analisado", state.name||""],
    ["Gerado em", new Date().toLocaleString("pt-BR")],
    [],
    ["VEREDITO", total?("REPROVADO - "+r.nDiagPep+" PEP(s) estouram saldo; "+r.nSoCls+" classe(s) estouram isoladamente; "+r.nInval+" linha(s) invalida(s)"):"APROVADO - sem apontamentos"],
    [],
    ["Linhas de reclassificacao", r.counts.linhas],
    ["Total reclassificado (R$)", r.totalGeral],
    ["Total em linhas invalidas, fora das somas (R$)", r.totalInval],
    ["PEPs de origem distintos", r.counts.peps],
    ["Linhas lidas em SALDO DO PEP", r.counts.saldo],
    ["Linhas lidas em CJI3", r.counts.cji3],
    ["Tolerancia aplicada (R$)", r.tol],
    [],
    ["CAUSA RAIZ",""],
    ["PEPs que estouram o saldo contabil", r.nDiagPep],
    ["  - classes arrastadas em CASCATA", r.nCasc],
    ["  - classes que estouram ALEM do PEP (PEP + CLASSE)", r.nPC],
    ["  - classes que cabem, mas bloqueadas pelo PEP (SO PEP)", r.nSoPep],
    ["Classes que estouram com PEP OK (SO CLASSE)", r.nSoCls],
    ["Linhas com dados invalidos (branco/zero/negativo/nao numerico)", r.nInval],
    ["PEPs fora da mascara (avisos)", r.nMask],
    [],
    ["Premissa 1 - saldo do PEP: falhas", r.p1Falhas],
    ["Premissa 2 - saldo da classe: falhas", r.p2Falhas],
    ["Linhas bloqueadas", r.linFalhas]
  ];
  var wsR=XLSX.utils.aoa_to_sheet(res);
  wsR["!cols"]=[{wch:58},{wch:70}];
  if(wsR["A1"]) wsR["A1"].s=TIT;
  if(wsR["A5"]) wsR["A5"].s=HDR;
  if(wsR["B5"]) wsR["B5"].s=HDR;
  if(wsR["A15"]) wsR["A15"].s=HDR;
  if(wsR["B15"]) wsR["B15"].s=HDR;
  XLSX.utils.book_append_sheet(wb,wsR,"Resumo");

  // Diagnostico por PEP
  var dg=[["PEP origem","Tipo","Status PEP","Reclassificado","Saldo do PEP","Excedente PEP","Classe","Reclassificado classe","Saldo CJI3","Excedente classe","Status classe","Causa","Leitura","Linhas","Unidade","Aprovador"]];
  if(!r.diag.length) dg.push(["","OK","","","","","","","","","","","Nenhum PEP com problema de saldo.","","",""]);
  for(var g0=0;g0<r.diag.length;g0++){ var G=r.diag[g0];
    for(var g1=0;g1<G.classes.length;g1++){ var W=G.classes[g1];
      dg.push([G.pep, G.tipo==="PEP"?"ESTOURO DO PEP":"SO CLASSE", G.p1.status, G.p1.usado, G.p1.disp, G.p1.exc,
               W.classe, W.usado, W.disp, W.exc, W.status, W.causa, W.causaTxt||W.obs||"", W.linhas.join(", "),
               G.unidades.join(", "), G.aprovadores.join(", ")]);
    }
  }
  var wsD=XLSX.utils.aoa_to_sheet(dg);
  wsD["!cols"]=[{wch:24},{wch:16},{wch:16},{wch:18},{wch:18},{wch:16},{wch:14},{wch:18},{wch:18},{wch:16},{wch:20},{wch:14},{wch:90},{wch:30},{wch:12},{wch:24}];
  estilizaCabecalho(wsD,0,16);
  formatCols(wsD,dg.length,[3,4,5,7,8,9],"#,##0.00");
  XLSX.utils.book_append_sheet(wb,wsD,"Diagnostico por PEP");

  // Apontamentos
  var ap=[["PEP","Tipo","Titulo","Descricao","Classes afetadas","Unidade - Aprovador"]];
  if(!r.aps.length) ap.push(["","OK","Sem apontamentos","Todas as reclassificacoes respeitam os saldos.","",""]);
  for(var i=0;i<r.aps.length;i++){
    ap.push([r.aps[i].pep||"", r.aps[i].tipo||"",
             String(r.aps[i].titulo).replace(/&mdash;/g,"-").replace(/&oacute;/g,"o").replace(/&aacute;/g,"a"),
             String(r.aps[i].msg).replace(/<[^>]+>/g,"").replace(/&oacute;/g,"o").replace(/&aacute;/g,"a"), r.aps[i].sub||"", r.aps[i].quem||""]);
  }
  var wsA=XLSX.utils.aoa_to_sheet(ap);
  wsA["!cols"]=[{wch:24},{wch:8},{wch:34},{wch:90},{wch:100},{wch:40}];
  estilizaCabecalho(wsA,0,6);
  XLSX.utils.book_append_sheet(wb,wsA,"Apontamentos");

  // Premissa 1
  var a1=[["PEP origem","Reclassificado","Saldo do PEP","Folga","Excedente","Uso %","Status","Classes","Classes que estouram","Unidade","Aprovador"]];
  for(var b=0;b<r.p1.length;b++){ var q=r.p1[b];
    a1.push([q.pep,q.usado,q.disp,q.folga,q.exc,(q.pct===Infinity?null:q.pct/100),q.status,q.nCls,q.nClsErr,q.unidades.join(", "),q.aprovadores.join(", ")]); }
  var ws1=XLSX.utils.aoa_to_sheet(a1);
  ws1["!cols"]=[{wch:24},{wch:18},{wch:18},{wch:18},{wch:16},{wch:10},{wch:18},{wch:9},{wch:12},{wch:12},{wch:24}];
  estilizaCabecalho(ws1,0,11);
  formatCols(ws1,a1.length,[1,2,3,4],"#,##0.00"); formatCols(ws1,a1.length,[5],"0.00%");
  XLSX.utils.book_append_sheet(wb,ws1,"Premissa 1 - Saldo PEP");

  // Premissa 2
  var a2=[["PEP origem","Classe","Reclassificado","Saldo CJI3","Folga","Excedente","Linhas","Status","Causa","Leitura","Observacao"]];
  for(var c2=0;c2<r.p2.length;c2++){ var w=r.p2[c2];
    a2.push([w.pep,w.classe,w.usado,w.disp,w.folga,w.exc,w.linhas.join(", "),w.status,w.causa,w.causaTxt||"",w.obs||""]); }
  var ws2=XLSX.utils.aoa_to_sheet(a2);
  ws2["!cols"]=[{wch:24},{wch:14},{wch:18},{wch:18},{wch:18},{wch:16},{wch:20},{wch:20},{wch:14},{wch:80},{wch:80}];
  estilizaCabecalho(ws2,0,11);
  formatCols(ws2,a2.length,[2,3,4,5],"#,##0.00");
  XLSX.utils.book_append_sheet(wb,ws2,"Premissa 2 - Saldo Classe");

  // Detalhe por linha
  var a3=[["Linha","Unidade","PEP origem","PEP destino","Classe","Valor","Saldo do PEP","Saldo da classe","Status","Causa","Aprovador"]];
  for(var d=0;d<r.lin.length;d++){ var y=r.lin[d];
    a3.push([y.linha,y.unidade,y.origem,y.destino,y.classe,y.valor,y.saldoPep,y.saldoCls,y.status,y.causa,y.aprovador]); }
  var ws3=XLSX.utils.aoa_to_sheet(a3);
  ws3["!cols"]=[{wch:8},{wch:10},{wch:24},{wch:24},{wch:14},{wch:18},{wch:18},{wch:18},{wch:28},{wch:16},{wch:24}];
  estilizaCabecalho(ws3,0,11);
  formatCols(ws3,a3.length,[5,6,7],"#,##0.00");
  XLSX.utils.book_append_sheet(wb,ws3,"Detalhe por linha");

  // Qualidade
  var a4=[["Nivel","Linha","Observacao"]];
  if(!r.qa.length) a4.push(["OK","","Sem observacoes de qualidade de dados."]);
  for(var e=0;e<r.qa.length;e++) a4.push([r.qa[e].lvl.toUpperCase(), r.qa[e].linha||"", r.qa[e].msg]);
  var ws4=XLSX.utils.aoa_to_sheet(a4);
  ws4["!cols"]=[{wch:8},{wch:10},{wch:110}];
  estilizaCabecalho(ws4,0,3);
  XLSX.utils.book_append_sheet(wb,ws4,"Qualidade dos dados");

  var base=(state.name||"reclassificacoes.xlsx").replace(/\.(xlsx|xlsm|xls)$/i,"");
  XLSX.writeFile(wb, base+"_CONFERENCIA.xlsx", {cellStyles:true});
}
function formatCols(ws,nRows,cols,z){
  for(var r=1;r<nRows;r++){
    for(var i=0;i<cols.length;i++){
      var addr=XLSX.utils.encode_cell({r:r,c:cols[i]});
      if(ws[addr] && ws[addr].t==="n") ws[addr].z=z;
    }
  }
}

/* ---------- UI ---------- */
function showMsg(t){ var m=document.getElementById("msg"); m.textContent=t; m.className="msg show error"; }
function clearMsg(){ document.getElementById("msg").className="msg"; }

function readWorkbook(file){
  return new Promise(function(res,rej){
    var fr=new FileReader();
    fr.onload=function(e){ try{ res(XLSX.read(new Uint8Array(e.target.result),{type:"array",cellDates:true})); }catch(err){ rej(err); } };
    fr.onerror=rej;
    fr.readAsArrayBuffer(file);
  });
}
function loadFile(file){
  clearMsg();
  var drop=document.getElementById("drop-x"), meta=document.getElementById("meta-x");
  readWorkbook(file).then(function(wb){
    state.wb=wb; state.name=file.name;
    drop.classList.add("loaded"); meta.classList.remove("hidden");
    meta.textContent="OK: "+file.name+"  -  guias: "+wb.SheetNames.join(", ");
    document.getElementById("btn-run").disabled=false;
  }).catch(function(err){ showMsg("Erro ao ler "+file.name+": "+err.message); });
}

window.addEventListener("dragover",function(e){ e.preventDefault(); }, false);
window.addEventListener("drop",function(e){ e.preventDefault(); }, false);

function applyTheme(t){
  if(t==="light") document.body.classList.add("light"); else document.body.classList.remove("light");
  var opts=document.querySelectorAll(".theme-toggle .opt");
  for(var i=0;i<opts.length;i++){ opts[i].classList.toggle("on", opts[i].getAttribute("data-th")===t); }
  try{ localStorage.setItem("reclasscheck-theme", t); }catch(e){}
}
(function(){ var t="dark"; try{ t=localStorage.getItem("reclasscheck-theme")||"dark"; }catch(e){} if(t==="light") document.body.classList.add("light"); })();

document.addEventListener("DOMContentLoaded",function(){
  var saved="dark"; try{ saved=localStorage.getItem("reclasscheck-theme")||"dark"; }catch(e){}
  applyTheme(saved);
  document.getElementById("theme-toggle").addEventListener("click",function(){
    applyTheme(document.body.classList.contains("light")?"dark":"light");
  });

  var drop=document.getElementById("drop-x"), inp=document.getElementById("file-x");
  drop.addEventListener("click",function(){ inp.click(); });
  drop.addEventListener("mousemove",function(e){ var rc=drop.getBoundingClientRect(); drop.style.setProperty("--mx",(e.clientX-rc.left)+"px"); drop.style.setProperty("--my",(e.clientY-rc.top)+"px"); });
  drop.addEventListener("dragover",function(e){ e.preventDefault(); drop.classList.add("drag"); });
  drop.addEventListener("dragleave",function(){ drop.classList.remove("drag"); });
  drop.addEventListener("drop",function(e){ e.preventDefault(); drop.classList.remove("drag"); if(e.dataTransfer.files[0]) loadFile(e.dataTransfer.files[0]); });
  inp.addEventListener("change",function(e){ if(e.target.files[0]) loadFile(e.target.files[0]); });

  document.getElementById("btn-run").addEventListener("click",function(){
    clearMsg();
    var btn=this; btn.disabled=true; btn.textContent="Processando...";
    setTimeout(function(){
      try{ processar(); document.getElementById("results").scrollIntoView({behavior:"smooth"}); }
      catch(err){ showMsg("Erro ao processar: "+err.message); console.error(err); }
      btn.disabled=false; btn.innerHTML="&#9889; Validar";
    },30);
  });
  document.getElementById("btn-download").addEventListener("click",baixar);
  document.getElementById("btn-reset").addEventListener("click",function(){ location.reload(); });

  document.getElementById("f-dg").addEventListener("input",renderDiag);
  document.getElementById("o-dg").addEventListener("change",renderDiag);
  document.getElementById("dg-toggle").addEventListener("click",toggleDiag);
  document.getElementById("f-p1").addEventListener("input",renderP1);
  document.getElementById("o-p1").addEventListener("change",renderP1);
  document.getElementById("f-p2").addEventListener("input",renderP2);
  document.getElementById("o-p2").addEventListener("change",renderP2);
  document.getElementById("f-lin").addEventListener("input",renderLin);
  document.getElementById("o-lin").addEventListener("change",renderLin);

  var tabs=document.querySelectorAll(".tab");
  for(var i=0;i<tabs.length;i++){
    tabs[i].addEventListener("click",function(){
      var all=document.querySelectorAll(".tab"); for(var k=0;k<all.length;k++) all[k].classList.remove("active");
      var bodies=document.querySelectorAll(".panel-body"); for(var b=0;b<bodies.length;b++) bodies[b].classList.remove("active");
      this.classList.add("active"); document.getElementById("pb-"+this.getAttribute("data-t")).classList.add("active");
    });
  }
});
</script>
</body>
</html>
