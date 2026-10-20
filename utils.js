const ARQUIVO_BOLETIM = 'Boletim.xlsx';
const ARQUIVO_CUSTEIO = 'CUSTEIO_2026.xlsx';


/* 2. REGISTRO GLOBAL DO PLUGIN DE RÓTULOS (DATALABELS) */
if(typeof ChartDataLabels !== 'undefined'){
  Chart.register(ChartDataLabels);
  console.log('✅ chartjs-plugin-datalabels registrado.');
} else {
  console.warn('⚠ chartjs-plugin-datalabels NÃO foi carregado. Verifique o CDN no index.html.');
}


/* 3. UTILITÁRIOS */
const MESES = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho',
               'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
const MESES_ABREV = ['Jan','Fev','Mar','Abr','Mai','Jun',
                     'Jul','Ago','Set','Out','Nov','Dez'];

const fmtBRL = v => (v || 0).toLocaleString('pt-BR',
  { style:'currency', currency:'BRL', maximumFractionDigits:0 });
const fmtNum = v => (v || 0).toLocaleString('pt-BR');
const fmtPct = v => (v || 0).toFixed(1) + '%';

const fmtBRLcompact = v => {
  if(v === null || v === undefined || v === '') return '';
  const abs = Math.abs(v);
  if(abs >= 1_000_000) return 'R$ ' + (v/1_000_000).toFixed(1).replace('.',',') + 'M';
  if(abs >= 1_000)     return 'R$ ' + (v/1_000).toFixed(1).replace('.',',') + 'k';
  return 'R$ ' + v.toFixed(0);
};

function toNumber(v){
  if(v === null || v === undefined || v === '') return 0;
  if(typeof v === 'number') return v;
  let s = String(v).replace(/[R$\s]/g,'');
  if(s.includes(',')) s = s.replace(/\./g,'').replace(',','.');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function normalizarMes(v){
  if(v === null || v === undefined) return null;
  const s = String(v).trim();
  const low = s.toLowerCase();
  for(let i = 0; i < MESES.length; i++){
    if(low.startsWith(MESES[i].toLowerCase().slice(0,3)) ||
       low === MESES[i].toLowerCase()) return i;
  }
  const m = s.match(/\d{1,2}/);
  if(m){
    const n = parseInt(m[0]);
    if(n >= 1 && n <= 12) return n - 1;
  }
  return null;
}

function acharColuna(headers, ...candidatos){
  const norm = s => String(s).toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/[^a-z0-9]/g,'');
  const H = headers.map(h => ({ orig:h, norm:norm(h) }));

  for(const cand of candidatos){
    const c = norm(cand);
    const f = H.find(h => h.norm === c);
    if(f) return f.orig;
  }
  for(const cand of candidatos){
    const c = norm(cand);
    const f = H.find(h => h.norm.includes(c));
    if(f) return f.orig;
  }
  return null;
}

/* Opções base dos gráficos — COM datalabels global ativo */
function chartBaseOptions(){
  return {
    responsive:true,
    maintainAspectRatio:false,
    interaction:{ mode:'index', intersect:false },
    layout:{ padding:{ top: 26 } },
    plugins:{
      legend:{
        labels:{ color:'#e8eefc', font:{ size:12 }, boxWidth:14, boxHeight:10 }
      },
      tooltip:{
        backgroundColor:'#0b1220',
        borderColor:'#2a3554',
        borderWidth:1,
        titleColor:'#e8eefc',
        bodyColor:'#c7d4ee',
        callbacks:{
          label: ctx => `${ctx.dataset.label}: ${fmtBRL(ctx.parsed.y)}`
        }
      },
      /* datalabels global — vale para TODOS os datasets,
         cada dataset pode sobrescrever cor/display */
      datalabels:{
        color:'#e8eefc',
        anchor:'end',
        align:'top',
        offset:2,
        clamp:true,
        font:{ size:10, weight:'600' },
        formatter: v => (v === null || v === undefined || v === 0)
                        ? ''
                        : fmtBRLcompact(v)
      }
    },
    scales:{
      x:{ ticks:{ color:'#8ea0c4' }, grid:{ color:'rgba(255,255,255,.04)' } },
      y:{
        ticks:{ color:'#8ea0c4', callback: v => 'R$ ' + fmtNum(v) },
        grid:{ color:'rgba(255,255,255,.06)' }
      }
    }
  };
}

function mostrarErro(elId, msg){
  const el = document.getElementById(elId);
  if(!el) return;
  el.innerHTML = `
    <div class="status-msg">⚠ ${msg}</div>
    <div class="status-msg status-ok"></div>`;
}


/* 4. CARREGAMENTO DE EXCEL */
async function carregarExcel(url){
  console.log('[loader] Buscando:', url);
  const resp = await fetch(url, { cache:'no-cache' });
  if(!resp.ok) throw new Error(`HTTP ${resp.status} ao buscar ${url}`);

  const buf = await resp.arrayBuffer();
  const wb  = XLSX.read(buf, { type:'array', cellDates:true });
  const ws  = wb.Sheets[wb.SheetNames[0]];
  if(!ws) throw new Error('Planilha vazia: ' + url);

  const rows = XLSX.utils.sheet_to_json(ws, { defval:null });
  console.log(`[loader] OK — ${rows.length} linhas de ${url}`);
  if(rows.length) console.log('[loader] Colunas detectadas:', Object.keys(rows[0]));
  return rows;
}


/* 5. SEÇÃO BOLETINS */
const stateBoletim = {
  dados: [],
  chartBarras: null
};

async function initBoletim(){
  try{
    const rows = await carregarExcel(ARQUIVO_BOLETIM);
    if(!rows.length){
      mostrarErro('boletimStatus', 'Boletim.xlsx está vazio.');
      return;
    }

    const headers  = Object.keys(rows[0]);
    const colMes   = acharColuna(headers, 'mes', 'mês', 'month', 'periodo', 'data');
    const colValor = acharColuna(headers, 'recebido', 'receita', 'valorrecebido', 'valor');
    const colMeta  = acharColuna(headers, 'meta', 'metamensal', 'objetivo');

    if(!colMes || !colValor){
      mostrarErro('boletimStatus',
        'Colunas de Mês ou Valor Recebido não encontradas. Colunas: ' +
        headers.join(' | '));
      return;
    }

    stateBoletim.dados = rows
      .map(r => {
        const mi = normalizarMes(r[colMes]);
        return {
          mesIdx:   mi,
          mes:      mi !== null ? MESES[mi] : String(r[colMes]),
          recebido: toNumber(r[colValor]),
          meta:     colMeta ? toNumber(r[colMeta]) : 0
        };
      })
      .filter(x => x.mesIdx !== null)
      .sort((a,b) => a.mesIdx - b.mesIdx);

    const sel = document.getElementById('bFilterMes');
    sel.innerHTML = '<option value="">Todos os meses</option>' +
      stateBoletim.dados.map(b =>
        `<option value="${b.mesIdx}">${b.mes}</option>`).join('');
    sel.onchange = renderBoletim;

    renderBoletim();
  }catch(e){
    mostrarErro('boletimStatus', 'Erro ao carregar Boletim: ' + e.message);
  }
}

function renderBoletim(){
  const filtro = document.getElementById('bFilterMes').value;
  let dados = stateBoletim.dados;
  if(filtro !== '') dados = dados.filter(d => d.mesIdx === parseInt(filtro));

  const totalRecebido = stateBoletim.dados.reduce((s,x) => s + x.recebido, 0);
  const metaAnual     = stateBoletim.dados.reduce((s,x) => s + x.meta, 0);
  const metaMensal    = stateBoletim.dados.length ? (metaAnual / stateBoletim.dados.length) : 0;


  const mesesComRecebido = stateBoletim.dados.filter(d => d.recebido > 0);
  const ultimoMesIdx = mesesComRecebido.length
    ? mesesComRecebido[mesesComRecebido.length - 1].mesIdx
    : -1;

  const mesesRestantes = Math.max(0, 11 - ultimoMesIdx);
  const gap            = Math.max(0, metaAnual - totalRecebido);
  const gapPorMes      = mesesRestantes > 0 ? gap / mesesRestantes : 0;

  document.getElementById('boletimCards').innerHTML = `
    <div class="card green">
      <div class="label">Total Recebido</div>
      <div class="value">${fmtBRL(totalRecebido)}</div>
      <div class="sub">Acumulado até ${mesesComRecebido.length
        ? mesesComRecebido[mesesComRecebido.length-1].mes : '—'}</div>
    </div>
    <div class="card">
      <div class="label">Meta Anual</div>
      <div class="value">${fmtBRL(metaAnual)}</div>
      <div class="sub">Meta mensal: ${fmtBRL(metaMensal)}</div>
    </div>
    <div class="card ${gap > 0 ? 'red' : 'green'}">
      <div class="label">Gap Restante</div>
      <div class="value">${fmtBRL(gap)}</div>
      <div class="sub">${mesesRestantes > 0
        ? `Faltam ${mesesRestantes} meses` : 'Meta atingida 🎉'}</div>
    </div>
    <div class="card orange">
      <div class="label">Meta p/ Mês Restante</div>
      <div class="value">${fmtBRL(gapPorMes)}</div>
      <div class="sub">${mesesRestantes} meses restantes</div>
    </div>`;

  const labels    = dados.map(d => d.mes);
  const recebidos = dados.map(d => d.recebido);
  const metas     = dados.map(d => d.meta);

  const necessarioMes = mesesRestantes > 0 ? gap / mesesRestantes : 0;
  const projetado = dados.map(d =>
    (d.recebido === 0 && d.mesIdx > ultimoMesIdx) ? necessarioMes : null
  );

  if(stateBoletim.chartBarras) stateBoletim.chartBarras.destroy();
  stateBoletim.chartBarras = new Chart(
    document.getElementById('chartBoletim'),
    {
      data:{
        labels,
        datasets:[
          {
            type:'bar', label:'Recebido', data:recebidos,
            backgroundColor:'rgba(77,163,255,.75)',
            borderRadius:6, maxBarThickness:46,
            grouped: false, barPercentage: 0.7, categoryPercentage: 0.8,
            order: 3,
            datalabels:{ color:'#4da3ff' }   /* só sobrescreve a cor */
          },
          {
            type:'bar', label:'Necessário p/ fechar gap',
            data:projetado,
            backgroundColor:'rgba(129, 65, 182, 0.85)',
            borderRadius:6, maxBarThickness:46,
            grouped: false, barPercentage: 0.7, categoryPercentage: 0.8,
            order: 2,
            datalabels:{ color:'#a047f3' }
          },
          {
            type:'line', label:'Meta Mensal', data:metas,
            borderColor:'#ef4444', borderWidth:2, tension:.25,
            pointBackgroundColor:'#ef4444', pointRadius:4,
            order: 1,
            datalabels:{ display:false }     /* esconde na linha */
          }
        ]
      },
      options: chartBaseOptions()
    }
  );
}


/* 6. SEÇÃO CUSTEIO */
const stateCusteio = {
  dados: [],
  chartBarras: null
};

async function initCusteio(){
  try{
    const rows = await carregarExcel(ARQUIVO_CUSTEIO);
    if(!rows.length){
      mostrarErro('custeioStatus', 'CUSTEIO_2026.xlsx está vazio.');
      return;
    }

    const headers    = Object.keys(rows[0]);
    const colSetor   = acharColuna(headers, 'segmento', 'setor', 'equipe', 'time', 'area');
    const colPrev    = acharColuna(headers, 'previsto', 'valorprevisto', 'apontado',
                                   'validado', 'planejado');
    const colPago    = acharColuna(headers, 'recebido', 'pago', 'valorpago');
    const colNaoPago = acharColuna(headers, 'naopago', 'nao pago', 'naopag',
                                   'pendente', 'naorecebido');
    const colExtra   = acharColuna(headers, 'extra', 'ordensextras', 'ordemextra',
                                   'fora', 'pagoamais', 'amais');
    const colMes     = acharColuna(headers, 'mes', 'mês', 'month', 'periodo', 'data');

    if(!colSetor || !colPago){
      mostrarErro('custeioStatus',
        'Colunas de Setor ou Valor Pago não encontradas. Colunas: ' +
        headers.join(' | '));
      return;
    }

    stateCusteio.dados = rows.map(r => {
      const mesIdx   = colMes ? normalizarMes(r[colMes]) : null;
      const previsto = colPrev ? toNumber(r[colPrev]) : 0;
      const pago     = toNumber(r[colPago]);
      let   naoPago  = colNaoPago ? toNumber(r[colNaoPago]) : 0;
      if(!colNaoPago && previsto) naoPago = Math.max(0, previsto - pago);
      const extra    = colExtra ? toNumber(r[colExtra]) : 0;

      return {
        setor: String(r[colSetor] || 'Sem setor').trim(),
        mesIdx,
        mes: mesIdx !== null ? MESES[mesIdx] : '—',
        previsto, pago, naoPago, extra
      };
    });

    const setores = [...new Set(stateCusteio.dados.map(x => x.setor))].sort();
    const selSetor = document.getElementById('cFilterSetor');
    selSetor.innerHTML = '<option value="">Todos os setores</option>' +
      setores.map(s => `<option value="${s}">${s}</option>`).join('');
    selSetor.onchange = renderCusteio;

    const mesesDisp = [...new Set(stateCusteio.dados
      .filter(x => x.mesIdx !== null).map(x => x.mesIdx))].sort((a,b) => a-b);
    const selMes = document.getElementById('cFilterMes');
    selMes.innerHTML = '<option value="">Todos os meses</option>' +
      mesesDisp.map(m => `<option value="${m}">${MESES[m]}</option>`).join('');
    selMes.onchange = renderCusteio;

    renderCusteio();
  }catch(e){
    mostrarErro('custeioStatus', 'Erro ao carregar Custeio: ' + e.message);
  }
}

function renderCusteio(){
  const fs = document.getElementById('cFilterSetor').value;
  const fm = document.getElementById('cFilterMes').value;
  let dados = stateCusteio.dados;
  if(fs)        dados = dados.filter(d => d.setor === fs);
  if(fm !== '') dados = dados.filter(d => d.mesIdx === parseInt(fm));

  const totalPrev    = dados.reduce((s,x) => s + x.previsto, 0);
  const totalPago    = dados.reduce((s,x) => s + x.pago, 0);
  const totalNaoPago = dados.reduce((s,x) => s + x.naoPago, 0);
  const totalExtra   = dados.reduce((s,x) => s + x.extra, 0);

  document.getElementById('custeioCards').innerHTML = `
    <div class="card">
      <div class="label">Total Validado</div>
      <div class="value">${fmtBRL(totalPrev)}</div>
      <div class="sub">${fmtNum(dados.length)} registros</div>
    </div>
    <div class="card green">
      <div class="label">Total Pago</div>
      <div class="value">${fmtBRL(totalPago)}</div>
      <div class="sub">${totalPrev ? fmtPct(totalPago/totalPrev*100) : '0%'} do validado</div>
    </div>
    <div class="card red">
      <div class="label">Não Pago</div>
      <div class="value">${fmtBRL(totalNaoPago)}</div>
      <div class="sub">${totalPrev ? fmtPct(totalNaoPago/totalPrev*100) : '0%'} do validado</div>
    </div>
    <div class="card purple">
      <div class="label">Ordens Extras</div>
      <div class="value">${fmtBRL(totalExtra)}</div>
      <div class="sub">Valores pagos a mais</div>
    </div>`;

  const pct  = totalPrev > 0 ? (totalPago / totalPrev * 100) : 0;
  const circ = 2 * Math.PI * 70;
  const ring = document.getElementById('pctRing');
  ring.setAttribute('stroke-dasharray', circ);
  ring.setAttribute('stroke-dashoffset', circ * (1 - Math.min(pct,100)/100));
  ring.setAttribute('stroke',
    pct >= 90 ? '#22c55e' : pct >= 70 ? '#f79604' : '#ef4444');

  document.getElementById('pctVal').textContent    = fmtPct(pct);
  document.getElementById('pctDetail').textContent =
    `${fmtBRL(totalPago)} pagos de ${fmtBRL(totalPrev)} validados`;

  const porMes = {};
  dados.forEach(d => {
    const key = d.mesIdx !== null ? d.mesIdx : -1;
    if(!porMes[key]) porMes[key] = { prev:0, pago:0 };
    porMes[key].prev += d.previsto;
    porMes[key].pago += d.pago;
  });
  const chaves  = Object.keys(porMes).map(Number).sort((a,b) => a - b);
  const labels  = chaves.map(k => k >= 0 ? MESES_ABREV[k] : '—');
  const valPrev = chaves.map(k => porMes[k].prev);
  const valPago = chaves.map(k => porMes[k].pago);

  if(stateCusteio.chartBarras) stateCusteio.chartBarras.destroy();
  stateCusteio.chartBarras = new Chart(
    document.getElementById('chartCusteio'),
    {
      type:'bar',
      data:{
        labels,
        datasets:[
          {
            label:'Validado', data:valPrev,
            backgroundColor:'rgba(77,163,255,.75)',
            borderRadius:6, maxBarThickness:36,
            datalabels:{ color:'#4da3ff' }   /* só cor */
          },
          {
            label:'Pago', data:valPago,
            backgroundColor:'rgba(34,197,94,.8)',
            borderRadius:6, maxBarThickness:36,
            datalabels:{ color:'#22c55e' }   /* só cor */
          }
        ]
      },
      options: chartBaseOptions()
    }
  );

  const agrupar = campo => {
    const m = {};
    dados.forEach(d => { m[d.setor] = (m[d.setor] || 0) + d[campo]; });
    const total = Object.values(m).reduce((s,v) => s + v, 0);
    return { m, total };
  };

  const renderTabela = (elId, campo, cor) => {
    const { m, total } = agrupar(campo);
    const linhas = Object.entries(m)
      .map(([setor,val]) => ({
        setor, val,
        pct: total > 0 ? val / total * 100 : 0
      }))
      .sort((a,b) => b.val - a.val);

    document.getElementById(elId).innerHTML = `
      <thead>
        <tr>
          <th>Setor</th>
          <th class="num">Valor</th>
          <th class="num">Participação</th>
        </tr>
      </thead>
      <tbody>
        ${linhas.length
          ? linhas.map(l => `
              <tr>
                <td>${l.setor}</td>
                <td class="num">${fmtBRL(l.val)}</td>
                <td class="num">
                  <span class="bar-inline"
                    style="width:${Math.min(l.pct * 1.6, 80)}px;background:${cor}"></span>
                  ${fmtPct(l.pct)}
                </td>
              </tr>`).join('')
          : `<tr><td colspan="3"
                style="text-align:center;color:var(--muted);padding:20px">
                Sem dados</td></tr>`}
      </tbody>
      <tfoot>
        <tr style="background:#182238;font-weight:700">
          <td>Total</td>
          <td class="num">${fmtBRL(total)}</td>
          <td class="num">100%</td>
        </tr>
      </tfoot>`;
  };

  renderTabela('tblPago',    'pago',    '#22c55e');
  renderTabela('tblNaoPago', 'naoPago', '#ef4444');
  renderTabela('tblExtras',  'extra',   '#a855f7');
}


/* 7. INICIALIZAÇÃO */
async function reloadAll(){
  if(stateBoletim.chartBarras) stateBoletim.chartBarras.destroy();
  stateBoletim.chartBarras = null;

  if(stateCusteio.chartBarras) stateCusteio.chartBarras.destroy();
  stateCusteio.chartBarras = null;

  await Promise.all([ initBoletim(), initCusteio() ]);

  document.getElementById('updateInfo').textContent =
    'Atualizado em ' + new Date().toLocaleString('pt-BR');
}

document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('btnReload').addEventListener('click', reloadAll);
  reloadAll();
});
