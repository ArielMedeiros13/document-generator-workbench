/* ══════════════════════════════════════════════════════════════
   DOCUMENT GENERATOR WORKBENCH · JS Engine v3
   Módulos: Importação PDF · Extração · Mapeamento Heurístico ·
            Revisão · Templates por Tipo · Geração · Exportação
   ══════════════════════════════════════════════════════════════ */

/* ── Versão para diagnóstico de cache ───────────────────────────
   Confirme no console: window.DOC_WORKBENCH_VERSION
   Se retornar undefined ou valor antigo → browser está servindo
   uma versão cacheada do arquivo. Force Ctrl+Shift+R ou adicione
   ?v= diferente na tag <script src="document-generator-workbench.js?v=...">
   ────────────────────────────────────────────────────────────── */
window.DOC_WORKBENCH_VERSION = '2026-03-13.02';

/* ══════════════════════════════════════════════════════════════
   MÓDULO: CARREGAMENTO DO PDF.JS — robusto, com fallback de CDN
   ══════════════════════════════════════════════════════════════
   Estratégia:
   1. Tenta CDN primário (cdnjs)
   2. Se falhar em 8s, tenta CDN secundário (jsdelivr)
   3. Se ambos falharem, sinaliza claramente: import fica bloqueado
      com mensagem específica, sem crash silencioso.

   Por que dinâmico em vez de <script> estático:
   - <script integrity="hash"> descarta o arquivo inteiro se o hash
     não bater, sem erro visível → pdfjsLib fica undefined silenciosamente.
   - Carregamento dinâmico detecta o erro e tenta fallback.
   - Permite configurar workerSrc do mesmo CDN que carregou.
   ══════════════════════════════════════════════════════════════ */

const PDFJS_VERSION = '3.11.174';
const PDFJS_CDNS = [
  {
    lib:    `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${PDFJS_VERSION}/pdf.min.js`,
    worker: `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${PDFJS_VERSION}/pdf.worker.min.js`,
    name:   'cdnjs',
  },
  {
    lib:    `https://cdn.jsdelivr.net/npm/pdfjs-dist@${PDFJS_VERSION}/build/pdf.min.js`,
    worker: `https://cdn.jsdelivr.net/npm/pdfjs-dist@${PDFJS_VERSION}/build/pdf.worker.min.js`,
    name:   'jsdelivr',
  },
];

// Estado global do carregamento — consultado por handlePdfFile
window._gudiPdfJs = {
  status: 'loading',   // 'loading' | 'ready' | 'failed'
  workerSrc: null,
  error: null,
};

function _loadScriptWithTimeout(src, timeoutMs) {
  return new Promise((resolve, reject) => {
    const el = document.createElement('script');
    el.src = src;
    el.crossOrigin = 'anonymous';
    const timer = setTimeout(() => {
      el.remove();
      reject(new Error(`timeout após ${timeoutMs}ms`));
    }, timeoutMs);
    el.onload = () => { clearTimeout(timer); resolve(); };
    el.onerror = (e) => { clearTimeout(timer); el.remove(); reject(e); };
    document.head.appendChild(el);
  });
}

async function _initPdfJs() {
  for (const cdn of PDFJS_CDNS) {
    try {
      console.info(`[DOCS] Tentando carregar pdf.js via ${cdn.name}…`);
      await _loadScriptWithTimeout(cdn.lib, 8000);
      if (typeof window.pdfjsLib === 'undefined') {
        throw new Error('script carregado mas pdfjsLib não definido');
      }
      pdfjsLib.GlobalWorkerOptions.workerSrc = cdn.worker;
      window._gudiPdfJs.status = 'ready';
      window._gudiPdfJs.workerSrc = cdn.worker;
      console.info(`[DOCS] pdf.js carregado via ${cdn.name} (v${pdfjsLib.version})`);
      return;
    } catch (err) {
      console.warn(`[DOCS] pdf.js falhou via ${cdn.name}:`, err.message || err);
    }
  }
  // Todos os CDNs falharam
  window._gudiPdfJs.status = 'failed';
  window._gudiPdfJs.error = 'Nenhum CDN disponível';
  console.error('[DOCS] pdf.js não carregou em nenhum CDN. Import de PDF indisponível.');
}

// Iniciar carregamento imediatamente (não bloqueia o resto do JS)
const _pdfJsReady = _initPdfJs();

/* ──────────────────────────────────────────────────────────────
   MÓDULO: STATE
   ────────────────────────────────────────────────────────────── */
const state = {
  docType: 'relatorio',
  theme: 'grafite',
  sections: [],
  highlights: [],
  documentModel: [],
  structureMode: 'editorial',
  sectionIdCounter: 0,
  highlightIdCounter: 0,
  reviewBlocks: [],
  reviewBlockCounter: 0,
};

/* ──────────────────────────────────────────────────────────────
   MÓDULO: METADADOS DE TIPO DOCUMENTAL
   ────────────────────────────────────────────────────────────── */
const DOC_TYPE_LABELS = {
  relatorio:   'RELATÓRIO EXECUTIVO',
  parecer:     'PARECER INSTITUCIONAL',
  diagnostico: 'DIAGNÓSTICO ESTRATÉGICO',
  proposta:    'PROPOSTA',
  sintese:     'SÍNTESE ESTRATÉGICA',
  impacto:     'RELATÓRIO DE IMPACTO',
  crm:         'RELATÓRIO DE CRM — INTELIGÊNCIA PÓS-EVENTO',
};

/* Estrutura editorial sugerida por tipo documental */
const DOC_TYPE_SCHEMA = {
  relatorio: {
    summaryLabel: 'Resumo Executivo',
    nextLabel: 'Próximos Passos',
    defaultSections: ['Contexto e Objetivos', 'Análise de Resultados', 'Conclusões'],
    highlightsLabel: 'Indicadores-Chave',
    intro: null,
    signature: false,
  },
  parecer: {
    summaryLabel: 'Fundamentação',
    nextLabel: 'Recomendações',
    defaultSections: ['Objeto do Parecer', 'Análise', 'Conclusão e Parecer Final'],
    highlightsLabel: 'Pontos Avaliados',
    intro: 'Este parecer foi elaborado com base nas informações disponibilizadas e tem caráter institucional. As conclusões aqui apresentadas são de responsabilidade da equipe signatária.',
    signature: true,
  },
  diagnostico: {
    summaryLabel: 'Síntese Diagnóstica',
    nextLabel: 'Recomendações',
    defaultSections: ['Contexto e Metodologia', 'Achados Principais', 'Análise Crítica', 'Leitura Estratégica'],
    highlightsLabel: 'Indicadores Diagnósticos',
    intro: null,
    signature: false,
  },
  proposta: {
    summaryLabel: 'Visão Geral da Proposta',
    nextLabel: 'Próximos Passos para Aprovação',
    defaultSections: ['Objetivo e Escopo', 'Entregáveis', 'Cronograma', 'Investimento'],
    highlightsLabel: 'Resumo da Proposta',
    intro: null,
    signature: false,
  },
  sintese: {
    summaryLabel: 'Síntese Executiva',
    nextLabel: 'Ações Prioritárias',
    defaultSections: ['Situação Atual', 'Análise Estratégica', 'Decisões Recomendadas'],
    highlightsLabel: 'Fatores Críticos',
    intro: null,
    signature: false,
  },
  impacto: {
    summaryLabel: 'Sumário de Impacto',
    nextLabel: 'Próximos Ciclos',
    defaultSections: ['Contexto do Projeto', 'Resultados Alcançados', 'Repercussão Institucional', 'Lições Aprendidas'],
    highlightsLabel: 'Indicadores de Impacto',
    intro: null,
    signature: false,
  },
  crm: {
    summaryLabel: 'Panorama Geral — Inteligência CRM',
    nextLabel: 'Ações de Reengajamento',
    defaultSections: [
      'Ausentes Totais — 6 Temperaturas',
      'Segmentos de Maior Valor Estratégico',
      'Hipóteses sobre Comparecimento',
      'Públicos Prioritários — Próxima Edição',
      'Proposta de Segmentação CRM',
      'Recomendações de Próximos Movimentos',
    ],
    highlightsLabel: 'KPIs Analíticos',
    intro: null,
    signature: false,
  },
};

const THEME_DESCS = {
  grafite:  '<strong>Grafite</strong> — fundo grafite escuro, contraste máximo, lateral colorida. Para relatórios executivos e materiais de impacto interno.',
  editorial:'<strong>Editorial</strong> — fundo branco, tipografia refinada, linha laranja de acento. Para documentos formais de circulação externa.',
  navy:     '<strong>Navy</strong> — azul institucional profundo, identidade corporativa e governamental. Para pareceres e documentos sérios.',
  pdf:      '<strong>Modo PDF</strong> — language inspired by polished presentation decks: bold type, gradient bar and decorative accents.',
};

const BULLET_COLORS = ['#E16E1A','#2ED9C3','#81C458','#E5B92B','#2D4F76'];

const RE_TABULAR_SEP     = / {2,}|\t|[|;]/;
const RE_TABULAR_SPLIT   = / {2,}|\t|\s+\|\s+|;/;
const MAX_X_GAP_FOR_LINE = 8; // Reduzido para evitar merge de colunas distintas
const MAX_Y_GAP_FOR_LINE = 2; // Reduzido para evitar merge de linhas separadas
const MIN_FONT_SIZE_RATIO_FOR_HEADING = 1.3; // Aumentado para evitar falsos headings
const MAX_FONT_SIZE_VARIATION_IN_LINE = 1.15; // Reduzido para evitar merge de tamanhos diferentes
const RE_SLASH_SEP       = /\s+\/\s+/;
const RE_DASH_SEP        = /\s+[—–]{1,2}\s+/;

/* ──────────────────────────────────────────────────────────────
   MÓDULO: LOGO INLINE (resolve bug de imagem quebrada no print)
   Em vez de <img src="...">, usamos um logotipo tipográfico SVG
   inline que sempre renderiza corretamente em qualquer contexto.
   ────────────────────────────────────────────────────────────── */
function buildLogoSVG(variant = 'light', size = 'header') {
  // variant: 'light' = texto branco (fundos escuros)
  //          'dark'  = texto laranja (fundos claros)
  // size: 'header' | 'footer'
  const h = size === 'footer' ? '17' : '26';
  const colors = variant === 'light'
    ? ['#81C458','#2ED9C3','#2D4F76','#E16E1A']
    : ['#E16E1A','#E5B92B','#2D4F76','#2ED9C3'];

  return `<svg height="${h}" viewBox="0 0 72 ${h === '17' ? '17' : '26'}" xmlns="http://www.w3.org/2000/svg" class="doc-header__logo-svg" aria-label="DOCS">
    <text y="${h === '17' ? '13' : '20'}" font-family="'Nunito',sans-serif" font-weight="900" font-size="${h === '17' ? '14' : '22'}" letter-spacing="2" fill="url(#lgrd-${variant}-${size})">DOCS</text>
    <defs>
      <linearGradient id="lgrd-${variant}-${size}" x1="0%" y1="0%" x2="100%" y2="0%">
        <stop offset="0%"   stop-color="${colors[0]}"/>
        <stop offset="33%"  stop-color="${colors[1]}"/>
        <stop offset="66%"  stop-color="${colors[2]}"/>
        <stop offset="100%" stop-color="${colors[3]}"/>
      </linearGradient>
    </defs>
  </svg>`;
}

function logoVariant(theme) {
  return (theme === 'editorial' || theme === 'pdf') ? 'dark' : 'light';
}

/* ──────────────────────────────────────────────────────────────
   MÓDULO: ASSETS EXTERNOS (ornamentos)
   ────────────────────────────────────────────────────────────── */
const ASSETS = {
  degrade06: 'degradês-06.svg',
  degrade07: 'degradês-07.svg',
};

/* Ornamento SVG inline (fallback caso asset não carregue) */
const SVG_STAR_ORNAMENT = `<svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg" class="doc-ornament doc-ornament--star" aria-hidden="true">
  <polygon points="50,0 58,35 95,38 68,60 78,95 50,73 22,95 32,60 5,38 42,35" fill="url(#grd-orn-s)"/>
  <defs><linearGradient id="grd-orn-s" x1="0%" y1="0%" x2="100%" y2="100%">
    <stop offset="0%" stop-color="#81C458"/><stop offset="50%" stop-color="#2ED9C3"/><stop offset="100%" stop-color="#2D4F76"/>
  </linearGradient></defs>
</svg>`;

const SVG_SWIRL_ORNAMENT = `<svg viewBox="0 0 200 120" xmlns="http://www.w3.org/2000/svg" class="doc-ornament doc-ornament--swirl" aria-hidden="true">
  <path d="M10,80 C10,20 70,5 100,40 C130,75 170,20 190,50" fill="none" stroke="url(#grd-orn-sw)" stroke-width="16" stroke-linecap="round"/>
  <defs><linearGradient id="grd-orn-sw" x1="0%" y1="0%" x2="100%" y2="0%">
    <stop offset="0%" stop-color="#E16E1A"/><stop offset="50%" stop-color="#81C458"/><stop offset="100%" stop-color="#2ED9C3"/>
  </linearGradient></defs>
</svg>`;

/* ──────────────────────────────────────────────────────────────
   DOM REFS
   ────────────────────────────────────────────────────────────── */
const $ = id => document.getElementById(id);

const els = {
  docTypes:        $('docTypes'),
  themeOpts:       $('themeOpts'),
  themeDesc:       $('themeDesc'),
  fldTitle:        $('fldTitle'),
  fldSubtitle:     $('fldSubtitle'),
  fldDate:         $('fldDate'),
  fldVersion:      $('fldVersion'),
  fldAuthor:       $('fldAuthor'),
  fldDept:         $('fldDept'),
  fldSummary:      $('fldSummary'),
  fldNextSteps:    $('fldNextSteps'),
  sectionsList:    $('sectionsList'),
  highlightsList:  $('highlightsList'),
  addSectionBtn:   $('addSectionBtn'),
  addHighlightBtn: $('addHighlightBtn'),
  generateBtn:     $('generateBtn'),
  exampleBtn:      $('exampleBtn'),
  printBtn:        $('printBtn'),
  stageHint:       $('stageHint'),
  docWrap:         $('docWrap'),
  // Import
  importPdfBtn:    $('importPdfBtn'),
  pdfFileInput:    $('pdfFileInput'),
  importZone:      $('importZone'),
  dropZoneGlobal:  $('dropZoneGlobal'),
  loadingOverlay:  $('loadingOverlay'),
  loadingLabel:    $('loadingLabel'),
  // Review modal
  reviewOverlay:   $('reviewOverlay'),
  reviewClose:     $('reviewClose'),
  reviewCancel:    $('reviewCancel'),
  reviewApply:     $('reviewApply'),
  reviewAddSec:    $('reviewAddSec'),
  reviewSectionsList: $('reviewSectionsList'),
  reviewSubtitle:  $('reviewSubtitle'),
  reviewWarning:   $('reviewWarning'),
  reviewWarningText: $('reviewWarningText'),
  reviewPageInfo:  $('reviewPageInfo'),
  rv_docType:  $('rv_docType'),
  rv_structureMode: $('rv_structureMode'),
  rv_title:    $('rv_title'),
  rv_subtitle: $('rv_subtitle'),
  rv_date:     $('rv_date'),
  rv_version:  $('rv_version'),
  rv_author:   $('rv_author'),
  rv_dept:     $('rv_dept'),
  rv_summary:  $('rv_summary'),
  rv_nextSteps:$('rv_nextSteps'),
  rv_raw:      $('rv_raw'),
};

/* ══════════════════════════════════════════════════════════════
   MÓDULO: INIT — liga todos os eventos
   ══════════════════════════════════════════════════════════════ */
function init() {

  /* ── Diagnóstico de ambiente — visível no console do browser ── */
  console.info(
    `%c[DOCS] JS carregado · versão ${window.DOC_WORKBENCH_VERSION}`,
    'color:#2ED9C3;font-weight:700'
  );
  // pdf.js pode ainda estar carregando de forma assíncrona —
  // o status real é consultado em handlePdfFile via _pdfJsReady
  console.info(`[DOCS] pdf.js status inicial: ${window._gudiPdfJs?.status}`);
  console.info(`[DOCS] setupDragDropZone: ${typeof setupDragDropZone}`);
  if (typeof setupDragDropZone === 'undefined') {
    console.error('[DOCS] CRÍTICO: setupDragDropZone não definida — versão errada do JS em execução!');
  }

  /* pdf.js worker já configurado por _initPdfJs() — não reconfigurar aqui */

  /* ── Tipo de documento ──────────────────────────────────────── */
  els.docTypes.addEventListener('click', e => {
    const btn = e.target.closest('.doc-type');
    if (!btn) return;
    document.querySelectorAll('.doc-type').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    state.docType = btn.dataset.type;
    if (els.docWrap.style.display !== 'none') generateDocument();
  });

  /* ── Tema visual ──────────────────────────────────────────── */
  els.themeOpts.addEventListener('click', e => {
    const btn = e.target.closest('.theme-card');
    if (!btn) return;
    document.querySelectorAll('.theme-card').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    state.theme = btn.dataset.theme;
    if (els.themeDesc) els.themeDesc.innerHTML = THEME_DESCS[state.theme] || '';
    // Aplica instantaneamente sem regenerar
    const docEl = els.docWrap.querySelector('.gudi-doc');
    if (docEl) {
      docEl.dataset.theme = state.theme;
      // Atualiza logos inline (eles são SVG, só troca a variante)
      docEl.querySelectorAll('[data-logo-variant]').forEach(el => {
        const sz = el.dataset.logoSize || 'header';
        el.outerHTML = buildLogoSVG(logoVariant(state.theme), sz);
      });
    }
  });

  /* ── Seções e indicadores ──────────────────────────────────── */
  els.addSectionBtn.addEventListener('click', () => addSection());
  els.addHighlightBtn.addEventListener('click', () => addHighlight());

  /* ── Gerar ─────────────────────────────────────────────────── */
  els.generateBtn.addEventListener('click', generateDocument);

  /* ── Exemplo ───────────────────────────────────────────────── */
  els.exampleBtn.addEventListener('click', loadExample);

  /* ── Exportar PDF — versão robusta (preserva cores no print) ── */
  els.printBtn.addEventListener('click', async () => {
    const btn = els.printBtn;
    const orig = btn.innerHTML;
    btn.innerHTML = 'Preparando…';
    btn.disabled = true;
    preparePrintMode();
    try {
      if (document.fonts?.ready) await document.fonts.ready;
      await waitForAssetsReady(els.docWrap);
      await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
      window.print();
    } finally {
      setTimeout(() => {
        btn.innerHTML = orig;
        btn.disabled = false;
        cleanupPrintMode();
      }, 300);
    }
  });

  /* ── Import PDF: botão no header ──────────────────────────── */
  els.importPdfBtn.addEventListener('click', () => els.pdfFileInput.click());

  /* ── Import PDF: zona no painel ─────────────────────────────── */
  els.importZone.addEventListener('click', () => els.pdfFileInput.click());

  /* ── Input de arquivo ──────────────────────────────────────── */
  els.pdfFileInput.addEventListener('change', e => {
    const file = e.target.files[0];
    if (file) handlePdfFile(file);
    e.target.value = ''; // reset para permitir re-upload do mesmo arquivo
  });

  /* ── Drag and drop: zona do painel ─────────────────────────── */
  setupDragDropZone(els.importZone);
  runAutotestFromQuery();

  /* ── Drag and drop: global (body) ──────────────────────────── */
  let dragCount = 0;
  document.addEventListener('dragenter', e => {
    if (e.dataTransfer.types.includes('Files')) {
      dragCount++;
      els.dropZoneGlobal.style.display = 'flex';
    }
  });
  document.addEventListener('dragleave', () => {
    dragCount--;
    if (dragCount <= 0) { dragCount = 0; els.dropZoneGlobal.style.display = 'none'; }
  });
  document.addEventListener('dragover', e => e.preventDefault());
  document.addEventListener('drop', e => {
    e.preventDefault();
    dragCount = 0;
    els.dropZoneGlobal.style.display = 'none';
    const file = e.dataTransfer.files[0];
    if (file && file.type === 'application/pdf') {
      handlePdfFile(file);
    } else if (file) {
      showWarning('Apenas arquivos PDF são suportados para importação.');
    }
  });

  /* ── Review modal ────────────────────────────────────────── */
  els.reviewClose.addEventListener('click', closeReview);
  els.reviewCancel.addEventListener('click', closeReview);
  els.reviewApply.addEventListener('click', applyReviewToTemplate);
  els.reviewOverlay.addEventListener('click', e => {
    if (e.target === els.reviewOverlay) closeReview();
  });
  els.reviewAddSec.addEventListener('click', () => addReviewBlock({ type: 'generic', confidence: 'low', title: '', text: '', enabled: true }));

  window.addEventListener('beforeprint', preparePrintMode);
  window.addEventListener('afterprint', cleanupPrintMode);

  const params = new URLSearchParams(window.location.search);
  if (params.get('example') === '1') {
    loadExample();
  }
}

function setupDragDropZone(zone) {
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('import-zone--active'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('import-zone--active'));
  zone.addEventListener('drop', e => {
    e.preventDefault(); zone.classList.remove('import-zone--active');
    const file = e.dataTransfer.files[0];
    if (file && file.type === 'application/pdf') handlePdfFile(file);
  });
}

/* ══════════════════════════════════════════════════════════════
   MÓDULO: IMPORTAÇÃO E EXTRAÇÃO DE PDF
   Usa pdf.js para extração real de texto página por página
   ══════════════════════════════════════════════════════════════ */
async function handlePdfFile(file) {
  if (!file || file.type !== 'application/pdf') {
    showWarning('O arquivo selecionado não é um PDF válido.');
    return;
  }

  showLoading('Carregando leitor de PDF…');

  // ── Aguardar o carregamento dinâmico do pdf.js ─────────────────────────
  // _pdfJsReady é a promise de _initPdfJs() lançada no topo do arquivo.
  // Se o usuário clicou em importar logo após a página carregar,
  // o CDN pode ainda não ter respondido — aguardamos aqui sem travar a UI.
  try { await _pdfJsReady; } catch (_) { /* _initPdfJs nunca rejeita */ }

  // ── Verificar resultado ────────────────────────────────────────────────
  if (window._gudiPdfJs?.status !== 'ready' || !window.pdfjsLib) {
    hideLoading();
    const gudiStatus = window._gudiPdfJs?.status || 'desconhecido';
    console.error(`[DOCS] pdf.js indisponível. Status: ${gudiStatus}`);

    const isEdge   = /Edg\//.test(navigator.userAgent);
    const isFF     = /Firefox\//.test(navigator.userAgent);
    let msg;
    if (gudiStatus === 'failed') {
      msg = 'O leitor de PDF não carregou em nenhum servidor. '
          + 'Verifique sua conexão e recarregue (Ctrl+Shift+R). '
          + 'Redes corporativas podem bloquear CDNs externos.';
    } else if (isEdge) {
      msg = 'O leitor de PDF não carregou. O Edge pode estar bloqueando CDNs '
          + '(Tracking Prevention). Tente: Configurações → Privacidade → '
          + 'Prevenção de rastreamento → Básico. Ou abra em aba normal.';
    } else if (isFF) {
      msg = 'O leitor de PDF não carregou. O Firefox pode bloquear CDNs em modo Privativo. '
          + 'Abra em aba normal ou desative o bloqueio para este site.';
    } else {
      msg = 'O leitor de PDF não carregou. Recarregue a página e tente novamente. '
          + 'Se persistir, tente em outro navegador.';
    }
    showWarning(msg);
    return;
  }

  setLoading('Lendo o arquivo PDF…');

  try {
    const arrayBuffer = await file.arrayBuffer();
    setLoading('Extraindo texto do PDF…');

    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    const totalPages = pdf.numPages;
    setLoading(`Processando ${totalPages} página(s)…`);

    let fullText = '';
    const pageTexts = [];
    const pageStructures = [];

    for (let i = 1; i <= totalPages; i++) {
      setLoading(`Extraindo página ${i} de ${totalPages}…`);
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();

      const pageStructure = reconstructPageStructure(content.items, i);
      const pageText = pageStructure.blocks.map(block => block.text).join('\n\n');
      pageStructures.push(pageStructure);
      pageTexts.push(pageText);
      fullText += pageText + '\n\n';
    }

    hideLoading();

    // ── Qualidade da extração ──────────────────────────────────────────
    const charCount = fullText.trim().length;
    let warning = null;
    if (charCount === 0) {
      warning = 'Nenhum texto extraído. O PDF provavelmente é uma imagem escaneada '
              + 'ou está protegido. Preencha os campos manualmente.';
    } else if (charCount < 50) {
      warning = 'Texto muito escasso. O PDF pode ser escaneado ou protegido. '
              + 'Revise e complete os campos antes de aplicar.';
    } else if (charCount < 300) {
      warning = 'Texto curto extraído. Pode haver conteúdo predominantemente visual. '
              + 'Revise o mapeamento antes de aplicar.';
    }

    const mapped = heuristicMap({
      text: fullText,
      pageTexts,
      pageStructures,
      filename: file.name,
      selectedDocType: state.docType,
    });

    openReview(mapped, fullText, warning, totalPages);

  } catch (err) {
    hideLoading();
    console.error('[DOCS] Erro durante extração de PDF:', err);

    let userMsg;
    const errMsg = (err.message || '').toLowerCase();
    if (/password|encrypted/.test(errMsg)) {
      userMsg = 'O PDF está protegido por senha. Remova a proteção antes de importar.';
    } else if (err.name === 'UnknownErrorException' || /worker.*crash|worker.*terminat/.test(errMsg)) {
      userMsg = 'O processador de PDF travou. Recarregue a página e tente novamente.';
    } else if (/invalid pdf|missing pdf|not a pdf|unexpected end/.test(errMsg)) {
      userMsg = 'O arquivo não é um PDF válido ou está corrompido.';
    } else if (/network|fetch|failed to fetch|load failed/.test(errMsg)) {
      userMsg = 'Falha de rede durante o processamento. Verifique sua conexão.';
    } else if (/not defined|referenceerror/.test(errMsg)) {
      userMsg = 'Erro interno: função não reconhecida. Force a recarga com Ctrl+Shift+R.';
    } else {
      userMsg = `Erro ao processar o PDF: ${err.message || 'erro desconhecido'}. `
              + 'Tente outro arquivo ou preencha os campos manualmente.';
    }

    openReviewWithError(userMsg, file.name);
  }
}



function reconstructPageStructure(items, pageNumber) {
  const lines = reconstructPdfLines(items);
  const blocks = buildEditorialPipeline(lines, pageNumber);
  return { pageNumber, lines, blocks };
}

function reconstructPdfLines(items) {
  if (!items || items.length === 0) return [];

  const filtered = items
    .map(item => ({
      text: cleanPdfToken(item.str || ''),
      x: Number(item.transform?.[4] || 0),
      y: Number(item.transform?.[5] || 0),
      width: Number(item.width || 0),
      height: Math.abs(Number(item.height || item.transform?.[0] || 0)) || 0,
      fontName: item.fontName || '',
      raw: item,
    }))
    .filter(item => item.text);

  const sorted = filtered.sort((a, b) => {
    const dy = Math.abs(b.y - a.y);
    if (dy > 3) return b.y - a.y;
    return a.x - b.x;
  });

  const lines = [];
  for (const token of sorted) {
    const target = lines.find(line => Math.abs(line.y - token.y) <= Math.max(2.5, Math.min(token.height || 0, 6)));
    if (target) {
      target.tokens.push(token);
      target.y = (target.y + token.y) / 2;
      target.maxHeight = Math.max(target.maxHeight, token.height || 0);
    } else {
      lines.push({
        y: token.y,
        maxHeight: token.height || 0,
        tokens: [token],
      });
    }
  }

  return lines
    .sort((a, b) => b.y - a.y)
    .map((line, index, arr) => {
      const tokens = line.tokens.sort((a, b) => a.x - b.x);
      const text = joinPdfLineTokens(tokens);
      const next = arr[index + 1];
      return {
        text,
        y: line.y,
        x: tokens[0]?.x || 0,
        width: Math.max(0, (tokens[tokens.length - 1]?.x || 0) - (tokens[0]?.x || 0)),
        fontSize: median(tokens.map(t => t.height).filter(Boolean)) || 0,
        gapAfter: next ? Math.abs(line.y - next.y) : 0,
        tokens,
      };
    })
    .filter(line => line.text);
}

function cleanPdfToken(text) {
  return String(text || '')
    .replace(/\u00A0/g, ' ')
    .replace(/[ \t]+/g, ' ')
    .trim();
}

function joinPdfLineTokens(tokens) {
  let out = '';
  let prev = null;
  for (const token of tokens) {
    if (!token.text) continue;
    if (!prev) {
      out = token.text;
      prev = token;
      continue;
    }
    const gap = token.x - (prev.x + prev.width);
    const noSpaceBefore = /^[,.;:!?%)\]}]/.test(token.text);
    const noSpaceAfterPrev = /[(\[{\/-]$/.test(prev.text);
    const needsSpace = gap > Math.max(1.5, (prev.height || 0) * 0.14) && !noSpaceBefore && !noSpaceAfterPrev;
    out += `${needsSpace ? ' ' : ''}${token.text}`;
    prev = token;
  }
  return normalizeImportedLine(out);
}

function normalizeImportedLine(text) {
  return String(text || '')
    .replace(/\s+([,.;:!?])/g, '$1')
    .replace(/([(\[{])\s+/g, '$1')
    .replace(/\s+/g, ' ')
    .replace(/-\s+$/g, '-')
    .trim();
}

function sanitizeEditorialLine(text) {
  return normalizeImportedLine(String(text || '')
    .replace(/\b(?:id|uuid|guid|field|label|name|value|type)\s*[:=]\s*[A-Za-z0-9_\-./]+/gi, ' ')
    .replace(/\b(?:id|uuid|guid)\b$/i, ' ')
    .replace(/\s+/g, ' '));
}

function buildEditorialPipeline(lines, pageNumber) {
  const normalized = normalizeEditorialLines(lines);
  const grouped = segmentRawEditorialBlocks(normalized, pageNumber);
  const recomposed = grouped
    .map((block, index) => finalizeEditorialBlock(recomposeHeadingLines(block.lines), index, pageNumber))
    .filter(Boolean);
  const fused = fuseOrphanEditorialBlocks(recomposed);
  return dedupeEditorialBlocks(fused);
}

function normalizeEditorialLines(lines) {
  const output = [];
  for (const line of lines || []) {
    const text = sanitizeEditorialLine(line?.text || '');
    if (!text || isLikelyNoiseLine(text) || isTechnicalLeakLine(text)) continue;
    output.push({
      ...line,
      text,
      lexicalDensity: estimateLexicalDensity(text),
      uppercaseRatio: computeUppercaseRatio(text),
      sentenceLike: isSentenceLike(text),
      tableCandidate: isLikelyTabularLine(text),
      headingCandidate: false,
    });
  }
  const medianFont = median(output.map(item => item.fontSize).filter(Boolean)) || 10;
  return output.map((line, index, arr) => ({
    ...line,
    headingCandidate: isEditorialHeadingLine(line, arr[index + 1], arr[index - 1], medianFont),
  }));
}

function segmentRawEditorialBlocks(lines, pageNumber) {
  if (!lines.length) return [];
  const blocks = [];
  let current = null;
  const bodyFont = median(lines.map(line => line.fontSize).filter(Boolean)) || 10;
  const largeGap = Math.max(18, bodyFont * 1.7);

  lines.forEach((line, index) => {
    const next = lines[index + 1];
    const prevLine = current?.lines?.[current.lines.length - 1] || null;
    const blockBreak = !current
      || current.lastGap >= largeGap
      || shouldStartEditorialBlock(prevLine, line, next, bodyFont);

    if (blockBreak) {
      current = { page: pageNumber, lines: [], lastGap: 0 };
      blocks.push(current);
    }

    current.lines.push(line);
    current.lastGap = line.gapAfter;
  });

  return blocks.filter(block => block.lines.length);
}

function shouldStartEditorialBlock(prevLine, line, nextLine, bodyFont = 10) {
  if (!prevLine || !line) return true;
  const prevText = sanitizeEditorialLine(prevLine.text || '');
  const text = sanitizeEditorialLine(line.text || '');
  if (!prevText || !text) return true;
  if (looksLikeBullet(text) || looksLikeStepLine(text)) return true;
  // Tabular lines stay together with each other
  const currentTabular = isLikelyTabularLine(text) || isLikelyTableHeaderLine(text);
  const prevTabular = isLikelyTabularLine(prevText) || isLikelyTableHeaderLine(prevText);
  if (currentTabular && prevTabular) return false;
  // Mergeable heading continuations stay together
  if (shouldMergeHeadingLines(prevLine, line, bodyFont)) return false;
  // Strong heading (numbered or clearly uppercase non-tabular) starts a new block
  const isHeading = isEditorialHeadingLine(line, nextLine, prevLine, bodyFont);
  if (isHeading && (looksLikeNumberedHeading(text) || (computeUppercaseRatio(text) >= 0.6 && !currentTabular))) return true;
  // Large gap starts new block (except between tabular lines already handled above)
  const gap = Number(prevLine.gapAfter) || 0;
  const largeGapThreshold = Math.max(18, Number(prevLine.fontSize || line.fontSize || bodyFont) * 1.5);
  if (gap >= largeGapThreshold && !(currentTabular && (prevTabular || lineHasHeadingTraits(prevLine, bodyFont, line)))) return true;
  
  return false;
}

function createFallbackStructure(pageLines, pageIndex, issues) {
  // Fallback mais simples: agrupar linhas por proximidade sem tentar ser inteligente
  const blocks = [];
  const allText = pageLines.map(line => sanitizeEditorialLine(line.text || '')).filter(Boolean);
  
  if (allText.length === 0) return [];
  
  // Criar um único bloco genérico com todo o conteúdo
  const combinedText = allText.join('\n');
  
  // Se o texto for muito longo, tentar dividir em parágrafos
  if (combinedText.length > 500) {
    const paragraphs = splitTextIntoParagraphs(combinedText);
    paragraphs.forEach((paragraph, index) => {
      if (paragraph.trim()) {
        blocks.push({
          heading: '',
          text: paragraph,
          lines: [],
          pageNumber: pageIndex + 1,
          extractionQuality: 'poor',
          fallbackMode: true,
          fallbackReason: issues.join(',')
        });
      }
    });
  } else {
    // Texto curto: um único bloco
    blocks.push({
      heading: '',
      text: combinedText,
      lines: pageLines,
      pageNumber: pageIndex + 1,
      extractionQuality: 'poor',
      fallbackMode: true,
      fallbackReason: issues.join(',')
    });
  }
  
  return blocks;
}

function assessPageExtractionQuality(pageText, pageBlocks) {
  if (!pageText || !pageText.trim()) return { quality: 'empty', confidence: 0, issues: ['empty_page'] };
  
  const issues = [];
  let confidence = 1.0;
  
  // Verificar conteúdo muito curto (pode ser ruído)
  if (pageText.length < 50) {
    issues.push('very_short_content');
    confidence -= 0.4;
  }
  
  // Verificar conteúdo fragmentado (muitas linhas curtas)
  const lines = pageText.split('\n').filter(line => line.trim().length > 0);
  const avgLineLength = lines.reduce((sum, line) => sum + line.length, 0) / lines.length;
  if (avgLineLength < 20 && lines.length > 10) {
    issues.push('fragmented_content');
    confidence -= 0.3;
  }
  
  // Verificar proporção de ruído
  const noiseLines = lines.filter(line => isLikelyNoiseLine(line)).length;
  const noiseRatio = noiseLines / lines.length;
  if (noiseRatio > 0.4) {
    issues.push('high_noise_ratio');
    confidence -= 0.3;
  }
  
  // Verificar se há blocos estruturais
  const structuredBlocks = pageBlocks.filter(block => 
    block.type === 'table' || block.type === 'matrix' || block.type === 'comparison'
  ).length;
  if (structuredBlocks === 0 && pageText.length > 500) {
    issues.push('no_structured_blocks');
    confidence -= 0.2;
  }
  
  // Verificar se há headings detectados
  const headingBlocks = pageBlocks.filter(block => 
    block.heading || block.type === 'section'
  ).length;
  if (headingBlocks === 0 && pageText.length > 300) {
    issues.push('no_headings_detected');
    confidence -= 0.1;
  }
  
  // Verificar se o conteúdo parece ser apenas metadata/ruído
  if (looksLikeFooterCluster(pageText) || looksLikeMostlyHeaderFooter(pageBlocks)) {
    issues.push('mostly_metadata');
    confidence -= 0.5;
  }
  
  // Determinar qualidade final
  let quality = 'good';
  if (confidence < 0.3) quality = 'poor';
  else if (confidence < 0.6) quality = 'moderate';
  
  return { quality, confidence: Math.max(0, confidence), issues };
}

function heuristicMap({ text, pageTexts, pageStructures, filename, selectedDocType }) {
  const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
  const blocks = (pageStructures || []).flatMap(page => page.blocks || []);
  const extraction = assessExtractionQuality(text, blocks, pageTexts);
  const result = {
    title: '', subtitle: '', date: '', version: '', author: '', dept: '',
    summary: '', nextSteps: '', docType: selectedDocType || 'relatorio', sections: [],
    highlights: [],
    blocks: [],
    confidence: extraction.confidence,
    extractionNotes: extraction.notes,
  };

  if (lines.length === 0) return result;

  const titleCandidate = findTitleCandidate(blocks, lines, filename);
  const subtitleCandidate = findSubtitleCandidate(blocks, titleCandidate?.text || '');

  result.title = titleCandidate?.confidence === 'high' ? titleCandidate.text : filenameToTitle(filename);
  result.subtitle = subtitleCandidate?.confidence !== 'low' ? subtitleCandidate.text : '';
  result.date = findDate(text);
  result.version = findVersion(text);
  result.author = findAuthor(text);
  result.dept = findDept(text);

  const inferredDocType = inferDocType(text, result.title, blocks);
  result.docType = selectedDocType || inferredDocType;

  const classifiedBlocks = classifyBlocks(blocks, {
    title: result.title,
    subtitle: result.subtitle,
    docType: result.docType,
    text,
  });
  result.blocks = classifiedBlocks;
  result.sections = classifiedBlocks
    .filter(block => block.type === 'section' && block.enabled)
    .slice(0, 12)
    .map(block => ({
      title: block.title || fallbackSectionTitle(block, result.docType),
      body: block.text,
    }));

  result.summary = deriveSummaryFromBlocks(classifiedBlocks).slice(0, 900);
  result.nextSteps = deriveNextStepsFromBlocks(classifiedBlocks).slice(0, 900);

  result.highlights = classifiedBlocks
    .filter(block => block.type === 'highlight' && block.enabled)
    .slice(0, 4)
    .map(block => extractHighlightFromBlock(block))
    .filter(Boolean);

  return result;
}

function classifyBlocks(blocks, context) {
  return (blocks || []).map((block, index) => classifySingleBlock(block, index, context));
}

function classifySingleBlock(block, index, context) {
  const text = (block.text || '').trim();
  const firstLine = block.heading || block.firstLine || text.split('\n')[0] || '';
  const normalizedFirstLine = normalizeHeadingText(firstLine).toLowerCase();
  const lowered = text.toLowerCase();
  let type = 'generic';
  let confidence = 'medium';
  let title = normalizeContentText(block.title || block.heading || '');
  const enabled = block.enabled !== false;

  // Classificação mais conservadora para evitar falsos positivos

  // Summary - critérios mais restritos
  if (normalizedFirstLine.includes('resumo') || normalizedFirstLine.includes('sumário') || 
      normalizedFirstLine.includes('executivo') || normalizedFirstLine.includes('síntese')) {
    if (text.length <= 800 && text.split('\n').length <= 6) {
      type = 'summary';
      confidence = text.length > 100 ? 'high' : 'medium';
    }
  }

  // Next Steps - critérios mais restritos
  else if (normalizedFirstLine.includes('próximos passos') || normalizedFirstLine.includes('próximo') || 
           normalizedFirstLine.includes('ações') || normalizedFirstLine.includes('recomendações')) {
    if (looksLikeExplicitStepBlock(text)) {
      type = 'nextSteps';
      confidence = 'high';
    }
  }

  // Highlights - apenas se for claramente métrica
  else if (looksLikeMetricBlock(text) && text.length <= 200) {
    type = 'highlight';
    confidence = 'high';
  } else if (looksLikeComparisonBlock(text)) {
    type = 'comparison';
    confidence = 'medium';
  } else if (looksLikeKeyValueBlock(text)) {
    type = 'matrix';
    confidence = 'medium';
  } else if (/^(nota\b|observa[çc][aã]o|metodologia|nota metodol[oó]gica|crit[eé]rio de an[aá]lise|ressalva|limita[çc][oõ]es)/i.test(normalizedFirstLine)) {
    type = 'note';
    confidence = 'medium';
  } else if (block.metricsLike && supportsHighlights(context.docType, lowered) && !looksLikeSummarySentence(text) && hasStrongHighlightEvidence(text)) {
    type = 'highlight';
    confidence = 'medium';
  } else if ((block.strongHeading || /^(\d+[\.\)]|[IVXLC]+[\.\)])\s+/i.test(firstLine)) && !isEditorialMetadataBlock(firstLine)) {
    type = 'section';
    confidence = 'high';
    title = firstLine;
  } else if (index <= 1 && text.length >= 220 && text.length <= 1200 && !looksLikeFooterCluster(text) && supportsSummary(context.docType) && !looksLikeHeading(text)) {
    type = 'summary';
    confidence = 'low';
  }

  if (type === 'highlight' && !extractHighlightFromBlock({ text })) {
    type = 'generic';
    confidence = 'low';
  }

  if (type === 'table' && !RE_TABULAR_SEP.test(text) && !RE_SLASH_SEP.test(text) && !RE_DASH_SEP.test(text)) {
    if (!hasStrongTabularEvidence(text)) {
      type = 'generic';
      confidence = 'low';
    }
  }

  return {
    ...block,
    type,
    confidence,
    enabled,
    title: normalizeContentText(title),
    label: blockLabelForType(type),
    format: determineInitialBlockFormat({ ...block, type, title }, context.docType),
    preserve: true,
    keepHeading: true,
    lockAutoTransform: type === 'section' || type === 'table' || type === 'nextSteps' || type === 'matrix' || type === 'comparison' || type === 'note',
    editorialRole: (type === 'table' || type === 'matrix' || type === 'comparison') ? 'table_block' : (title ? 'heading_and_body' : 'body'),
    tableData: block.tableData || null,
  };
}

function supportsSummary(docType) {
  return ['relatorio', 'diagnostico', 'proposta', 'sintese', 'impacto', 'crm', 'parecer'].includes(docType);
}

function supportsHighlights(docType, text) {
  if (docType === 'parecer') return /crit[eé]rio|ponto|indicador|valor|resultado/i.test(text);
  if (docType === 'crm') return /taxa|lead|segment|base|engajamento|convers[aã]o|reten[çc][aã]o|abertura/i.test(text);
  return true;
}

function stripSectionMarkers(text) {
  return String(text || '')
    .replace(/^(\d+[\.\)]|[IVXLC]+[\.\)]|#{1,3})\s*/i, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function blockLabelForType(type) {
  return ({
    summary: 'Resumo',
    section: 'Seção',
    highlight: 'Destaque',
    nextSteps: 'Próximos passos',
    table: 'Tabela/grade',
    matrix: 'Matriz/atributos',
    comparison: 'Quadro comparativo',
    note: 'Nota metodológica',
    generic: 'Texto genérico',
  })[type] || 'Texto genérico';
}

function fallbackSectionTitle(block, docType) {
  const line = stripSectionMarkers(block.firstLine || '');
  if (looksLikeHeading(line)) return line;
  const schema = DOC_TYPE_SCHEMA[docType] || DOC_TYPE_SCHEMA.relatorio;
  const nextIndex = Math.min((state.reviewBlocks || []).filter(item => item.type === 'section').length, schema.defaultSections.length - 1);
  return schema.defaultSections[nextIndex] || 'Seção';
}

function extractHighlightFromBlock(block) {
  const text = String(block.text || '').trim();
  if (isLikelyNoiseBlock(text) || /https?:\/\//i.test(text) || /\bv\d+(?:\.\d+)?\b/i.test(text)) return null;
  const match = text.match(/([A-Za-zÀ-ÿ][A-Za-zÀ-ÿ0-9 %/\-]{2,40})[:\-]\s*([+\-]?\d[\d.,/%x ]{0,18})(?:\s*[—–-]\s*(.+))?/);
  if (match) {
    if (/\bhttps?:/i.test(match[1])) return null;
    return {
      label: match[1].trim(),
      value: match[2].trim(),
      desc: (match[3] || '').trim(),
    };
  }
  const metric = text.match(/([+\-]?\d[\d.,/%x ]{0,18})/);
  if (!metric) return null;
  const parts = text.split(metric[0]);
  const label = parts[0]?.trim().replace(/[:\-–—]+$/, '');
  if (!label || label.length > 48 || /^https?:/i.test(label) || isDateLine(label)) return null;
  return {
    label,
    value: metric[0].trim(),
    desc: parts[1]?.trim() || '',
  };
}

function hasStrongHighlightEvidence(text) {
  const value = String(text || '').trim();
  if (!value || value.length > 220) return false;
  const numericTokens = value.match(/([+\-]?\d[\d.,]*\s?(?:%|x|mil|mi|k|m|R\$)?)/gi) || [];
  if (numericTokens.length !== 1) return false;
  const separators = (value.match(/[:\-–—]/g) || []).length;
  return separators >= 1;
}

function hasStrongTabularEvidence(text) {
  const value = String(text || '').trim();
  const rows = value.split('\n').filter(Boolean);
  const separatedRows = rows.filter(row => RE_TABULAR_SEP.test(row) || RE_SLASH_SEP.test(row) || RE_DASH_SEP.test(row));
  const headerLikeRows = rows.filter(row => isLikelyTableHeaderLine(row));
  const kvRows = rows.filter(row => /^[^:]{3,40}\s*[:\u2013\u2014]\s+.{3,}$/.test(row.trim()));
  return separatedRows.length >= 2 || (rows.length >= 3 && headerLikeRows.length >= 1) || kvRows.length >= 3;
}

function looksLikeComparisonBlock(text) {
  const value = String(text || '').trim();
  const lines = value.split('\n').filter(Boolean);
  if (lines.length < 3 || lines.length > 30) return false;
  const slashLines = lines.filter(line => RE_SLASH_SEP.test(line) && !isSentenceLike(line));
  const dashPairLines = lines.filter(line => RE_DASH_SEP.test(line) && !isSentenceLike(line) && line.trim().length <= 160);
  if (slashLines.length >= 2) return true;
  if (dashPairLines.length >= 2) return true;
  if (slashLines.length >= 1 && (dashPairLines.length >= 1 || lines.length <= 8)) return true;
  return false;
}

function looksLikeKeyValueBlock(text) {
  const value = String(text || '').trim();
  const lines = value.split('\n').filter(Boolean);
  if (lines.length < 2 || lines.length > 30) return false;
  const kvLines = lines.filter(line => {
    const trimmed = line.trim();
    return /^[^:]{3,50}\s*:\s+.{2,}$/.test(trimmed) && trimmed.length <= 180 && !isSentenceLike(trimmed);
  });
  return kvLines.length >= Math.max(2, Math.ceil(lines.length * 0.5));
}

function determineInitialBlockFormat(block, docType) {
  if (block.type === 'nextSteps') return 'list';
  if (block.type === 'table') return 'table';
  if (block.type === 'matrix') return 'matrix';
  if (block.type === 'comparison') return 'comparison';
  if (block.type === 'note') return 'note';
  if (block.type === 'summary') return 'preserve';
  if (block.type === 'section') return 'section';
  if (block.type === 'highlight') return shouldAllowVisualFormatForDocType(docType) ? 'auto' : 'preserve';
  if (String(block.text || '').split('\n').filter(line => looksLikeBullet(line)).length >= 2) return 'list';
  return 'preserve';
}

function shouldAllowVisualFormatForDocType(docType) {
  return docType === 'impacto' || docType === 'crm';
}

function extractStepLines(text) {
  const lines = String(text || '')
    .split('\n')
    .map(line => line.replace(/^[\-\•\*\d\.\)\s]+/, '').trim())
    .filter(line => line.length > 4);
  return lines.length ? lines : [String(text || '').trim()].filter(Boolean);
}

function deriveSummaryFromBlocks(blocks) {
  const summaryIndex = (blocks || []).findIndex(block => block.type === 'summary' && block.enabled);
  if (summaryIndex === -1) return '';
  const summaryBlock = blocks[summaryIndex];
  if (!looksLikeHeading(summaryBlock.text) && summaryBlock.text.length > 40) return summaryBlock.text;

  const collected = [];
  for (let i = summaryIndex + 1; i < blocks.length; i++) {
    const block = blocks[i];
    if (!block.enabled) continue;
    if (block.type === 'section' || block.type === 'summary' || block.type === 'nextSteps') break;
    if (isLikelyNoiseBlock(block.text)) continue;
    if (block.text.length < 30) continue;
    collected.push(block.text);
    if (collected.join(' ').length > 700) break;
  }
  return collected.join(' ').trim();
}

function deriveNextStepsFromBlocks(blocks) {
  const explicit = (blocks || [])
    .filter(block => block.type === 'nextSteps' && block.enabled)
    .flatMap(block => extractStepLines(block.text));

  const filteredExplicit = explicit.filter(line => line.length > 8 && !looksLikeSummarySentence(line));
  if (filteredExplicit.length >= 2) return filteredExplicit.slice(0, 8).join('\n');

  const nextHeadingIndex = (blocks || []).findIndex(block => {
    const candidate = normalizeHeadingText(block.title || block.text).toLowerCase();
    return isNextStepsHeading(candidate);
  });
  if (nextHeadingIndex === -1) return '';

  const steps = [];
  for (let i = nextHeadingIndex + 1; i < blocks.length; i++) {
    const block = blocks[i];
    if (!block.enabled) continue;
    if (block.type === 'section' && i !== nextHeadingIndex + 1 && !looksLikeStepLine(block.text)) break;
    if (isLikelyNoiseBlock(block.text)) continue;
    const lines = extractStepLines(block.text);
    if (!lines.length) continue;
    steps.push(...lines);
    if (steps.length >= 8) break;
  }
  return steps.slice(0, 8).join('\n');
}

function assessExtractionQuality(text, blocks, pageTexts) {
  const notes = [];
  const plainText = String(text || '').trim();
  let score = 0;
  if (plainText.length > 400) score += 2;
  else if (plainText.length > 120) score += 1;
  else notes.push('Pouco texto útil foi extraído do PDF.');

  if ((blocks || []).length >= 4) score += 2;
  else if ((blocks || []).length >= 2) score += 1;
  else notes.push('A separação em blocos ficou fraca, sugerindo ordem de leitura ruim.');

  if ((pageTexts || []).some(page => page.length > 80)) score += 1;
  if ((plainText.match(/[�]/g) || []).length) notes.push('Há caracteres corrompidos na extração.');
  if (looksLikeMostlyHeaderFooter(blocks)) {
    score = Math.min(score, 1);
    notes.push('A extração parece ter capturado majoritariamente cabeçalhos, rodapés ou URLs.');
  }

  const confidence = score >= 4 ? 'high' : score >= 2 ? 'medium' : 'low';
  if (confidence === 'low') notes.push('Importação com baixa confiabilidade: prefira revisar manualmente antes de aplicar.');
  return { confidence, notes };
}

function normalizeHeadingText(text) {
  const value = String(text || '').trim();
  if (/^(?:[A-ZÀ-Ý]\s){3,}[A-ZÀ-Ý]$/i.test(value)) {
    return value.replace(/\s+/g, '');
  }
  return value.replace(/\s{2,}/g, ' ').trim();
}

function looksLikeHeading(text, nextText = '') {
  const value = normalizeHeadingText(text);
  const next = normalizeHeadingText(nextText);
  if (!value || value.length > 140) return false;
  if (isLikelyNoiseLine(value) || isMetadataLine(value) || isTitleNoise(value)) return false;

  const uppercaseRatio = (value.match(/[A-ZÀ-Ý]/g) || []).length / Math.max((value.match(/[A-Za-zÀ-ÿ]/g) || []).length, 1);
  const numbered = /^\d+(?:\.\d+)?(?:\s*[-–—.:)]\s*|\s+)/.test(value);
  const keywordish = /\b(resumo|sum[aá]rio|introdu[cç][aã]o|contexto|objetivos?|an[aá]lise|diagn[oó]stico|metodologia|resultados?|conclus[aã]o|recomenda[cç][oõ]es|pr[oó]ximos?\s+passos|impacto|crm|segmenta[cç][aã]o|parcerias?)\b/i.test(value);
  const shortHeading = value.length <= 72 && !/[.!?;]$/.test(value);
  const nextLooksBody = !!next && next.length > value.length && /[a-zà-ÿ]{4,}\s+[a-zà-ÿ]{4,}/i.test(next);

  return numbered || keywordish || (shortHeading && uppercaseRatio >= 0.55) || (shortHeading && nextLooksBody);
}

function looksLikeNumberedHeading(text) {
  const value = normalizeHeadingText(text);
  return /^\d+(?:\.\d+)?\s*[-–—.:)]\s*[A-Za-zÀ-ÿ]/.test(value)
    || /^[IVXLC]+\s*[-–—.:)]\s*[A-Za-zÀ-ÿ]/i.test(value);
}

function computeUppercaseRatio(text) {
  const letters = String(text || '').match(/[A-Za-zÀ-ÿ]/g) || [];
  if (!letters.length) return 0;
  const uppers = String(text || '').match(/[A-ZÀ-Ý]/g) || [];
  return uppers.length / letters.length;
}

function estimateLexicalDensity(text) {
  const tokens = String(text || '').trim().match(/[A-Za-zÀ-ÿ0-9]+/g) || [];
  if (!tokens.length) return 0;
  const contentish = tokens.filter(token => token.length >= 4 || /^\d/.test(token));
  return contentish.length / tokens.length;
}

function isSentenceLike(text) {
  const value = String(text || '').trim();
  if (!value) return false;
  if (/[.!?;:]$/.test(value)) return true;
  if (value.length > 90 && /[a-zà-ÿ]{3,}\s+[a-zà-ÿ]{3,}\s+[a-zà-ÿ]{3,}/i.test(value)) return true;
  if (/\b(que|para|com|como|onde|porque|foram|será|estão|estava|identificados?|considerando|houve|apresenta|indica)\b/i.test(value) && value.length > 48) return true;
  return false;
}

function isTechnicalLeakLine(text) {
  const value = String(text || '').trim();
  if (!value) return true;
  return /^(?:id|uuid|guid|field|label|name|value|type)$/i.test(value)
    || /^(?:id|uuid|guid|field|label|name|value|type)\s*[:=]/i.test(value)
    || (/\b(?:json|payload|serialized|schema|undefined|null)\b/i.test(value) && value.length < 60);
}

function normalizeHeadingLabel(text) {
  let value = normalizeHeadingText(text);
  if (!value) return '';
  value = value
    .replace(/\b(id|uuid|guid|field|label|name|value|type)\b/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  const parts = value.split(/\s+/);
  if (parts.length >= 2) {
    const deduped = [];
    for (const part of parts) {
      if (foldForMatch(deduped[deduped.length - 1] || '') === foldForMatch(part)) continue;
      deduped.push(part);
    }
    value = deduped.join(' ');
  }
  const doubled = value.match(/^(.{4,80}?)\s+\1$/i);
  if (doubled) value = doubled[1].trim();
  return value.trim();
}

function lineHasHeadingTraits(line, bodyFont = 10, nextLine = null) {
  if (!line?.text) return false;
  const text = sanitizeEditorialLine(line.text);
  if (!text || isLikelyNoiseLine(text) || isTechnicalLeakLine(text) || isMetadataLine(text) || isTitleNoise(text)) return false;
  const lineFont = Number(line.fontSize) || bodyFont || 10;
  const nextFont = Number(nextLine?.fontSize) || bodyFont || 10;
  const uppercaseRatio = Number(line.uppercaseRatio ?? computeUppercaseRatio(text));
  const lexicalDensity = Number(line.lexicalDensity ?? estimateLexicalDensity(text));
  const sentenceLike = typeof line.sentenceLike === 'boolean' ? line.sentenceLike : isSentenceLike(text);
  const visuallyProminent = lineFont >= bodyFont * MIN_FONT_SIZE_RATIO_FOR_HEADING || lineFont > nextFont * 1.03;
  const continuation = /^[A-ZÀ-Ý0-9][A-ZÀ-Ýa-zà-ÿ0-9'’\-—–,:()]{1,36}$/.test(text) && !/[.!?;:]$/.test(text);
  return (
    looksLikeNumberedHeading(text)
    || looksLikeHeading(text, nextLine?.text || '')
    || (uppercaseRatio >= 0.72 && !sentenceLike)
    || (continuation && lexicalDensity >= 0.45 && visuallyProminent)
    || (continuation && uppercaseRatio >= 0.58 && !sentenceLike)
  );
}

  function shouldMergeHeadingLines(prevLine, line, bodyFont = 10, skipHeadingCheck = false) {
  if (!prevLine || !line) return false;
  const prevText = sanitizeEditorialLine(prevLine.text || '');
  const text = sanitizeEditorialLine(line.text || '');
  if (!prevText || !text) return false;
  if (looksLikeBullet(text) || looksLikeStepLine(text)) return false;
  if (isSentenceLike(text) && !looksLikeNumberedHeading(text)) return false;
  const prevFont = Number(prevLine.fontSize || bodyFont || 10);
  const lineFont = Number(line.fontSize || bodyFont || 10);
  const similarFont = Math.abs(prevFont - lineFont) <= Math.max(1.5, bodyFont * 0.18);
  const closeGap = Number(prevLine.gapAfter || 0) <= Math.max(18, prevFont * 1.45);
  const similarX = Math.abs(Number(prevLine.x || 0) - Number(line.x || 0)) <= Math.max(18, prevFont * 1.6);
  const prevHeadingish = skipHeadingCheck
    ? (looksLikeNumberedHeading(prevText) || computeUppercaseRatio(prevText) >= 0.6 || lineHasHeadingTraits(prevLine, bodyFont, line))
    : lineHasHeadingTraits(prevLine, bodyFont, line);
  const dashContinuation = /^[—–\-]\s*/.test(text);
  const lowercaseContinuation = /^[a-zà-ÿ]/.test(text) && text.length <= 42 && !isSentenceLike(text);
  const lineHeadingish = skipHeadingCheck
    ? (/^[A-ZÀ-Ý0-9]/.test(text) || lineHasHeadingTraits(line, bodyFont, null) || dashContinuation || lowercaseContinuation)
    : (lineHasHeadingTraits(line, bodyFont, null) || dashContinuation || lowercaseContinuation);
  const prevOpen = !/[.!?;:]$/.test(prevText);
  const shortContinuation = text.length <= 48 || /^[A-ZÀ-Ý0-9a-zà-ÿ—–\-][A-Za-zÀ-ÿ0-9''\-—–,:() ]{1,54}$/.test(text);
  return similarFont && closeGap && similarX && prevHeadingish && lineHeadingish && prevOpen && shortContinuation;
}

function isHeadingOrphanFragment(text) {
  const value = sanitizeEditorialLine(text || '');
  if (!value) return true;
  if (isTechnicalLeakLine(value)) return true;
  if (looksLikeNumberedHeading(value)) return false;
  const words = value.split(/\s+/).filter(Boolean);
  return words.length <= 3
    && value.length <= 24
    && !/[.!?;:]/.test(value)
    && computeUppercaseRatio(value) >= 0.65;
}

function shouldAttachBlockToPreviousHeading(prev, block) {
  if (!prev?.heading || !block?.text || block.heading) return false;
  const text = sanitizeEditorialLine(block.text);
  if (!text || isSentenceLike(text)) return false;
  if (text.length > 42) return false;
  return isHeadingOrphanFragment(text)
    || (computeUppercaseRatio(text) >= 0.7 && /^[A-ZÀ-Ý0-9]/.test(text));
}

function shouldPullLeadingFragmentIntoHeading(prev, block) {
  if (!prev?.heading || !block?.text || block.tableLike) return false;
  const paragraphs = splitTextIntoParagraphs(block.text);
  if (!paragraphs.length) return false;
  const first = sanitizeEditorialLine(paragraphs[0]);
  if (!first || !isHeadingOrphanFragment(first)) return false;
  const rest = paragraphs.slice(1).join(' ');
  return !rest || isSentenceLike(rest) || rest.length > first.length;
}

function reattachLeadingHeadingFragment(heading, text) {
  const paragraphs = splitTextIntoParagraphs(text);
  if (!paragraphs.length) return { heading: normalizeHeadingLabel(heading), text: normalizeContentText(text) };
  const [first, ...rest] = paragraphs;
  return {
    heading: normalizeHeadingLabel(`${heading} ${first}`),
    text: rest.join('\n\n').trim(),
  };
}

function foldForMatch(text) {
  return normalizeHeadingText(text)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9]/g, '')
    .toLowerCase();
}

function isLikelyNoiseLine(text) {
  const value = String(text || '').trim();
  return !value
    || /^https?:\/\//i.test(value)
    || /\b\d+\/\d+\b/.test(value)
    || /^[✦◆◈◉●•\-\s]+$/.test(value)
    || /^official portfolio document$/i.test(value)
    || /^g u d i$/i.test(value)
    || isBrandingLine(value);
}

function isLikelyNoiseBlock(text) {
  const value = String(text || '').trim();
  return isLikelyNoiseLine(value)
    || looksLikeFooterCluster(value)
    || /^autoria\s+[aá]rea\s+data\s+vers[aã]o$/i.test(normalizeHeadingText(value))
    || /^documentos oficiais$/i.test(value)
    || isDocumentTypeBadge(value);
}

function looksLikeFooterCluster(text) {
  const value = String(text || '').trim();
  return /^https?:\/\//i.test(value)
    || /(example\.com|127\.0\.0\.1|autotest-raster-source\.html)/i.test(value)
    || (/\b\d{1,2}\/\d{1,2}\/\d{2,4}\b/i.test(value) && value.length < 90);
}

function looksLikeMostlyHeaderFooter(blocks) {
  const items = (blocks || []).filter(block => block.text);
  if (!items.length) return true;
  const noisy = items.filter(block => isLikelyNoiseBlock(block.text) || looksLikeFooterCluster(block.text));
  return noisy.length / items.length >= 0.5;
}

function looksLikeExplicitStepBlock(text) {
  const lines = String(text || '').split('\n').map(line => line.trim()).filter(Boolean);
  const bulletLines = lines.filter(line => looksLikeStepLine(line));
  return bulletLines.length >= 2;
}

function looksLikeMetricBlock(text) {
  const value = String(text || '').trim();
  if (!value || isLikelyNoiseBlock(value)) return false;
  const hasNumber = /(?:^|[\s:])(?:[+\-]?\d[\d.,]*\s?(?:%|x|mil|mi|k|m|R\$)?)(?:$|[\s,.;:])/i.test(value);
  const metricKeywordCount = (value.match(/\b(kpi|indicador|m[eé]trica|taxa|nps|receita|abertura|clique|reten[cç][aã]o|convers[aã]o|ticket|volume|usuarios?|leads?|alcance|impacto)\b/gi) || []).length;
  const compactPattern = value.length <= 160 && /[:\-]/.test(value);
  return hasNumber && (metricKeywordCount >= 1 || compactPattern);
}

function looksLikeTableBlock(lines) {
  if (!Array.isArray(lines) || lines.length < 2) return false;
  const withSeparators = lines.filter(line => RE_TABULAR_SEP.test(String(line?.text || '')));
  const headerLike = lines.filter(line => isLikelyTableHeaderLine(String(line?.text || '')));
  const shortRows = lines.filter(line => isLikelyTabularLine(String(line?.text || '')));
  return withSeparators.length >= 2 || (headerLike.length >= 1 && shortRows.length >= 3);
}

function isLikelyTableHeaderLine(text) {
  const value = sanitizeEditorialLine(text || '');
  if (!value) return false;
  const tokens = value.split(/\s{2,}|\t|\s+\|\s+|;/).map(part => part.trim()).filter(Boolean);
  const slashTokens = value.split(/\s+\/\s+/).map(part => part.trim()).filter(Boolean);
  const dashTokens = value.split(/\s+[—–]{1,2}\s+/).map(part => part.trim()).filter(Boolean);
  const compactTokens = value.split(/\s+/).filter(Boolean);
  const knownHeaders = /(?:atributo|detalhes|grupo|segmento|crit[eé]rio|estrat[eé]gia|inscri[cç][oõ]es|inscritos|presentes|ausentes|taxa|absten[cç][aã]o|temperatura|perfil|volume|prioridade|tipo de patrocinador|argumento|garantido|infer[eê]ncia|raz[aã]o|categoria|interpreta[çc][aã]o|indicador|resultado|meta|status|descri[çc][aã]o|respons[aá]vel|prazo|a[çc][aã]o|objetivo|entrega|fase|etapa)/i.test(value);
  const slashSeparated = slashTokens.length >= 2 && slashTokens.length <= 8 && !isSentenceLike(value);
  const dashSeparated = dashTokens.length >= 2 && dashTokens.length <= 6 && computeUppercaseRatio(value) >= 0.40 && !isSentenceLike(value);
  const upper = computeUppercaseRatio(value) >= 0.62;
  return knownHeaders || tokens.length >= 2 || slashSeparated || dashSeparated || (compactTokens.length >= 2 && compactTokens.length <= 8 && upper && !isSentenceLike(value));
}

function isLikelyTabularLine(text) {
  const value = sanitizeEditorialLine(text || '');
  if (!value || value.length > 220) return false;
  if (looksLikeHeading(value) || looksLikeBullet(value) || isSentenceLike(value)) return false;
  if (RE_TABULAR_SEP.test(value)) return true;
  if (RE_SLASH_SEP.test(value) && value.length <= 160) return true;
  if (RE_DASH_SEP.test(value) && value.length <= 160) return true;
  if (/^[^:]{3,40}\s*[:–—]\s+.{3,}$/.test(value) && value.length <= 160) return true;
  const numericDensity = (value.match(/\d+/g) || []).length;
  const chunks = value.split(RE_TABULAR_SPLIT).filter(Boolean);
  return chunks.length >= 2 || (numericDensity >= 1 && value.split(/\s+/).length <= 14);
}

function splitTableCells(text) {
  const value = sanitizeEditorialLine(text || '');
  if (!value) return [];
  const explicit = value.split(RE_TABULAR_SPLIT).map(part => part.trim()).filter(Boolean);
  if (explicit.length >= 2) return explicit;
  if (RE_SLASH_SEP.test(value)) {
    const slashParts = value.split(RE_SLASH_SEP).map(part => part.trim()).filter(Boolean);
    if (slashParts.length >= 2) return slashParts;
  }
  if (RE_DASH_SEP.test(value)) {
    const dashParts = value.split(RE_DASH_SEP).map(part => part.trim()).filter(Boolean);
    if (dashParts.length >= 2 && dashParts.length <= 6) return dashParts;
  }
  if (/^([^:]{3,40})\s*:\s+(.{3,})$/.test(value)) {
    const match = value.match(/^([^:]{3,40})\s*:\s+(.{3,})$/);
    if (match) return [match[1].trim(), match[2].trim()];
  }
  const knownPairs = value.match(/(?:TOTAL|INSCRIÇÕES|PRESENTES|AUSENTES|TAXA|ABSTENÇÃO|TEMPERATURA|PERFIL|VOLUME|PRIORIDADE|GRUPO|INSCRITOS|SEGMENTO|CRITÉRIO|ESTRATÉGIA|ATRIBUTO|DETALHES|RAZÃO|ARGUMENTO|GARANTIDO|INFERÊNCIA|TIPO DE PATROCINADOR|RELEVANTE|INDICADOR|RESULTADO|META|STATUS|OBJETIVO|ENTREGA|FASE|ETAPA)/gi) || [];
  if (knownPairs.length >= 2 && computeUppercaseRatio(value) >= 0.50) return knownPairs.map(item => item.trim());
  return [value];
}

function extractTableStructureFromLines(lines) {
  if (!Array.isArray(lines) || lines.length < 2) return null;
  const normalized = lines.map(line => ({ ...line, text: sanitizeEditorialLine(line.text || '') })).filter(line => line.text);
  if (normalized.length < 2) return null;
  let captionLines = [];
  let startIndex = 0;
  const medianFs = median(normalized.map(line => line.fontSize).filter(Boolean)) || 10;
  if (normalized.length >= 3 && isEditorialHeadingLine(normalized[0], normalized[1], null, medianFs) && isLikelyTableHeaderLine(normalized[1].text)) {
    captionLines = [normalized[0]];
    startIndex = 1;
  }
  if (startIndex === 0 && normalized.length >= 3 && isLikelyTableHeaderLine(normalized[0].text)) {
    const firstCells = splitTableCells(normalized[0].text);
    if (firstCells.length >= 2) startIndex = 0;
  }
  const candidateLines = normalized.slice(startIndex);
  const tableishCount = candidateLines.filter(line => isLikelyTabularLine(line.text) || isLikelyTableHeaderLine(line.text)).length;
  const threshold = Math.max(2, Math.ceil(candidateLines.length * 0.55));
  if (tableishCount < threshold) return null;
  const header = splitTableCells(candidateLines[0].text);
  if (header.length < 2 && !isLikelyTableHeaderLine(candidateLines[0].text)) return null;
  const rows = candidateLines.slice(1)
    .map(line => splitTableCells(line.text))
    .filter(cells => cells.length >= 1);
  if (!rows.length) return null;
  return {
    caption: normalizeHeadingLabel(captionLines.map(line => line.text).join(' ')),
    captionLines,
    header,
    rows,
  };
}

function serializeTableStructure(tableData) {
  if (!tableData) return '';
  return [tableData.header.join(' | '), ...tableData.rows.map(row => row.join(' | '))].join('\n');
}

function looksLikeSummarySentence(text) {
  const value = String(text || '').trim();
  return value.length > 80 && /(?:\.\s|,\s)/.test(value);
}

function normalizeParagraphText(text) {
  return String(text || '')
    .replace(/\s+/g, ' ')
    .replace(/\s+([,.;:!?])/g, '$1')
    .trim();
}

function shouldKeepParagraphFlow(paragraph, lineText) {
  if (/^[a-zà-ÿ0-9]/i.test(lineText) && /[a-zà-ÿ0-9,\)]$/i.test(paragraph)) return true;
  if (/^[A-ZÀ-Ý][a-zà-ÿ]/.test(lineText) && /[,;]$/.test(paragraph)) return true;
  if (/^\(?[0-9a-zà-ÿ]/i.test(lineText) && !looksLikeHeading(lineText)) return true;
  return false;
}

function shouldMergeIndentedLine(prevText, lineText) {
  return !!prevText && !!lineText && !looksLikeHeading(lineText) && !looksLikeBullet(lineText) && !/[.:;]$/.test(prevText);
}

function looksLikeBullet(text) {
  const value = String(text || '').trim();
  return /^([\-\*\u2022\u25E6\u25A0\u25AA\u25B8\u25B9]|[a-z]\)|\d+[\.\)])\s+/i.test(value);
}

function isBrandingLine(text) {
  const value = foldForMatch(text);
  return ['docs', 'documentgeneratorworkbench', 'officialdocuments', 'examplecom'].includes(value);
}

function isDocumentTypeBadge(text) {
  const value = foldForMatch(text);
  return [
    'relatorioexecutivo',
    'parecerinstitucional',
    'relatoriodecrm',
    'diagnostico',
    'proposta',
    'sinteseestrategica',
    'relatoriodeimpacto',
  ].includes(value);
}

function isTitleNoise(text) {
  const raw = normalizeHeadingText(text);
  const folded = foldForMatch(text);
  return isBrandingLine(raw)
    || isDocumentTypeBadge(raw)
    || isEditorialMetadataBlock(raw)
    || folded.includes('vivaobomdavida')
    || folded.includes('documentosoficiais')
    || ['officialportfoliodocument', 'executivereport', 'institutionalreview', 'crmanalysisreport'].includes(folded);
}

function looksLikeRealDocumentTitle(text) {
  const value = String(text || '').trim();
  return value.length >= 16
    && value.length <= 120
    && !isTitleNoise(value)
    && !isMetadataLine(value)
    && /[A-Za-zÀ-ÿ]/.test(value)
    && (/[·:\-]/.test(value) || /\b(relat[oó]rio|an[aá]lise|plano|estrat[eé]gia|crescimento|engajamento|contexto|mercado|crm)\b/i.test(value));
}

function isEditorialMetadataBlock(text) {
  const value = foldForMatch(text);
  return value.startsWith('autoriaareadataversao')
    || value.startsWith('autoriadataversao')
    || value.startsWith('areadataversao')
    || value.startsWith('autoriaarea')
    || value === 'autoria'
    || value === 'versao';
}

function isNextStepsHeading(text) {
  const value = foldForMatch(text);
  return [
    'proximospassos',
    'proximopasso',
    'proximasacoes',
    'recomendacoes',
    'acoesrecomendadas',
    'encaminhamentos',
    'planodeacao',
  ].includes(value);
}

function looksLikeStepLine(text) {
  return looksLikeBullet(text) || /^\d+\s+/.test(String(text || '').trim()) || /^(priorizar|lançar|expandir|iniciar|implementar|consolidar|revisar|ajustar|escalar)\b/i.test(String(text || '').trim());
}

function isMetadataLine(line) {
  return isDateLine(line) || isAuthorLine(line) || /vers[aã]o|área|departamento|equipe|setor/i.test(String(line || ''));
}

function isDateLine(line) { return /\b\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}\b|\b\d{4}\b|janeiro|fevereiro|março|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro/i.test(line) && line.length < 42; }
function isAuthorLine(line) { return /autoria|autor|elaborado|responsável|assinado/i.test(line); }
function capitalizeWords(str) { return str.toLowerCase().replace(/\b\w/g, c => c.toUpperCase()); }
function median(values) {
  if (!values?.length) return 0;
  const arr = [...values].sort((a, b) => a - b);
  const mid = Math.floor(arr.length / 2);
  return arr.length % 2 ? arr[mid] : (arr[mid - 1] + arr[mid]) / 2;
}

/* ══════════════════════════════════════════════════════════════
   MÓDULO: REVIEW MODAL
   ══════════════════════════════════════════════════════════════ */
function openReview(mapped, rawText, warning, pages) {
  els.rv_title.value    = mapped.title;
  els.rv_subtitle.value = mapped.subtitle;
  els.rv_date.value     = mapped.date;
  els.rv_version.value  = mapped.version;
  els.rv_author.value   = mapped.author;
  els.rv_dept.value     = mapped.dept;
  els.rv_summary.value  = mapped.summary;
  els.rv_nextSteps.value= mapped.nextSteps;
  els.rv_raw.value      = rawText.slice(0, 5000);
  els.rv_docType.value  = mapped.docType;
  els.rv_structureMode.value = inferStructureMode(mapped.docType, mapped.blocks || []);

  els.reviewSectionsList.innerHTML = '';
  state.reviewBlocks = [];
  state.reviewBlockCounter = 0;
  (mapped.blocks || []).forEach(block => addReviewBlock(block));

  const noteText = (mapped.extractionNotes || []).join(' ');
  if (pages) {
    const confidenceLabel = mapped.confidence === 'high' ? 'alta' : mapped.confidence === 'medium' ? 'média' : 'baixa';
    els.reviewPageInfo.textContent = `${pages} página(s) extraída(s) · confiança ${confidenceLabel}`;
  }

  if (warning || noteText) {
    els.reviewWarning.style.display = 'flex';
    els.reviewWarningText.textContent = [warning, noteText].filter(Boolean).join(' ');
  } else {
    els.reviewWarning.style.display = 'none';
  }

  els.reviewOverlay.style.display = 'flex';
  document.body.style.overflow = 'hidden';
}

function openReviewWithError(errorMsg, filename) {
  openReview(
    { title: filenameToTitle(filename), subtitle: '', date: '', version: '', author: '', dept: '', summary: '', nextSteps: '', docType: state.docType || 'relatorio', sections: [], blocks: [], extractionNotes: [], confidence: 'low' },
    '',
    `Não foi possível extrair texto automaticamente: ${errorMsg} Preencha os campos manualmente.`,
    null
  );
}

function closeReview() {
  els.reviewOverlay.style.display = 'none';
  document.body.style.overflow = '';
}

function addReviewBlock(block = {}) {
  const id = ++state.reviewBlockCounter;
  const item = {
    id,
    type: block.type || 'generic',
    confidence: block.confidence || 'low',
    title: block.title || '',
    text: block.text || '',
    enabled: block.enabled !== false,
    format: block.format || defaultBlockFormat(block.type),
    preserve: block.preserve !== false,
    lockAutoTransform: !!block.lockAutoTransform,
    keepHeading: block.keepHeading !== false,
    source: block.source || 'manual',
    page: block.page || null,
    editorialRole: block.editorialRole || (block.tableLike ? 'table_block' : (block.title ? 'heading_and_body' : 'body')),
    tableData: block.tableData || null,
  };
  state.reviewBlocks.push(item);

  const el = document.createElement('div');
  el.className = 'review-sec-item';
  el.dataset.id = id;
  el.innerHTML = `
    <div class="review-sec-item-hd">
      <input type="text" placeholder="Título do bloco" value="${escHtml(item.title)}" data-field="title"/>
      <select class="review-block-type" data-field="type">
        <option value="generic"${item.type === 'generic' ? ' selected' : ''}>Texto genérico</option>
        <option value="section"${item.type === 'section' ? ' selected' : ''}>Seção</option>
        <option value="summary"${item.type === 'summary' ? ' selected' : ''}>Resumo</option>
        <option value="highlight"${item.type === 'highlight' ? ' selected' : ''}>Destaque</option>
        <option value="nextSteps"${item.type === 'nextSteps' ? ' selected' : ''}>Próximos passos</option>
        <option value="table"${item.type === 'table' ? ' selected' : ''}>Tabela</option>
        <option value="matrix"${item.type === 'matrix' ? ' selected' : ''}>Matriz/atributos</option>
        <option value="comparison"${item.type === 'comparison' ? ' selected' : ''}>Quadro comparativo</option>
        <option value="note"${item.type === 'note' ? ' selected' : ''}>Nota metodológica</option>
      </select>
      <select class="review-block-confidence" data-field="confidence">
        <option value="high"${item.confidence === 'high' ? ' selected' : ''}>Alta</option>
        <option value="medium"${item.confidence === 'medium' ? ' selected' : ''}>Média</option>
        <option value="low"${item.confidence === 'low' ? ' selected' : ''}>Baixa</option>
      </select>
      <label class="review-block-toggle">
        <input type="checkbox" data-field="enabled"${item.enabled ? ' checked' : ''}/>
        <span>Usar</span>
      </label>
      <button class="review-sec-del" title="Remover">✕</button>
    </div>
    <div class="review-sec-item-meta">
      <select class="review-block-format" data-field="format">
        <option value="auto"${item.format === 'auto' ? ' selected' : ''}>Auto prudente</option>
        <option value="section"${item.format === 'section' ? ' selected' : ''}>Seção textual</option>
        <option value="preserve"${item.format === 'preserve' ? ' selected' : ''}>Bloco corrido preservado</option>
        <option value="visual"${item.format === 'visual' ? ' selected' : ''}>Componente visual</option>
        <option value="list"${item.format === 'list' ? ' selected' : ''}>Lista</option>
        <option value="table"${item.format === 'table' ? ' selected' : ''}>Tabela</option>
        <option value="matrix"${item.format === 'matrix' ? ' selected' : ''}>Matriz/atributos</option>
        <option value="comparison"${item.format === 'comparison' ? ' selected' : ''}>Quadro comparativo</option>
        <option value="note"${item.format === 'note' ? ' selected' : ''}>Nota metodológica</option>
      </select>
      <label class="review-block-toggle">
        <input type="checkbox" data-field="preserve"${item.preserve ? ' checked' : ''}/>
        <span>Preservar bloco</span>
      </label>
      <label class="review-block-toggle">
        <input type="checkbox" data-field="keepHeading"${item.keepHeading ? ' checked' : ''}/>
        <span>Preservar heading</span>
      </label>
      <select class="review-block-format" data-field="editorialRole">
        <option value="heading_and_body"${item.editorialRole === 'heading_and_body' ? ' selected' : ''}>Heading + corpo</option>
        <option value="heading_only"${item.editorialRole === 'heading_only' ? ' selected' : ''}>Heading único</option>
        <option value="table_block"${item.editorialRole === 'table_block' ? ' selected' : ''}>Tabela coesa</option>
        <option value="table_header"${item.editorialRole === 'table_header' ? ' selected' : ''}>Cabeçalho de tabela</option>
        <option value="table_row"${item.editorialRole === 'table_row' ? ' selected' : ''}>Linha de tabela</option>
        <option value="body"${item.editorialRole === 'body' ? ' selected' : ''}>Somente corpo</option>
        <option value="noise"${item.editorialRole === 'noise' ? ' selected' : ''}>Ruído</option>
      </select>
      <label class="review-block-toggle">
        <input type="checkbox" data-field="lockAutoTransform"${item.lockAutoTransform ? ' checked' : ''}/>
        <span>Travar auto</span>
      </label>
    </div>
    <div class="review-seg-actions">
      <button type="button" class="review-seg-btn" data-action="merge-up">Unir acima</button>
      <button type="button" class="review-seg-btn" data-action="merge-down">Unir abaixo</button>
      <button type="button" class="review-seg-btn" data-action="extract-heading">Heading único</button>
      <button type="button" class="review-seg-btn" data-action="continue-heading">Continua heading anterior</button>
      <button type="button" class="review-seg-btn" data-action="mark-table">Tabela</button>
      <button type="button" class="review-seg-btn" data-action="mark-table-header">Cabeçalho tabela</button>
      <button type="button" class="review-seg-btn" data-action="mark-table-row">Linha tabela</button>
      <button type="button" class="review-seg-btn" data-action="mark-body">Marcar corpo</button>
      <button type="button" class="review-seg-btn review-seg-btn--warn" data-action="mark-noise">Ruído</button>
      <button type="button" class="review-seg-btn" data-action="split-block">Separar bloco</button>
    </div>
    <textarea placeholder="Conteúdo do bloco..." data-field="text">${escHtml(item.text)}</textarea>`;

  el.querySelector('.review-sec-del').addEventListener('click', () => {
    state.reviewBlocks = state.reviewBlocks.filter(s => s.id !== id);
    el.remove();
  });
  el.querySelectorAll('[data-field]').forEach(inp => {
    const eventName = inp.type === 'checkbox' || inp.tagName === 'SELECT' ? 'change' : 'input';
    inp.addEventListener(eventName, () => {
      const s = state.reviewBlocks.find(s => s.id === id);
      if (s) s[inp.dataset.field] = inp.type === 'checkbox' ? inp.checked : inp.value;
    });
  });

  el.querySelectorAll('[data-action]').forEach(btn => {
    btn.addEventListener('click', () => handleReviewSegmentationAction(btn.dataset.action, el));
  });

  els.reviewSectionsList.appendChild(el);
}

function handleReviewSegmentationAction(action, el) {
  if (!el) return;
  if (action === 'merge-up') return mergeReviewBlockWithSibling(el, -1);
  if (action === 'merge-down') return mergeReviewBlockWithSibling(el, 1);
  if (action === 'extract-heading') return extractHeadingForReviewBlock(el);
  if (action === 'continue-heading') return continueHeadingFromPreviousReviewBlock(el);
  if (action === 'mark-table') return markReviewBlockAsTable(el);
  if (action === 'mark-table-header') return markReviewBlockAsTableHeader(el);
  if (action === 'mark-table-row') return markReviewBlockAsTableRow(el);
  if (action === 'mark-body') return markReviewBlockAsBody(el);
  if (action === 'mark-noise') return markReviewBlockAsNoise(el);
  if (action === 'split-block') return splitReviewBlock(el);
}

function getReviewBlockElements() {
  return Array.from(els.reviewSectionsList.querySelectorAll('.review-sec-item'));
}

function getReviewBlockIndex(el) {
  return getReviewBlockElements().findIndex(item => item === el);
}

function mergeReviewBlockWithSibling(el, direction) {
  const blocks = getReviewBlockElements();
  const index = getReviewBlockIndex(el);
  const sibling = blocks[index + direction];
  if (!sibling) return;

  const primary = direction < 0 ? sibling : el;
  const secondary = direction < 0 ? el : sibling;

  const primaryTitle = primary.querySelector('[data-field="title"]');
  const primaryText = primary.querySelector('[data-field="text"]');
  const secondaryTitle = secondary.querySelector('[data-field="title"]');
  const secondaryText = secondary.querySelector('[data-field="text"]');

  const mergedTitle = normalizeHeadingLabel([primaryTitle?.value, secondaryTitle?.value].filter(Boolean).join(' '));
  const mergedText = [primaryText?.value, secondaryText?.value].filter(Boolean).join('\n\n');

  if (primaryTitle) primaryTitle.value = mergedTitle;
  if (primaryText) primaryText.value = normalizeContentText(mergedText);
  primary.querySelector('[data-field="keepHeading"]').checked = !!mergedTitle;
  primary.querySelector('[data-field="editorialRole"]').value = mergedTitle ? 'heading_and_body' : 'body';
  syncReviewBlockState(primary);

  const secondaryId = Number(secondary.dataset.id);
  state.reviewBlocks = state.reviewBlocks.filter(item => item.id !== secondaryId);
  secondary.remove();
}

function extractHeadingForReviewBlock(el) {
  const titleInput = el.querySelector('[data-field="title"]');
  const textArea = el.querySelector('[data-field="text"]');
  const currentTitle = normalizeHeadingLabel(titleInput?.value || '');
  const currentText = normalizeContentText(textArea?.value || '');
  const paragraphs = splitTextIntoParagraphs(currentText);
  const candidate = currentTitle || paragraphs[0] || '';
  const rest = currentTitle ? currentText : paragraphs.slice(1).join('\n\n');
  const extracted = normalizeHeadingLabel(candidate);
  if (titleInput) titleInput.value = extracted;
  if (textArea) textArea.value = normalizeContentText(removeHeadingEcho(rest || currentText, extracted));
  el.querySelector('[data-field="type"]').value = extracted ? 'section' : 'generic';
  el.querySelector('[data-field="keepHeading"]').checked = !!extracted;
  el.querySelector('[data-field="editorialRole"]').value = textArea?.value ? 'heading_and_body' : 'heading_only';
  syncReviewBlockState(el);
}

function continueHeadingFromPreviousReviewBlock(el) {
  const blocks = getReviewBlockElements();
  const index = getReviewBlockIndex(el);
  if (index <= 0) return;
  const prev = blocks[index - 1];
  const prevTitle = prev.querySelector('[data-field="title"]');
  const prevText = prev.querySelector('[data-field="text"]');
  const textArea = el.querySelector('[data-field="text"]');
  const titleInput = el.querySelector('[data-field="title"]');
  const fragment = normalizeHeadingLabel([titleInput?.value, textArea?.value].filter(Boolean).join(' '));
  if (!fragment) return;
  prevTitle.value = normalizeHeadingLabel(`${prevTitle?.value || ''} ${fragment}`);
  prev.querySelector('[data-field="keepHeading"]').checked = true;
  prev.querySelector('[data-field="editorialRole"]').value = normalizeContentText(prevText?.value || '') ? 'heading_and_body' : 'heading_only';
  syncReviewBlockState(prev);
  markReviewBlockAsNoise(el);
}

function markReviewBlockAsBody(el) {
  el.querySelector('[data-field="type"]').value = 'generic';
  el.querySelector('[data-field="title"]').value = '';
  el.querySelector('[data-field="keepHeading"]').checked = false;
  el.querySelector('[data-field="editorialRole"]').value = 'body';
  el.querySelector('[data-field="format"]').value = 'preserve';
  syncReviewBlockState(el);
}

function markReviewBlockAsTable(el) {
  el.querySelector('[data-field="type"]').value = 'table';
  el.querySelector('[data-field="editorialRole"]').value = 'table_block';
  el.querySelector('[data-field="format"]').value = 'table';
  el.querySelector('[data-field="lockAutoTransform"]').checked = true;
  syncReviewBlockState(el);
}

function markReviewBlockAsTableHeader(el) {
  el.querySelector('[data-field="type"]').value = 'table';
  el.querySelector('[data-field="editorialRole"]').value = 'table_header';
  el.querySelector('[data-field="format"]').value = 'table';
  syncReviewBlockState(el);
}

function markReviewBlockAsTableRow(el) {
  el.querySelector('[data-field="type"]').value = 'table';
  el.querySelector('[data-field="editorialRole"]').value = 'table_row';
  el.querySelector('[data-field="format"]').value = 'table';
  syncReviewBlockState(el);
}

function markReviewBlockAsNoise(el) {
  el.querySelector('[data-field="enabled"]').checked = false;
  el.querySelector('[data-field="type"]').value = 'generic';
  el.querySelector('[data-field="title"]').value = '';
  el.querySelector('[data-field="editorialRole"]').value = 'noise';
  syncReviewBlockState(el);
}

function splitReviewBlock(el) {
  const textArea = el.querySelector('[data-field="text"]');
  const text = normalizeContentText(textArea?.value || '');
  const lines = text.split('\n').map(line => line.trim()).filter(Boolean);
  if (lines.length < 2) return;
  const response = window.prompt('Separar a partir de qual linha? (2 até ' + lines.length + ')', String(Math.ceil(lines.length / 2)));
  const splitAt = Number(response);
  if (!Number.isInteger(splitAt) || splitAt <= 1 || splitAt > lines.length) return;
  const first = lines.slice(0, splitAt - 1).join('\n');
  const second = lines.slice(splitAt - 1).join('\n');
  textArea.value = normalizeContentText(first);
  syncReviewBlockState(el);

  const newBlock = collectSingleReviewBlock(el);
  newBlock.id = undefined;
  newBlock.title = '';
  newBlock.text = normalizeContentText(second);
  newBlock.type = 'generic';
  newBlock.editorialRole = 'body';
  insertReviewBlockAfter(el, newBlock);
}

function collectSingleReviewBlock(el) {
  return {
    id: Number(el.dataset.id),
    title: normalizeContentText(el.querySelector('[data-field="title"]')?.value || ''),
    type: el.querySelector('[data-field="type"]')?.value || 'generic',
    confidence: el.querySelector('[data-field="confidence"]')?.value || 'low',
    enabled: !!el.querySelector('[data-field="enabled"]')?.checked,
    format: el.querySelector('[data-field="format"]')?.value || 'auto',
    preserve: !!el.querySelector('[data-field="preserve"]')?.checked,
    keepHeading: !!el.querySelector('[data-field="keepHeading"]')?.checked,
    lockAutoTransform: !!el.querySelector('[data-field="lockAutoTransform"]')?.checked,
    editorialRole: el.querySelector('[data-field="editorialRole"]')?.value || 'body',
    text: normalizeContentText(el.querySelector('[data-field="text"]')?.value || ''),
  };
}

function insertReviewBlockAfter(referenceEl, block) {
  const temp = document.createElement('div');
  addReviewBlock(block);
  const inserted = els.reviewSectionsList.lastElementChild;
  if (referenceEl.nextSibling) {
    els.reviewSectionsList.insertBefore(inserted, referenceEl.nextSibling);
  }
}

function syncReviewBlockState(el) {
  const item = state.reviewBlocks.find(block => block.id === Number(el.dataset.id));
  if (!item) return;
  Object.assign(item, collectSingleReviewBlock(el));
}

/* Aplica o conteúdo revisado ao formulário principal e gera o documento */
function applyReviewToTemplate() {
  els.fldTitle.value     = normalizeContentText(els.rv_title.value);
  els.fldSubtitle.value  = normalizeContentText(els.rv_subtitle.value);
  els.fldDate.value      = els.rv_date.value;
  els.fldVersion.value   = els.rv_version.value;
  els.fldAuthor.value    = normalizeContentText(els.rv_author.value);
  els.fldDept.value      = normalizeContentText(els.rv_dept.value);
  els.fldSummary.value   = normalizeContentText(els.rv_summary.value);
  els.fldNextSteps.value = normalizeContentText(els.rv_nextSteps.value);

  const type = els.rv_docType.value;
  document.querySelectorAll('.doc-type').forEach(b => b.classList.toggle('active', b.dataset.type === type));
  state.docType = type;
  state.structureMode = normalizeStructureMode(els.rv_structureMode.value, type);

  state.sections = [];
  state.sectionIdCounter = 0;
  els.sectionsList.innerHTML = '';
  state.highlights = [];
  state.highlightIdCounter = 0;
  els.highlightsList.innerHTML = '';
  state.documentModel = [];

  const blocks = collectReviewBlocks();
  const safeSummary = normalizeContentText(els.rv_summary.value);
  const safeNextSteps = normalizeStepLines(normalizeContentText(els.rv_nextSteps.value)).join('\n');

  els.fldSummary.value = safeSummary;
  els.fldNextSteps.value = safeNextSteps;

  const structure = buildDocumentStructureFromReview(blocks, {
    docType: type,
    structureMode: state.structureMode,
    summary: safeSummary,
    nextSteps: safeNextSteps,
  });
  state.documentModel = structure.model;
  structure.sections.forEach(section => addSection(section.title, section.body));
  structure.highlights.forEach(item => addHighlight(item.label, item.value, item.desc));

  if (!els.fldSummary.value.trim()) {
    els.fldSummary.value = deriveSummaryFromBlocks(blocks);
  }

  if (!els.fldNextSteps.value.trim()) {
    els.fldNextSteps.value = deriveNextStepsFromBlocks(blocks);
  }

  closeReview();
  generateDocument();
}

function collectReviewBlocks() {
  return Array.from(els.reviewSectionsList.querySelectorAll('.review-sec-item')).map(el => ({
    id: Number(el.dataset.id),
    title: normalizeContentText(el.querySelector('[data-field="title"]')?.value || ''),
    type: el.querySelector('[data-field="type"]')?.value || 'generic',
    confidence: el.querySelector('[data-field="confidence"]')?.value || 'low',
    enabled: !!el.querySelector('[data-field="enabled"]')?.checked,
    format: el.querySelector('[data-field="format"]')?.value || 'auto',
    preserve: !!el.querySelector('[data-field="preserve"]')?.checked,
    keepHeading: !!el.querySelector('[data-field="keepHeading"]')?.checked,
    lockAutoTransform: !!el.querySelector('[data-field="lockAutoTransform"]')?.checked,
    editorialRole: el.querySelector('[data-field="editorialRole"]')?.value || 'body',
    text: normalizeContentText(el.querySelector('[data-field="text"]')?.value || ''),
  })).filter(block => block.editorialRole !== 'noise' && (block.text || block.title));
}

/* ──────────────────────────────────────────────────────────────
   Utilitário: loading overlay
   ────────────────────────────────────────────────────────────── */
function showLoading(msg) {
  els.loadingLabel.textContent = msg || 'Processando…';
  els.loadingOverlay.style.display = 'flex';
}
function setLoading(msg) { els.loadingLabel.textContent = msg; }
function hideLoading() { els.loadingOverlay.style.display = 'none'; }

function showWarning(msg) {
  // Mini toast não-intrusivo
  const toast = document.createElement('div');
  toast.style.cssText = 'position:fixed;bottom:24px;left:50%;transform:translateX(-50%);background:#252529;border:1px solid rgba(229,185,43,.4);color:rgba(229,185,43,.9);padding:10px 20px;border-radius:12px;font-size:.8rem;z-index:9999;font-family:Nunito,sans-serif;font-weight:700;box-shadow:0 8px 24px rgba(0,0,0,.4)';
  toast.textContent = msg;
  document.body.appendChild(toast);
  setTimeout(() => toast.remove(), 4000);
}

function filenameToTitle(filename) {
  return (filename || '').replace(/\.pdf$/i, '').replace(/[_\-]+/g, ' ').trim();
}

/* ══════════════════════════════════════════════════════════════
   MÓDULO: GERENCIAMENTO DE SEÇÕES (formulário principal)
   ══════════════════════════════════════════════════════════════ */
function addSection(title = '', body = '') {
  const id = ++state.sectionIdCounter;
  const item = { id, title, body };
  state.sections.push(item);
  renderSectionItem(item, els.sectionsList, 'sections');
  return item;
}

function addHighlight(label = '', value = '', desc = '') {
  const id = ++state.highlightIdCounter;
  const item = { id, label, value, desc };
  state.highlights.push(item);
  renderSectionItem(item, els.highlightsList, 'highlights');
  return item;
}

function renderSectionItem(item, container, listKey) {
  const isHighlight = listKey === 'highlights';
  const el = document.createElement('div');
  el.className = 'sec-item';
  el.dataset.id = item.id;

  if (isHighlight) {
    el.innerHTML = `
      <div class="sec-item__head">
        <input class="sec-item__name" type="text" placeholder="Label (ex: Usuários Ativos)"
               value="${escHtml(item.label)}" data-field="label"/>
        <button class="sec-item__del" title="Remover">✕</button>
      </div>
      <div class="sec-item__body">
        <input type="text" placeholder="Valor (ex: +347)" value="${escHtml(item.value)}" data-field="value"/>
        <textarea placeholder="Descrição adicional (opcional)" data-field="desc">${escHtml(item.desc)}</textarea>
      </div>`;
  } else {
    el.innerHTML = `
      <div class="sec-item__head">
        <input class="sec-item__name" type="text" placeholder="Título da seção"
               value="${escHtml(item.title)}" data-field="title"/>
        <button class="sec-item__del" title="Remover">✕</button>
      </div>
      <div class="sec-item__body">
        <textarea placeholder="Conteúdo da seção..." data-field="body">${escHtml(item.body)}</textarea>
      </div>`;
  }

  el.querySelector('.sec-item__del').addEventListener('click', () => {
    state[listKey] = state[listKey].filter(s => s.id !== item.id);
    el.remove();
  });
  el.querySelectorAll('[data-field]').forEach(input => {
    input.addEventListener('input', () => {
      const target = state[listKey].find(s => s.id === item.id);
      if (target) target[input.dataset.field] = input.value;
    });
  });

  container.appendChild(el);
}

/* ══════════════════════════════════════════════════════════════
   MÓDULO: GERAÇÃO DO DOCUMENTO
   ══════════════════════════════════════════════════════════════ */
function generateDocument() {
  const config = readConfig();
  const html = buildDocHTML(config);

  els.stageHint.style.display = 'none';
  els.docWrap.style.display = 'block';
  els.docWrap.innerHTML = html;

  // Feedback no botão
  const btn = els.generateBtn;
  const orig = btn.innerHTML;
  btn.innerHTML = '✓ Documento Gerado';
  btn.style.opacity = '.85';
  setTimeout(() => { btn.innerHTML = orig; btn.style.opacity = ''; }, 1800);
}

function readConfig() {
  const sections = normalizeSections(state.sections);
  const highlights = normalizeHighlights(state.highlights);
  const normalizedNextSteps = normalizeStepLines(els.fldNextSteps.value);
  const documentModel = normalizeDocumentModel(state.documentModel);

  return {
    docType:    state.docType,
    typeLabel:  DOC_TYPE_LABELS[state.docType] || 'DOCUMENTO OFICIAL',
    schema:     DOC_TYPE_SCHEMA[state.docType] || DOC_TYPE_SCHEMA.relatorio,
    structureMode: state.structureMode,
    theme:      state.theme,
    title:      normalizeContentText(els.fldTitle.value)    || 'Official Portfolio Document',
    subtitle:   normalizeContentText(els.fldSubtitle.value) || '',
    date:       els.fldDate.value.trim()     || formatDate(),
    version:    els.fldVersion.value.trim()  || 'v1.0',
    author:     normalizeContentText(els.fldAuthor.value)   || 'Strategy Team',
    dept:       normalizeContentText(els.fldDept.value)     || '',
    summary:    normalizeContentText(els.fldSummary.value)  || '',
    sections,
    highlights,
    documentModel,
    nextSteps:  normalizedNextSteps.join('\n'),
    nextStepItems: normalizedNextSteps,
  };
}

/* ══════════════════════════════════════════════════════════════
   MÓDULO: BUILD HTML DO DOCUMENTO
   ══════════════════════════════════════════════════════════════ */
function buildDocHTML(cfg) {
  const { theme, docType, schema } = cfg;
  const lv = logoVariant(theme);
  const contextual = buildContextualData(cfg);
  const bodyHTML = buildDocTypeBody(cfg, contextual);

  const headerHTML = theme === 'pdf'
    ? buildPdfHeaderHTML(cfg, lv)
    : buildStandardHeaderHTML(cfg, lv);

  const signatureHTML = schema.signature ? buildSignatureHTML(cfg) : '';

  return `
<div class="gudi-doc" data-theme="${theme}" data-doctype="${docType}">
  <div class="doc-layout">

    <div class="doc-sidebar">
      <div class="doc-sidebar__dots">
        <div class="doc-sidebar__dot"></div>
        <div class="doc-sidebar__dot"></div>
        <div class="doc-sidebar__dot"></div>
      </div>
      <div class="doc-sidebar__bar"></div>
      <span class="doc-sidebar__label">DOCS</span>
    </div>

    <div class="doc-main">

      <header class="doc-header">
        ${headerHTML}
      </header>

      <div class="doc-body">
        ${bodyHTML}
        ${signatureHTML}
      </div>

      ${SVG_SWIRL_ORNAMENT}

      <footer class="doc-footer">
        <div class="doc-footer__inner">
          <div class="doc-footer__brand">
            ${buildLogoSVG(lv, 'footer')}
            <span class="doc-footer__tagline">viva o bom da vida</span>
          </div>
          <div class="doc-footer__meta">
            <span>${escHtml(cfg.typeLabel)}</span>
            <span class="doc-footer__dot">·</span>
            <span>${escHtml(cfg.date)}</span>
            <span class="doc-footer__dot">·</span>
            <span>${escHtml(cfg.version)}</span>
          </div>
          <div class="doc-footer__url">example.com</div>
        </div>
      </footer>

    </div>
  </div>
</div>`;
}

function buildDocTypeBody(cfg, contextual) {
  if (cfg.documentModel.length) {
    return buildStructuredDocumentBody(cfg, contextual);
  }

  const sectionsHTML = cfg.sections.length ? buildSectionsHTML(cfg.sections, cfg.docType) : '';
  const summaryHTML = cfg.summary ? buildSummaryHTML(cfg.summary, cfg.schema.summaryLabel) : '';
  const highlightsHTML = cfg.highlights.length ? buildHighlightsHTML(cfg.highlights, cfg.schema.highlightsLabel, cfg.docType) : '';
  const nextStepsHTML = cfg.nextStepItems.length ? buildNextStepsHTML(cfg.nextStepItems, cfg.schema.nextLabel, cfg.docType) : '';

  if (cfg.docType === 'parecer') {
    return [
      cfg.schema.intro ? `<div class="doc-parecer-intro">${escHtml(cfg.schema.intro)}</div>` : '',
      summaryHTML,
      contextual.parecerMatrix,
      contextual.metaStrip,
      sectionsHTML,
      nextStepsHTML,
    ].filter(Boolean).join('');
  }

  if (cfg.docType === 'diagnostico') {
    return [
      summaryHTML,
      contextual.diagnosticHero,
      contextual.diagnosticGrid,
      sectionsHTML,
      nextStepsHTML,
    ].filter(Boolean).join('');
  }

  if (cfg.docType === 'proposta') {
    return [
      summaryHTML,
      contextual.proposalMeta,
      contextual.proposalRoadmap,
      sectionsHTML,
      nextStepsHTML,
    ].filter(Boolean).join('');
  }

  if (cfg.docType === 'sintese') {
    return [
      summaryHTML,
      contextual.metaStrip,
      contextual.strategicCallouts,
      sectionsHTML,
      nextStepsHTML,
    ].filter(Boolean).join('');
  }

  if (cfg.docType === 'impacto') {
    return [
      summaryHTML,
      highlightsHTML,
      contextual.impactNarrative,
      contextual.impactTimeline,
      sectionsHTML,
      nextStepsHTML,
    ].filter(Boolean).join('');
  }

  if (cfg.docType === 'crm') {
    return [
      summaryHTML,
      contextual.crmDeck,
      sectionsHTML,
      nextStepsHTML,
    ].filter(Boolean).join('');
  }

  return [
    summaryHTML,
    highlightsHTML,
    contextual.metaStrip,
    (summaryHTML || highlightsHTML) && sectionsHTML ? '<div class="doc-sep"></div>' : '',
    sectionsHTML,
    nextStepsHTML,
  ].filter(Boolean).join('');
}

function buildContextualData(cfg) {
  return {
    metaStrip: buildMetaStripHTML(cfg),
    parecerMatrix: buildParecerMatrixHTML(cfg),
    diagnosticHero: buildDiagnosticHeroHTML(cfg),
    diagnosticGrid: buildDiagnosticGridHTML(cfg),
    proposalMeta: buildProposalMetaHTML(cfg),
    proposalRoadmap: buildProposalRoadmapHTML(cfg),
    strategicCallouts: buildStrategicCalloutsHTML(cfg),
    impactNarrative: buildImpactNarrativeHTML(cfg),
    impactTimeline: buildImpactTimelineHTML(cfg),
    crmDeck: buildCRMBlockHTML(cfg),
  };
}

function buildMetaStripHTML(cfg) {
  const items = [
    cfg.author ? ['Autoria', cfg.author] : null,
    cfg.dept ? ['Área', cfg.dept] : null,
    cfg.date ? ['Data', cfg.date] : null,
    cfg.version ? ['Versão', cfg.version] : null,
  ].filter(Boolean);

  if (items.length < 2) return '';

  return `
<div class="doc-meta-strip">
  ${items.map(([label, value]) => `
  <div class="doc-meta-strip__item">
    <div class="doc-meta-strip__label">${escHtml(label)}</div>
    <div class="doc-meta-strip__value">${escHtml(value)}</div>
  </div>`).join('')}
</div>`;
}

function buildParecerMatrixHTML(cfg) {
  const cards = cfg.highlights.slice(0, 4).map((item, index) => `
  <div class="doc-opinion-card">
    <div class="doc-opinion-card__kicker">Ponto ${String(index + 1).padStart(2, '0')}</div>
    <div class="doc-opinion-card__title">${escHtml(item.label || item.value || 'Critério avaliado')}</div>
    ${(item.value || item.desc) ? `<div class="doc-opinion-card__body">${escHtml([item.value, item.desc].filter(Boolean).join(' · '))}</div>` : ''}
  </div>`).join('');

  return cards ? `<div class="doc-opinion-grid">${cards}</div>` : '';
}

function buildDiagnosticHeroHTML(cfg) {
  const leadSection = cfg.sections[0];
  const supportSection = cfg.sections[1];
  const insight = leadSection?.body?.split('\n')[0] || cfg.summary || '';
  const support = supportSection?.body?.split('\n')[0] || '';
  if (!insight && !support) return '';

  return `
<div class="doc-diagnostic-hero">
  ${insight ? `<div class="doc-diagnostic-hero__main">${escHtml(insight.slice(0, 240))}</div>` : ''}
  ${support ? `<div class="doc-diagnostic-hero__side">${escHtml(support.slice(0, 180))}</div>` : ''}
</div>`;
}

function buildDiagnosticGridHTML(cfg) {
  if (cfg.sections.length < 2) return '';
  const cards = cfg.sections.slice(0, 4).map(s => `
<div class="doc-diag-card">
  <div class="doc-diag-card__label">${escHtml(s.title)}</div>
  <div class="doc-diag-card__text">${escHtml((s.body || '').split('\n')[0].slice(0, 180))}</div>
</div>`).join('');
  return cards ? `<div class="doc-diagnostico-grid">${cards}</div>` : '';
}

function buildProposalMetaHTML(cfg) {
  const sec = cfg.sections;
  const cronSection = sec.find(s => /cronograma|prazo|data/i.test(s.title));
  const invSection = sec.find(s => /investimento|valor|custo|budget/i.test(s.title));
  const escSection = sec.find(s => /escopo|objetivo/i.test(s.title));
  const values = [
    ['Escopo', escSection ? firstMeaningfulLine(escSection.body, 80) : ''],
    ['Prazo', cronSection ? firstMeaningfulLine(cronSection.body, 44) : ''],
    ['Investimento', invSection ? firstMeaningfulLine(invSection.body, 44) : ''],
  ].filter(([, value]) => value);

  if (!values.length) return '';

  return `
<div class="doc-proposta-meta">
  ${values.map(([label, value]) => `
  <div class="doc-proposta-meta-item">
    <div class="doc-proposta-meta-label">${escHtml(label)}</div>
    <div class="doc-proposta-meta-value">${escHtml(value)}</div>
  </div>`).join('')}
</div>`;
}

function buildProposalRoadmapHTML(cfg) {
  const roadmapSections = cfg.sections.filter(s => /cronograma|entreg|etapa|fase|prazo/i.test(s.title)).slice(0, 4);
  if (!roadmapSections.length) return '';
  return `
<div class="doc-roadmap">
  <div class="doc-roadmap__label">Estrutura da Proposta</div>
  <div class="doc-roadmap__items">
    ${roadmapSections.map((section, index) => `
    <div class="doc-roadmap__item">
      <div class="doc-roadmap__step">${String(index + 1).padStart(2, '0')}</div>
      <div>
        <div class="doc-roadmap__title">${escHtml(section.title)}</div>
        <div class="doc-roadmap__text">${escHtml(firstMeaningfulLine(section.body, 150))}</div>
      </div>
    </div>`).join('')}
  </div>
</div>`;
}

function buildStrategicCalloutsHTML(cfg) {
  const sections = cfg.sections.slice(0, 3);
  if (!sections.length) return '';
  return `
<div class="doc-strategy-stack">
  ${sections.map((section, index) => `
  <div class="doc-strategy-card">
    <div class="doc-strategy-card__eyebrow">Prioridade ${index + 1}</div>
    <div class="doc-strategy-card__title">${escHtml(section.title)}</div>
    <div class="doc-strategy-card__text">${escHtml(firstMeaningfulLine(section.body, 180))}</div>
  </div>`).join('')}
</div>`;
}

function buildImpactNarrativeHTML(cfg) {
  const metrics = cfg.highlights.slice(0, 3);
  const section = cfg.sections.find(s => /resultado|impacto|repercuss|efeito|alcance/i.test(s.title)) || cfg.sections[0];
  if (!metrics.length && !section) return '';

  return `
<div class="doc-impact-panel">
  ${section ? `<div class="doc-impact-panel__summary">${escHtml(firstMeaningfulLine(section.body, 220))}</div>` : ''}
  ${metrics.length ? `<div class="doc-impact-panel__metrics">
    ${metrics.map(item => `
    <div class="doc-impact-panel__metric">
      <div class="doc-impact-panel__metric-value">${escHtml(item.value || '—')}</div>
      <div class="doc-impact-panel__metric-label">${escHtml(item.label || 'Indicador')}</div>
    </div>`).join('')}
  </div>` : ''}
</div>`;
}

function buildImpactTimelineHTML(cfg) {
  const timelineSections = cfg.sections.slice(0, 4);
  if (timelineSections.length < 2) return '';
  return `
<div class="doc-impact-timeline">
  <div class="doc-impact-timeline__label">Evolução do Impacto</div>
  ${timelineSections.map((section, index) => `
  <div class="doc-impact-timeline__item">
    <div class="doc-impact-timeline__index">${String(index + 1).padStart(2, '0')}</div>
    <div class="doc-impact-timeline__content">
      <div class="doc-impact-timeline__title">${escHtml(section.title)}</div>
      <div class="doc-impact-timeline__text">${escHtml(firstMeaningfulLine(section.body, 160))}</div>
    </div>
  </div>`).join('')}
</div>`;
}

function buildCRMBlockHTML(cfg) {
  const kpiHTML = cfg.highlights.length ? buildCRMKPIs(cfg.highlights) : '';
  const tempHTML = buildCRMSegments(cfg.sections, cfg.highlights);
  const funnelHTML = cfg.highlights.length >= 2 ? buildCRMFunnel(cfg.highlights) : '';
  const tableHTML = buildCRMTable(cfg.sections, cfg.highlights);
  const priorityHTML = buildCRMPriorityAudiences(cfg.sections);
  const recommendationsHTML = buildCRMRecommendations(cfg.nextStepItems);
  const insightHTML = cfg.summary ? buildCRMInsight(cfg.summary) : '';

  if (!kpiHTML && !tempHTML && !funnelHTML && !tableHTML && !priorityHTML && !recommendationsHTML && !insightHTML) {
    return '';
  }

  return `
<div class="doc-crm-block">
  ${insightHTML}
  <div class="doc-crm-grid">
    ${kpiHTML}
    ${funnelHTML}
  </div>
  ${tempHTML}
  ${tableHTML}
  ${priorityHTML}
  ${recommendationsHTML}
</div>`;
}

/* KPIs de temperatura: transforma highlights em cards com classificação */
function buildCRMKPIs(highlights) {
  const temps = ['hot', 'warm', 'mid', 'cool', 'cold', 'base'];

  const cards = highlights.map((h, i) => {
    const cls = temps[i % temps.length];
    return `
<div class="doc-crm-kpi doc-crm-kpi--${cls}">
  <span class="doc-crm-kpi__dot"></span>
  <div class="doc-crm-kpi__label">${escHtml(h.label || 'Indicador')}</div>
  <div class="doc-crm-kpi__value">${escHtml(h.value || '—')}</div>
  ${h.desc ? `<div class="doc-crm-kpi__desc">${escHtml(h.desc)}</div>` : ''}
</div>`;
  }).join('');

  return `
<div>
  <div class="doc-crm-section-label">Indicadores-Chave</div>
  <div class="doc-crm-kpi-grid">${cards}</div>
</div>`;
}

/* Segmentos: usa seções para montar blocos de temperatura */
function buildCRMSegments(sections, highlights) {
  if (!sections || sections.length === 0) return '';

  // Mapeia seções para temperaturas baseado no título
  const tempMap = {
    'hot':  /quente|aprovado|elite|núcleo|urgente|prioridade\s*máx/i,
    'warm': /morno|profissional|intenção|alta\s+intenção/i,
    'mid':  /médio|suplente|moderado/i,
    'cool': /frio.qualificado|qualificado/i,
    'cold': /frio.impulsivo|impulsivo|pico/i,
    'base': /frio.base|base|geral|amplo/i,
  };

  const segs = sections.slice(0, 6).map((sec, i) => {
    let cls = 'base';
    for (const [key, rx] of Object.entries(tempMap)) {
      if (rx.test(sec.title)) { cls = key; break; }
    }
    // Usa highlight correspondente para volume, se existir
    const vol = highlights[i]?.value || '';
    const volLabel = highlights[i]?.label || '';
    const bodyFirst = (sec.body || '').split('\n')[0].slice(0, 120);

    return `
<div class="doc-crm-seg doc-crm-seg--${cls}">
  <div class="doc-crm-seg__temp">
    <span class="doc-crm-seg__badge">●</span>
    ${escHtml(sec.title.split('—')[0].split('–')[0].trim().slice(0, 20))}
  </div>
  <div class="doc-crm-seg__info">
    <div class="doc-crm-seg__name">${escHtml(sec.title)}</div>
    ${bodyFirst ? `<div class="doc-crm-seg__desc">${escHtml(bodyFirst)}</div>` : ''}
  </div>
  ${vol ? `<div class="doc-crm-seg__vol">${escHtml(vol)}<span>${escHtml(volLabel)}</span></div>` : ''}
</div>`;
  }).join('');

  return `
<div>
  <div class="doc-crm-section-label">Temperatura e Segmentação da Base</div>
  <div class="doc-crm-segments">${segs}</div>
</div>`;
}

/* Funil CRM: barras de progresso com os highlights */
function buildCRMFunnel(highlights) {
  if (highlights.length < 2) return '';

  // Calcula percentuais relativos ao primeiro valor (base total)
  const vals = highlights.map(h => {
    const n = parseInt((h.value || '0').replace(/[^\d]/g, '')) || 0;
    return { label: h.label || '', val: n, desc: h.desc || '' };
  });
  const max = Math.max(...vals.map(v => v.val), 1);
  const colors = ['#2D4F76', '#2ED9C3', '#81C458', '#E5B92B', '#E16E1A'];

  const stages = vals.map((v, i) => {
    const pct = max > 0 ? Math.round((v.val / max) * 100) : (100 - i * 18);
    const color = colors[Math.min(i, colors.length - 1)];
    return `
<div class="doc-crm-stage">
  <span class="doc-crm-stage__label">${escHtml(v.label)}</span>
  <span class="doc-crm-stage__bar"><span class="doc-crm-stage__fill" style="width:${pct}%;background:${color}"></span></span>
  <span class="doc-crm-stage__val">${escHtml(highlights[i].value || '—')}</span>
</div>`;
  }).join('');

  return `
<div>
  <div class="doc-crm-section-label">Funil de Engajamento</div>
  <div class="doc-crm-funnel">${stages}</div>
</div>`;
}

/* Insight CRM: destaca dado analítico principal do resumo */
function buildCRMInsight(summary) {
  if (!summary || summary.length < 20) return '';
  // Pega a primeira frase com dado numérico ou conclusão
  const sentences = summary.split(/[.!?·]\s+/).filter(s => s.length > 20);
  const insight = sentences.slice(0, 2).join('. ').slice(0, 320);

  return `
<div class="doc-crm-insight">
  <div class="doc-crm-insight__label">Insight Analítico</div>
  <div class="doc-crm-insight__text">${escHtml(insight)}</div>
</div>`;
}

function buildCRMTable(sections, highlights) {
  const rows = sections.slice(0, 5).map((section, index) => {
    const volume = highlights[index]?.value || '—';
    const priority = crmPriorityTag(section.title, section.body);
    const hypothesis = firstMeaningfulLine(section.body, 110);
    if (!section.title && !hypothesis) return '';
    return `
    <tr>
      <td>${escHtml(section.title || `Segmento ${index + 1}`)}</td>
      <td>${escHtml(volume)}</td>
      <td><span class="doc-crm-tag doc-crm-tag--${priority.className}">${escHtml(priority.label)}</span></td>
      <td>${escHtml(hypothesis || 'Sem leitura consolidada.')}</td>
    </tr>`;
  }).filter(Boolean).join('');

  if (!rows) return '';

  return `
<div>
  <div class="doc-crm-section-label">Tabela Analítica</div>
  <div class="doc-crm-table-wrap">
    <table class="doc-crm-table">
      <thead>
        <tr>
          <th>Segmento</th>
          <th>Volume</th>
          <th>Prioridade</th>
          <th>Leitura</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  </div>
</div>`;
}

function buildCRMPriorityAudiences(sections) {
  const prioritySections = sections
    .filter(section => /priorit|quente|valor|segment|público|aud[ií]enc/i.test(`${section.title} ${section.body}`))
    .slice(0, 4);

  if (!prioritySections.length) return '';

  return `
<div>
  <div class="doc-crm-section-label">Públicos Prioritários</div>
  <div class="doc-crm-priority-grid">
    ${prioritySections.map((section, index) => `
    <div class="doc-crm-priority-card">
      <div class="doc-crm-priority-card__rank">${String(index + 1).padStart(2, '0')}</div>
      <div class="doc-crm-priority-card__title">${escHtml(section.title)}</div>
      <div class="doc-crm-priority-card__text">${escHtml(firstMeaningfulLine(section.body, 150))}</div>
    </div>`).join('')}
  </div>
</div>`;
}

function buildCRMRecommendations(nextStepItems) {
  if (!nextStepItems?.length) return '';
  return `
<div>
  <div class="doc-crm-section-label">Recomendações e Próximos Movimentos</div>
  <div class="doc-crm-reco-list">
    ${nextStepItems.slice(0, 6).map((step, index) => `
    <div class="doc-crm-reco-item">
      <div class="doc-crm-reco-item__index">${String(index + 1).padStart(2, '0')}</div>
      <div class="doc-crm-reco-item__text">${escHtml(step)}</div>
    </div>`).join('')}
  </div>
</div>`;
}

function buildSignatureHTML(cfg) {
  return `
<div class="doc-signature">
  <div class="doc-signature__line"></div>
  <div class="doc-signature__name">${escHtml(cfg.author || 'Responsável Técnico')}</div>
  <div class="doc-signature__role">${escHtml(cfg.dept || 'Portfolio Operations')} · ${escHtml(cfg.date)}</div>
</div>`;
}

/* ──────────────────────────────────────────────────────────────
   BUILDERS DE HEADER
   ────────────────────────────────────────────────────────────── */
function buildStandardHeaderHTML(cfg, lv) {
  const headerMeta = [
    cfg.author ? ['Autoria', cfg.author] : null,
    cfg.dept ? ['Área', cfg.dept] : null,
    cfg.date ? ['Data', cfg.date] : null,
    cfg.version ? ['Versão', cfg.version] : null,
  ].filter(Boolean);

  return `
    ${SVG_STAR_ORNAMENT}
    <div class="doc-header__standard">
      <div class="doc-header__top">
        <div class="doc-header__brand">
          ${buildLogoSVG(lv, 'header')}
          <span class="doc-header__divider"></span>
          <span class="doc-header__sub-brand">viva o bom da vida<br>Documentos Oficiais</span>
        </div>
        <div class="doc-header__meta">
          <span class="doc-header__badge">${escHtml(cfg.typeLabel)}</span>
          <div class="doc-header__date-ver">
            <span>${escHtml(cfg.date)}</span>
            <span>${escHtml(cfg.version)}</span>
          </div>
        </div>
      </div>
      <div class="doc-header__title-block">
        <h1 class="doc-header__title">${escHtml(cfg.title)}</h1>
        ${cfg.subtitle ? `<p class="doc-header__subtitle">${escHtml(normalizeContentText(cfg.subtitle))}</p>` : ''}
      </div>
      ${headerMeta.length ? `<div class="doc-header__meta-row">
        ${headerMeta.map(([label, value]) => `
        <div class="doc-header__meta-item">
          <span class="doc-header__meta-label">${escHtml(label)}</span>
          <span class="doc-header__meta-value">${escHtml(value)}</span>
        </div>`).join('')}
      </div>` : ''}
    </div>`;
}

function buildPdfHeaderHTML(cfg, lv) {
  return `
    <div class="doc-header__pdf-bar"></div>
    <div class="doc-header__pdf-inner">
      <div class="doc-header__pdf-brand">
        ${buildLogoSVG(lv, 'header')}
        <span class="doc-header__pdf-sep"></span>
        <span class="doc-header__pdf-tagline">viva o bom da vida</span>
      </div>
      <div class="doc-header__pdf-type">${escHtml(cfg.typeLabel)}</div>
      <h1 class="doc-header__pdf-title">${escHtml(cfg.title)}</h1>
      ${cfg.subtitle ? `<p class="doc-header__pdf-subtitle">${escHtml(normalizeContentText(cfg.subtitle))}</p>` : ''}
      <div class="doc-header__pdf-meta">
        ${cfg.author ? `<div class="doc-header__pdf-meta-item">
          <span class="doc-header__pdf-meta-label">Autoria</span>
          <span class="doc-header__pdf-meta-value">${escHtml(cfg.author)}</span>
        </div>` : ''}
        ${cfg.dept ? `<div class="doc-header__pdf-meta-item">
          <span class="doc-header__pdf-meta-label">Área</span>
          <span class="doc-header__pdf-meta-value">${escHtml(cfg.dept)}</span>
        </div>` : ''}
        <div class="doc-header__pdf-meta-item">
          <span class="doc-header__pdf-meta-label">Data</span>
          <span class="doc-header__pdf-meta-value">${escHtml(cfg.date)}</span>
        </div>
        <div class="doc-header__pdf-meta-item">
          <span class="doc-header__pdf-meta-label">Versão</span>
          <span class="doc-header__pdf-meta-value">${escHtml(cfg.version)}</span>
        </div>
      </div>
      <svg class="doc-header__pdf-star" viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
        <polygon points="50,0 58,35 95,38 68,60 78,95 50,73 22,95 32,60 5,38 42,35" fill="url(#pdfStar)"/>
        <defs>
          <linearGradient id="pdfStar" x1="0%" y1="0%" x2="100%" y2="100%">
            <stop offset="0%" stop-color="#81C458"/>
            <stop offset="50%" stop-color="#2ED9C3"/>
            <stop offset="100%" stop-color="#2D4F76"/>
          </linearGradient>
        </defs>
      </svg>
    </div>`;
}

/* ──────────────────────────────────────────────────────────────
   BUILDERS DE BLOCOS DE CONTEÚDO
   ────────────────────────────────────────────────────────────── */
function buildSummaryHTML(text, label) {
  const summaryText = normalizeContentText(text);
  if (!summaryText) return '';
  return `
<div class="doc-summary">
  <div class="doc-summary__label">${escHtml(normalizeContentText(label || 'Resumo Executivo'))}</div>
  <div class="doc-summary__text">${renderTextParagraphs(summaryText, 'doc-summary__paragraph')}</div>
</div>`;
}

function buildHighlightsHTML(highlights, label) {
  const icons = ['✦', '◈', '◉', '◆', '★', '▸', '◎'];
  const items = highlights.map((h, i) => `
<div class="doc-highlight">
  <span class="doc-highlight__icon">${icons[i % icons.length]}</span>
  ${h.label ? `<div class="doc-highlight__label">${escHtml(normalizeContentText(h.label))}</div>` : ''}
  ${h.value ? `<div class="doc-highlight__value">${escHtml(normalizeContentText(h.value))}</div>` : ''}
  ${h.desc  ? `<div class="doc-highlight__desc">${escHtml(normalizeContentText(h.desc))}</div>`  : ''}
</div>`).join('');

  return `
<div class="doc-highlights-wrap">
  <div class="doc-highlights-label">${escHtml(normalizeContentText(label || 'Destaques e Indicadores'))}</div>
  <div class="doc-highlights">${items}</div>
</div>`;
}

function buildSectionsHTML(sections, docType) {
  if (docType === 'parecer') return buildParecerSectionsHTML(sections);
  if (docType === 'proposta') return buildProposalSectionsHTML(sections);
  if (docType === 'sintese') return buildStrategicSectionsHTML(sections);
  if (docType === 'impacto') return buildImpactSectionsHTML(sections);
  if (docType === 'crm') return buildCRMSectionsHTML(sections);
  if (docType === 'diagnostico') return buildDiagnosticSectionsHTML(sections);

  return buildDefaultSectionsHTML(sections);
}

function buildDefaultSectionsHTML(sections) {
  return sections.map((sec, i) => {
    const bodyParagraphs = renderTextParagraphs(sec.body);

    return `
<div class="doc-section">
  <div class="doc-section__head">
    <div class="doc-section__num">${String(i + 1).padStart(2, '0')}</div>
    <h2 class="doc-section__title">${escHtml(normalizeContentText(sec.title || `Seção ${i + 1}`))}</h2>
    <div class="doc-section__rule"></div>
  </div>
  <div class="doc-section__body">${bodyParagraphs}</div>
</div>`;
  }).join('');
}

function buildDiagnosticSectionsHTML(sections) {
  return sections.map((sec, i) => {
    const paragraphs = renderTextParagraphs(sec.body);

    return `
<article class="doc-section doc-section--diagnostico">
  <div class="doc-section__head">
    <div class="doc-section__num">${String(i + 1).padStart(2, '0')}</div>
    <div class="doc-section__diagnostic-head">
      <h2 class="doc-section__title">${escHtml(normalizeContentText(sec.title || `Achado ${i + 1}`))}</h2>
      <div class="doc-section__diagnostic-kicker">Achado estrutural</div>
    </div>
    <div class="doc-section__rule"></div>
  </div>
  <div class="doc-section__body">${paragraphs}</div>
</article>`;
  }).join('');
}

function buildParecerSectionsHTML(sections) {
  return sections.map((sec, i) => {
    const paragraphs = renderTextParagraphs(sec.body);

    return `
<article class="doc-section doc-section--parecer">
  <div class="doc-section__parecer-mark">Item ${String(i + 1).padStart(2, '0')}</div>
  <h2 class="doc-section__title">${escHtml(normalizeContentText(sec.title || `Item ${i + 1}`))}</h2>
  <div class="doc-section__body">${paragraphs}</div>
</article>`;
  }).join('');
}

function buildProposalSectionsHTML(sections) {
  return sections.map((sec, i) => {
    const paragraphs = renderTextParagraphs(sec.body);

    return `
<article class="doc-section doc-section--proposta">
  <div class="doc-section__proposal-index">${String(i + 1).padStart(2, '0')}</div>
  <div class="doc-section__proposal-main">
    <div class="doc-section__proposal-kicker">Frente de trabalho</div>
    <h2 class="doc-section__title">${escHtml(normalizeContentText(sec.title || `Leitura ${i + 1}`))}</h2>
    <div class="doc-section__body">${paragraphs}</div>
  </div>
</article>`;
  }).join('');
}

function buildStrategicSectionsHTML(sections) {
  return sections.map((sec, i) => `
<article class="doc-section doc-section--sintese-card">
  <div class="doc-section__strategy-order">P${i + 1}</div>
  <div class="doc-section__strategy-main">
    <h2 class="doc-section__title">${escHtml(normalizeContentText(sec.title || `Prioridade ${i + 1}`))}</h2>
    <div class="doc-section__strategy-text">${escHtml(normalizeContentText(firstMeaningfulLine(sec.body, 220)))}</div>
  </div>
</article>`).join('');
}

function buildImpactSectionsHTML(sections) {
  return sections.map((sec, i) => {
    const paragraphs = renderTextParagraphs(sec.body);

    return `
<article class="doc-section doc-section--impacto">
  <div class="doc-section__impacto-top">
    <div class="doc-section__impacto-index">${String(i + 1).padStart(2, '0')}</div>
    <div>
      <div class="doc-section__impacto-kicker">Evidência</div>
      <h2 class="doc-section__title">${escHtml(normalizeContentText(sec.title || `Impacto ${i + 1}`))}</h2>
    </div>
  </div>
  <div class="doc-section__body">${paragraphs}</div>
</article>`;
  }).join('');
}

function buildCRMSectionsHTML(sections) {
  return sections.map((sec, i) => `
<article class="doc-section doc-section--crm">
  <div class="doc-section__crm-head">
    <div class="doc-section__crm-index">${String(i + 1).padStart(2, '0')}</div>
    <h2 class="doc-section__title">${escHtml(normalizeContentText(sec.title || `Frente ${i + 1}`))}</h2>
  </div>
  <div class="doc-section__crm-text">${escHtml(normalizeContentText(firstMeaningfulLine(sec.body, 240)))}</div>
</article>`).join('');
}

function buildNextStepsHTML(text, label, docType = 'relatorio') {
  const steps = Array.isArray(text) ? text : normalizeStepLines(text);
  if (!steps.length) return '';
  const safeLabel = escHtml(normalizeContentText(label || 'Próximos Passos'));

  if (docType === 'proposta') {
    return `
<div class="doc-next doc-next--roadmap">
  <div class="doc-next__label">${safeLabel}</div>
  <div class="doc-next__roadmap">
    ${steps.map((step, i) => `
    <div class="doc-next__roadmap-item">
      <div class="doc-next__roadmap-step">${String(i + 1).padStart(2, '0')}</div>
      <div class="doc-next__roadmap-text">${escHtml(step)}</div>
    </div>`).join('')}
  </div>
</div>`;
  }

  if (docType === 'parecer') {
    return `
<div class="doc-next doc-next--formal">
  <div class="doc-next__label">${safeLabel}</div>
  <div class="doc-next__formal-list">
    ${steps.map((step, i) => `
    <div class="doc-next__formal-item">
      <span class="doc-next__formal-marker">${String(i + 1).padStart(2, '0')}</span>
      <span class="doc-next__formal-text">${escHtml(step)}</span>
    </div>`).join('')}
  </div>
</div>`;
  }

  if (docType === 'sintese' || docType === 'crm') {
    return `
<div class="doc-next doc-next--compact">
  <div class="doc-next__label">${safeLabel}</div>
  <div class="doc-next__pill-list">
    ${steps.map((step, i) => `<span class="doc-next__pill">${String(i + 1).padStart(2, '0')} · ${escHtml(step)}</span>`).join('')}
  </div>
</div>`;
  }

  const items = steps.map((step, i) => {
    const color = BULLET_COLORS[i % BULLET_COLORS.length];
    return `
<li class="doc-next__item">
  <span class="doc-next__bullet" style="background:${color}">${i + 1}</span>
  <span>${escHtml(step)}</span>
</li>`;
  }).join('');

  return `
<div class="doc-next">
  <div class="doc-next__label">${safeLabel}</div>
  <ul class="doc-next__list">${items}</ul>
</div>`;
}

function buildStructuredDocumentBody(cfg, contextual) {
  const leadingMeta = shouldRenderMetaStripInStructuredFlow(cfg)
    ? buildStructuredDocumentItemHTML({ kind: 'meta' }, cfg, contextual)
    : '';
  const ordered = (cfg.documentModel || [])
    .map(item => buildStructuredDocumentItemHTML(item, cfg, contextual))
    .filter(Boolean)
    .join('');

  if (!ordered) {
    return [
      cfg.summary ? buildSummaryHTML(cfg.summary, cfg.schema.summaryLabel) : '',
      cfg.sections.length ? buildSectionsHTML(cfg.sections, cfg.docType) : '',
      cfg.nextStepItems.length ? buildNextStepsHTML(cfg.nextStepItems, cfg.schema.nextLabel, cfg.docType) : '',
    ].filter(Boolean).join('');
  }

  return `<div class="doc-flow doc-flow--${cfg.structureMode}">${[leadingMeta, ordered].filter(Boolean).join('')}</div>`;
}

function buildStructuredDocumentItemHTML(item, cfg, contextual) {
  if (!item || !item.kind) return '';
  if (item.kind === 'summary') return item.text ? buildSummaryHTML(item.text, cfg.schema.summaryLabel) : '';
  if (item.kind === 'highlights') return item.items?.length ? buildHighlightsHTML(item.items, cfg.schema.highlightsLabel, cfg.docType) : '';
  if (item.kind === 'nextSteps') return item.items?.length ? buildNextStepsHTML(item.items, cfg.schema.nextLabel, cfg.docType) : '';
  if (item.kind === 'meta') return contextual.metaStrip || '';
  if (item.kind === 'block') return buildPreservedBlockHTML(item, cfg.docType);
  return '';
}

function buildPreservedBlockHTML(block, docType) {
  const title = block.keepHeading === false ? '' : normalizeContentText(block.title || '');
  if (block.mode === 'visual' && block.highlight) {
    return buildHighlightsHTML([block.highlight], title || 'Destaque', docType);
  }

  const body = buildBlockBodyHTML(block.text || '', block.mode || 'preserve', title, block.tableData || null);
  if (!body && !title) return '';

  const headingLevel = block.headingLevel || 2;
  const headingTag = headingLevel <= 2 ? 'h2' : (headingLevel === 3 ? 'h3' : 'h4');
  const levelClass = ` doc-preserved-block--level-${headingLevel}`;

  return `
<article class="doc-preserved-block doc-preserved-block--${block.mode || 'preserve'}${title ? ' doc-preserved-block--headed' : ''}${levelClass}">
  ${title ? `<div class="doc-preserved-block__head">
    ${block.number ? `<div class="doc-preserved-block__index">${escHtml(block.number)}</div>` : ''}
    <${headingTag} class="doc-preserved-block__title doc-preserved-block__title--${headingTag}">${escHtml(title)}</${headingTag}>
  </div>` : ''}
  ${body ? `<div class="doc-preserved-block__body">${body}</div>` : ''}
</article>`;
}

function buildBlockBodyHTML(text, mode = 'preserve', headingTitle = '', tableData = null) {
  const normalizedText = normalizeContentText(text);
  const lines = normalizedText.split('\n').map(line => line.trim()).filter(Boolean);
  if (!lines.length) return '';
  if (mode === 'table') {
    const parsed = tableData || extractTableStructureFromSerializedText(normalizedText);
    if (parsed) return buildTableBlockHTML(parsed);
    return buildFallbackTableHTML(lines);
  }
  if (mode === 'matrix') {
    return buildMatrixBlockHTML(lines);
  }
  if (mode === 'comparison') {
    return buildComparisonBlockHTML(lines);
  }
  if (mode === 'list') {
    return `<ul class="doc-preserved-block__list">${lines.map(line => `<li>${escHtml(stripListMarker(line))}</li>`).join('')}</ul>`;
  }
  if (mode === 'note') {
    return `<div class="doc-preserved-block__note">${lines.map(line => `<p class="doc-preserved-block__note-line">${escHtml(line)}</p>`).join('')}</div>`;
  }
  return renderStructuredBlockText(removeHeadingEcho(normalizedText, headingTitle || ''), mode);
}

function extractTableStructureFromSerializedText(text) {
  const rows = String(text || '').split('\n').map(line => line.trim()).filter(Boolean);
  if (rows.length < 2) return null;
  const separators = [
    { rx: /\s*\|\s*/, test: /\|/ },
    { rx: /\s+\/\s+/, test: /\s+\/\s+/ },
    { rx: /\s+[\u2014\u2013]{1,2}\s+/, test: /\s+[\u2014\u2013]{1,2}\s+/ },
    { rx: /\s{2,}/, test: /\s{2,}/ },
  ];
  for (const sep of separators) {
    if (!sep.test.test(rows[0])) continue;
    const header = rows[0].split(sep.rx).map(cell => cell.trim()).filter(Boolean);
    if (header.length < 2) continue;
    const bodyRows = rows.slice(1).map(row => {
      const cells = row.split(sep.rx).map(cell => cell.trim()).filter(Boolean);
      return cells.length >= 1 ? cells : [row];
    }).filter(Boolean);
    if (!bodyRows.length) continue;
    return { caption: '', captionLines: [], header, rows: bodyRows };
  }
  if (isLikelyTableHeaderLine(rows[0])) {
    const header = splitTableCells(rows[0]);
    if (header.length >= 2) {
      const bodyRows = rows.slice(1).map(row => splitTableCells(row)).filter(cells => cells.length >= 1);
      if (bodyRows.length) return { caption: '', captionLines: [], header, rows: bodyRows };
    }
  }
  return null;
}

function buildTableBlockHTML(tableData) {
  if (!tableData?.header?.length) return '';
  const head = `<tr>${tableData.header.map(cell => `<th>${escHtml(cell)}</th>`).join('')}</tr>`;
  const body = (tableData.rows || [])
    .map(row => `<tr>${row.map(cell => `<td>${escHtml(cell)}</td>`).join('')}</tr>`)
    .join('');
  return `
<div class="doc-preserved-table-wrap">
  <table class="doc-preserved-table">
    <thead>${head}</thead>
    <tbody>${body}</tbody>
  </table>
</div>`;
}

function buildFallbackTableHTML(lines) {
  if (!lines.length) return '';
  return `
<div class="doc-preserved-table-wrap doc-preserved-table-wrap--fallback">
  <table class="doc-preserved-table doc-preserved-table--fallback">
    <tbody>${lines.map((line, i) => {
      const cells = splitTableCells(line);
      if (i === 0 && isLikelyTableHeaderLine(line)) {
        return `<tr class="doc-preserved-table__header-row">${cells.map(cell => `<th>${escHtml(cell)}</th>`).join('')}</tr>`;
      }
      return `<tr>${cells.map(cell => `<td>${escHtml(cell)}</td>`).join('')}</tr>`;
    }).join('')}</tbody>
  </table>
</div>`;
}

function buildMatrixBlockHTML(lines) {
  if (!lines.length) return '';
  const pairs = lines.map(line => {
    const kvMatch = line.match(/^([^:]{3,50})\s*:\s+(.+)$/);
    if (kvMatch) return { key: kvMatch[1].trim(), value: kvMatch[2].trim() };
    const dashMatch = line.match(/^(.+?)\s+[\u2014\u2013]{1,2}\s+(.+)$/);
    if (dashMatch) return { key: dashMatch[1].trim(), value: dashMatch[2].trim() };
    const slashMatch = line.match(/^(.+?)\s+\/\s+(.+)$/);
    if (slashMatch && !isSentenceLike(line)) return { key: slashMatch[1].trim(), value: slashMatch[2].trim() };
    return { key: line, value: '' };
  });
  return `
<div class="doc-preserved-matrix">
  ${pairs.map(pair => `
  <div class="doc-preserved-matrix__row">
    <div class="doc-preserved-matrix__key">${escHtml(pair.key)}</div>
    ${pair.value ? `<div class="doc-preserved-matrix__value">${escHtml(pair.value)}</div>` : ''}
  </div>`).join('')}
</div>`;
}

function buildComparisonBlockHTML(lines) {
  if (!lines.length) return '';
  const headerLine = lines[0];
  const headerCells = splitTableCells(headerLine);
  const isHeader = isLikelyTableHeaderLine(headerLine) && headerCells.length >= 2;
  if (isHeader && lines.length >= 2) {
    const dataRows = lines.slice(1).map(line => splitTableCells(line));
    return `
<div class="doc-preserved-comparison">
  <div class="doc-preserved-comparison__header">
    ${headerCells.map(cell => `<div class="doc-preserved-comparison__col-title">${escHtml(cell)}</div>`).join('')}
  </div>
  ${dataRows.map(row => `
  <div class="doc-preserved-comparison__row">
    ${row.map((cell, i) => `<div class="doc-preserved-comparison__cell${i === 0 ? ' doc-preserved-comparison__cell--label' : ''}">${escHtml(cell)}</div>`).join('')}
  </div>`).join('')}
</div>`;
  }
  const pairs = lines.map(line => {
    const parts = splitTableCells(line);
    if (parts.length >= 2) return { left: parts[0], right: parts.slice(1).join(' \u00b7 ') };
    return { left: line, right: '' };
  });
  return `
<div class="doc-preserved-comparison doc-preserved-comparison--pairs">
  ${pairs.map(pair => `
  <div class="doc-preserved-comparison__row">
    <div class="doc-preserved-comparison__cell doc-preserved-comparison__cell--label">${escHtml(pair.left)}</div>
    ${pair.right ? `<div class="doc-preserved-comparison__cell">${escHtml(pair.right)}</div>` : ''}
  </div>`).join('')}
</div>`;
}

function inferStructureMode(docType, blocks) {
  const enabledBlocks = (blocks || []).filter(block => block && block.enabled !== false);
  const denseBlocks = enabledBlocks.filter(block => {
    const text = String(block.text || '').trim();
    return text.length >= 260 || splitTextIntoParagraphs(text).length >= 2;
  }).length;
  const visualBlocks = enabledBlocks.filter(block => block.type === 'highlight' || block.type === 'table' || block.type === 'matrix' || block.type === 'comparison' || looksLikeMetricBlock(block.text || '')).length;
  if (docType === 'parecer' || docType === 'diagnostico' || docType === 'proposta') return 'editorial';
  if (docType === 'impacto' && visualBlocks >= denseBlocks + 2 && visualBlocks >= 3) return 'visual';
  if (docType === 'crm' && visualBlocks >= denseBlocks + 2 && visualBlocks >= 3) return 'visual';
  return 'editorial';
}

function normalizeStructureMode(mode, docType) {
  if (mode === 'editorial' || mode === 'visual') return mode;
  return inferStructureMode(docType, state.reviewBlocks || []);
}

function defaultBlockFormat(type) {
  if (type === 'section') return 'section';
  if (type === 'nextSteps') return 'list';
  if (type === 'table') return 'table';
  if (type === 'matrix') return 'matrix';
  if (type === 'comparison') return 'comparison';
  if (type === 'note') return 'note';
  if (type === 'highlight') return 'preserve';
  return 'preserve';
}

function buildDocumentStructureFromReview(blocks, options = {}) {
  const enabledBlocks = (blocks || []).filter(block =>
    block.enabled !== false
    && block.editorialRole !== 'noise'
    && (String(block.text || '').trim() || String(block.title || '').trim())
  );
  const structureMode = normalizeStructureMode(options.structureMode, options.docType);
  const model = [];
  const sections = [];
  const highlights = [];
  let summaryInserted = false;
  let nextStepsInserted = false;

  enabledBlocks.forEach((block, index) => {
    const explicitRole = block.editorialRole || 'body';
    const title = explicitRole === 'body'
      ? ''
      : (normalizeContentText(block.title || '') || (block.type === 'section' ? fallbackSectionTitle(block, options.docType) : ''));
    const text = explicitRole === 'heading_only'
      ? ''
      : normalizeContentText(block.text || '');
    const format = resolveBlockFormat(block, structureMode);
    const titlelessText = title ? removeHeadingEcho(text, title) : text;
    const tableData = (block.type === 'table' || /^table_/.test(explicitRole))
      ? extractTableStructureFromSerializedText(text || block.text || '')
      : null;

    if (block.type === 'summary' && !summaryInserted) {
      model.push({ kind: 'summary', text: normalizeContentText(options.summary || deriveSummaryFromBlocks(blocks) || text) });
      summaryInserted = true;
      return;
    }

    if (block.type === 'nextSteps' && !nextStepsInserted) {
      const items = normalizeStepLines(normalizeContentText(options.nextSteps || deriveNextStepsFromBlocks(blocks) || text));
      if (items.length) {
        model.push({ kind: 'nextSteps', items });
        nextStepsInserted = true;
      }
      return;
    }

    if (shouldPromoteBlockToHighlight(block, structureMode)) {
      const highlight = extractHighlightFromBlock(block);
      if (highlight) {
        highlights.push(highlight);
        model.push({ kind: 'block', mode: 'visual', title, text: titlelessText, keepHeading: block.keepHeading !== false, number: extractSectionNumber(title) || String(index + 1).padStart(2, '0'), highlight });
        return;
      }
    }

    const headingLevel = detectHeadingLevel(title, block, index);

    model.push({
      kind: 'block',
      mode: format,
      title,
      text: titlelessText || title,
      keepHeading: block.keepHeading !== false,
      number: extractSectionNumber(title) || (title ? String(index + 1).padStart(2, '0') : ''),
      tableData,
      headingLevel,
    });

    if (shouldIncludeAsSection({ ...block, title }, format, structureMode)) {
      sections.push({ title: title || fallbackSectionTitle(block, options.docType), body: titlelessText });
    }
  });

  if (!summaryInserted && options.summary) model.unshift({ kind: 'summary', text: normalizeContentText(options.summary) });
  if (!nextStepsInserted) {
    const items = normalizeStepLines(normalizeContentText(options.nextSteps || deriveNextStepsFromBlocks(blocks)));
    if (items.length) model.push({ kind: 'nextSteps', items });
  }

  if (structureMode === 'visual' && highlights.length >= 2) {
    model.splice(summaryInserted ? 1 : 0, 0, { kind: 'highlights', items: dedupeHighlights(highlights) });
  }

  return {
    model: normalizeDocumentModel(mergeSequentialPreservedBlocks(model)),
    sections,
    highlights: dedupeHighlights(highlights),
  };
}

function resolveBlockFormat(block, structureMode) {
  if (block.lockAutoTransform) return block.format === 'auto' ? 'preserve' : block.format;
  if (block.format && block.format !== 'auto') return block.format;
  
  // Resolução mais conservadora baseada na confiança
  if (block.type === 'nextSteps') return 'list';
  if (block.type === 'table' && block.confidence === 'high') return 'table';
  if (block.type === 'matrix' && block.confidence === 'high') return 'matrix';
  if (block.type === 'comparison' && block.confidence === 'high') return 'comparison';
  if (block.type === 'note' && block.confidence !== 'low') return 'note';
  
  // Apenas promover para visual se for realmente confiável
  if (block.type === 'highlight' && structureMode === 'visual' && block.confidence === 'high') {
    return 'visual';
  }
  
  // Sections em modo editorial, senão preserve
  if (block.type === 'section') {
    return structureMode === 'editorial' ? 'section' : 'preserve';
  }
  
  // Default para preserve para evitar problemas
  return 'preserve';
}

function detectHeadingLevel(title, block, index) {
  if (!title) return 3;
  if (/^\d+\.\d+/.test(title.trim())) return 3;
  if (/^[a-z]\)/i.test(title.trim())) return 4;
  if (/^(IVXLC)+[\.)]/i.test(title.trim())) return 3;
  if (/^\d+[\.)\s]/.test(title.trim())) return 2;
  if (block.type === 'note') return 4;
  if (block.type === 'summary') return 2;
  if (block.strongHeading || block.type === 'section') return 2;
  if (index <= 2 && title.length >= 20) return 2;
  return 3;
}

function shouldPromoteBlockToHighlight(block, structureMode) {
  if (block.lockAutoTransform) return false;
  if (block.format === 'visual') return !!extractHighlightFromBlock(block);
  
  // Promoção mais conservadora: apenas alta confiança e estrutura clara
  if (structureMode === 'visual' && block.type === 'highlight' && block.confidence === 'high') {
    const highlight = extractHighlightFromBlock(block);
    return !!highlight && highlight.value && highlight.label;
  }
  
  return false;
}

function shouldIncludeAsSection(block, format, structureMode) {
  if (block.type === 'summary' || block.type === 'nextSteps') return false;
  if (format === 'visual' || format === 'table' || format === 'matrix' || format === 'comparison') return false;
  if (format === 'section' || format === 'list' || format === 'note') return true;
  return structureMode === 'editorial' && !!String(block.title || '').trim();
}

function shouldRenderMetaStripInStructuredFlow(cfg) {
  return ['parecer', 'proposta', 'sintese', 'relatorio'].includes(cfg.docType) && !!buildMetaStripHTML(cfg);
}

function renderStructuredBlockText(text, mode) {
  const normalized = normalizeContentText(text);
  const paragraphs = splitTextIntoParagraphs(normalized);
  if (!paragraphs.length) return '';
  if (mode === 'preserve') {
    const lines = normalized.split('\n').filter(Boolean);
    const bulletLines = lines.filter(line => looksLikeBullet(line.trim()));
    if (bulletLines.length >= 2 && bulletLines.length >= lines.length * 0.5) {
      const prose = lines.filter(line => !looksLikeBullet(line.trim()));
      const listItems = lines.filter(line => looksLikeBullet(line.trim()));
      return [
        prose.length ? prose.map(line => `<p class="doc-preserved-block__paragraph">${escHtml(line)}</p>`).join('') : '',
        `<ul class="doc-preserved-block__list">${listItems.map(line => `<li>${escHtml(stripListMarker(line))}</li>`).join('')}</ul>`,
      ].filter(Boolean).join('');
    }
    if (looksLikeKeyValueBlock(normalized)) {
      return buildMatrixBlockHTML(lines);
    }
    return paragraphs.map(paragraph => `<p class="doc-preserved-block__paragraph">${renderInlinePreservedText(paragraph)}</p>`).join('');
  }
  return renderTextParagraphs(paragraphs.join('\n\n'));
}

function renderInlinePreservedText(text) {
  return escHtml(text)
    .replace(/\n/g, '<br>')
    .replace(/&lt;br\s*\/?&gt;/gi, '<br>');
}

function mergeSequentialPreservedBlocks(model) {
  const merged = [];
  for (const item of model || []) {
    const prev = merged[merged.length - 1];
    if (prev && item && prev.kind === 'block' && item.kind === 'block' && prev.mode === 'preserve' && item.mode === 'preserve' && !prev.title && !item.title) {
      prev.text = `${prev.text}\n\n${item.text}`.trim();
      continue;
    }
    merged.push(item);
  }
  return merged;
}

function normalizeDocumentModel(model) {
  return (model || []).map(item => {
    if (!item || !item.kind) return null;
    if (item.kind === 'summary') return { kind: 'summary', text: normalizeContentText(item.text || '') };
    if (item.kind === 'highlights') return { kind: 'highlights', items: dedupeHighlights(item.items || []) };
    if (item.kind === 'nextSteps') return { kind: 'nextSteps', items: normalizeStepLines(normalizeContentText((item.items || []).join('\n') || item.text || '')) };
    if (item.kind === 'block') {
      const normalizedTitle = normalizeContentText(item.title || '');
      const normalizedText = normalizeContentText(item.text || '');
      if (!normalizedTitle && !normalizedText) return null;
      return {
        kind: 'block',
        mode: item.mode || 'preserve',
        title: normalizedTitle,
        text: normalizedText,
        keepHeading: item.keepHeading !== false,
        number: item.number || '',
        highlight: item.highlight || null,
        tableData: item.tableData || null,
        headingLevel: item.headingLevel || 2,
      };
    }
    return item;
  }).filter(Boolean);
}

function dedupeHighlights(items) {
  const seen = new Set();
  return (items || []).filter(item => {
    const key = `${item.label || ''}|${item.value || ''}|${item.desc || ''}`.toLowerCase();
    if (!key.trim() || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function stripListMarker(line) {
  return String(line || '').replace(/^[\-\•\*\d]+[\.\)\-]?\s+/, '').trim();
}

function extractSectionNumber(text) {
  const match = String(text || '').trim().match(/^(\d+(?:\.\d+)?(?:\s*[-–—]\s*|\s+))/);
  return match ? match[1].trim() : '';
}

/* ══════════════════════════════════════════════════════════════
   MÓDULO: EXEMPLO
   ══════════════════════════════════════════════════════════════ */
function loadExample() {
  state.sections = []; state.highlights = [];
  state.sectionIdCounter = 0; state.highlightIdCounter = 0;
  els.sectionsList.innerHTML = ''; els.highlightsList.innerHTML = '';

  els.fldTitle.value    = 'Adoption and Engagement Review';
  els.fldSubtitle.value = 'Sample product operations report · Q1 2025';
  els.fldDate.value     = 'Março 2025';
  els.fldVersion.value  = 'v2.1';
  els.fldAuthor.value   = 'Product Strategy Team';
  els.fldDept.value     = 'Product · Growth · CRM';
  els.fldSummary.value  = `The first quarter of 2025 showed consistent product adoption, stronger retention across repeat users and healthier engagement in the highest-value journeys. This sample report summarizes the main signals, highlights what changed in user behavior and proposes the next operational moves for the team.`;

  addHighlight('Usuários Ativos', '+4.820', 'Crescimento de 34% em relação ao trimestre anterior');
  addHighlight('Eventos Catálogados', '1.247', 'Expansão para 3 novos segmentos culturais');
  addHighlight('Taxa de Retenção', '72%', 'Acima da meta de 65% estabelecida para Q1');
  addHighlight('NPS Médio', '8.4 / 10', 'Satisfação elevada entre usuários recorrentes');

  addSection('Contexto de Mercado',
    `The category expanded steadily during the quarter, driven by better onboarding, stronger referral loops and clearer positioning in the activation funnel.\n\nThis sample product positioned itself well by serving both high-intent returning users and first-time visitors who needed clearer guidance to reach value quickly.`);
  addSection('Análise de Engajamento',
    `User behavior suggests a meaningful shift in recurring usage: people who return to the platform more than three times per month show materially stronger conversion and deeper feature adoption.\n\nThe reward layer inside the product appears to be the main retention driver, showing up repeatedly in qualitative feedback and in the strongest repeat-visit cohorts.`);
  addSection('Parcerias Estratégicas',
    `The team launched 12 new partnerships during the quarter, expanding the catalog by 28% and reducing concentration in the two most saturated content categories.\n\nHighlights include one flagship launch partnership, one public-sector integration pilot and one distribution agreement that broadened the offering for underrepresented segments.`);

  els.fldNextSteps.value = `Launch a personalized recommendation layer by April 2025\nExpand the rewards program with new partner categories\nStart a paid acquisition pilot in two additional regions\nShip an analytics panel for external stakeholders\nConsolidate integrations with the main fulfillment systems`;

  document.querySelectorAll('.doc-type').forEach(b => b.classList.toggle('active', b.dataset.type === 'relatorio'));
  state.docType = 'relatorio';
  setTheme('grafite');
  generateDocument();
}

function runAutotestFromQuery() {
  const params = new URLSearchParams(window.location.search);
  if (!params.has('autotest')) return;

  const docType = params.get('type');
  const theme = params.get('theme');

  loadExample();

  if (docType && DOC_TYPE_LABELS[docType]) {
    state.docType = docType;
    document.querySelectorAll('.doc-type').forEach(b => b.classList.toggle('active', b.dataset.type === docType));
  }

  if (theme && THEME_DESCS[theme]) {
    setTheme(theme);
  }

  generateDocument();
}

function setTheme(theme) {
  state.theme = theme;
  document.querySelectorAll('.theme-card').forEach(b => b.classList.toggle('active', b.dataset.theme === theme));
  if (els.themeDesc) els.themeDesc.innerHTML = THEME_DESCS[theme] || '';
}

function normalizeSections(sections) {
  return (sections || [])
    .map(section => ({
      id: section.id,
      title: normalizeContentText(section.title || ''),
      body: normalizeContentText(section.body || ''),
    }))
    .filter(section => section.title || section.body);
}

function normalizeHighlights(highlights) {
  return (highlights || [])
    .map(item => ({
      id: item.id,
      label: normalizeContentText(item.label || ''),
      value: normalizeContentText(item.value || ''),
      desc: normalizeContentText(item.desc || ''),
    }))
    .filter(item => item.label || item.value || item.desc);
}

function normalizeStepLines(text) {
  return normalizeContentText(text || '')
    .split('\n')
    .map(line => line.replace(/^[\-\•\*\d\.\)\s]+/, '').trim())
    .filter(Boolean);
}

function firstMeaningfulLine(text, maxLength = 120) {
  return String(text || '')
    .split('\n')
    .map(line => line.trim())
    .find(line => line.length > 0)?.slice(0, maxLength) || '';
}

function crmPriorityTag(title, body) {
  const content = `${title || ''} ${body || ''}`.toLowerCase();
  if (/quente|priorit|máx|alto valor|elite|urgente/.test(content)) return { className: 'max', label: 'máxima' };
  if (/morno|alta|qualificado|potencial/.test(content)) return { className: 'high', label: 'alta' };
  if (/m[eé]dio|moderado|base ativa/.test(content)) return { className: 'mid', label: 'média' };
  return { className: 'low', label: 'monitorar' };
}

async function waitForAssetsReady(container) {
  if (!container) return;
  const images = Array.from(container.querySelectorAll('img'));
  if (!images.length) return;

  await Promise.all(images.map(img => {
    if (img.complete) return Promise.resolve();
    return new Promise(resolve => {
      img.addEventListener('load', resolve, { once: true });
      img.addEventListener('error', resolve, { once: true });
    });
  }));
}

function preparePrintMode() {
  document.body.classList.add('is-printing');
  const doc = els.docWrap?.querySelector('.gudi-doc');
  if (doc) doc.dataset.exportMode = 'print';
}

function cleanupPrintMode() {
  document.body.classList.remove('is-printing');
  const doc = els.docWrap?.querySelector('.gudi-doc');
  if (doc) delete doc.dataset.exportMode;
}

/* ══════════════════════════════════════════════════════════════
   UTILS
   ══════════════════════════════════════════════════════════════ */
function escHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function decodeHtmlEntities(str) {
  if (!str) return '';
  const textarea = document.createElement('textarea');
  textarea.innerHTML = String(str);
  return textarea.value;
}

function stripHtmlTags(str) {
  if (!str) return '';
  const div = document.createElement('div');
  div.innerHTML = String(str);
  return div.textContent || div.innerText || '';
}

function normalizeContentText(value) {
  if (value == null) return '';
  const withBreaks = String(value)
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>\s*<p[^>]*>/gi, '\n\n')
    .replace(/<\/li>\s*<li[^>]*>/gi, '\n')
    .replace(/<\/?(p|div|section|article|ul|ol|li|h1|h2|h3|h4|h5|h6|span|strong|em|b|i|small)[^>]*>/gi, match => /^(<\/(p|div|section|article|ul|ol|li|h1|h2|h3|h4|h5|h6))/.test(match) ? '\n' : '');
  const noTags = stripHtmlTags(withBreaks);
  return decodeHtmlEntities(noTags)
    .replace(/\r\n/g, '\n')
    .replace(/\u00a0/g, ' ')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n[ \t]+/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/[ \t]{2,}/g, ' ')
    .trim();
}

function renderTextParagraphs(text, className = '') {
  const paragraphs = splitTextIntoParagraphs(normalizeContentText(text));
  return paragraphs
    .map(paragraph => `<p${className ? ` class="${className}"` : ''}>${escHtml(paragraph).replace(/\n/g, '<br>')}</p>`)
    .join('');
}

function removeHeadingEcho(text, title = '') {
  const cleanText = normalizeContentText(text);
  const cleanTitle = normalizeContentText(title);
  if (!cleanText || !cleanTitle) return cleanText;
  const paragraphs = splitTextIntoParagraphs(cleanText);
  if (!paragraphs.length) return cleanText;
  const first = normalizeHeadingText(paragraphs[0]);
  const heading = normalizeHeadingText(cleanTitle);
  if (first === heading || first.startsWith(heading) || heading.startsWith(first)) {
    return paragraphs.slice(1).join('\n\n').trim();
  }
  return cleanText;
}

function formatDate() {
  const months = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho',
                  'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  const d = new Date();
  return `${months[d.getMonth()]} ${d.getFullYear()}`;
}

/* Boot */
document.addEventListener('DOMContentLoaded', init);
