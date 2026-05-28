// ==UserScript==
// @name         Expansão SP - Relatório de Turmas em Excel
// @namespace    https://expansao.educacao.sp.gov.br/
// @version      1.0.0
// @description  Gera relatório em Excel (.xlsx) das turmas do Data Analytics (Expansão 2026): resumo da turma, alunos e notas por disciplina. Exporta a turma atual ou todas as turmas.
// @author       você
// @match        https://expansao.educacao.sp.gov.br/local/data_analytics/*
// @require      https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js
// @grant        none
// @run-at       document-idle
// ==/UserScript==

(function () {
  'use strict';

  // ---------- Utilitários de parsing ----------

  const txt = (el) => (el ? el.textContent.trim() : '');

  // Converte "8.49" / "3,00" em número; retorna null se não houver número.
  function parseNum(s) {
    if (s == null) return null;
    const m = String(s).replace(/\s/g, '').match(/-?\d+(?:[.,]\d+)?/);
    if (!m) return null;
    const n = parseFloat(m[0].replace(',', '.'));
    return Number.isFinite(n) ? n : null;
  }

  // Extrai o número de "100%" -> 100; null se não houver.
  function parsePct(s) {
    if (s == null) return null;
    const m = String(s).match(/(\d+(?:[.,]\d+)?)\s*%/);
    if (!m) return null;
    const n = parseFloat(m[1].replace(',', '.'));
    return Number.isFinite(n) ? n : null;
  }

  function sanitizeFile(s) {
    return String(s || '')
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // remove acentos
      .replace(/[^\w\- ]+/g, '')
      .trim()
      .replace(/\s+/g, '_')
      .slice(0, 80) || 'relatorio';
  }

  function dataHoje() {
    const d = new Date();
    const p = (n) => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
  }

  function escapeHtml(s) {
    return String(s == null ? '' : s).replace(/[&<>"]/g, (c) => (
      { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]
    ));
  }

  function fmtNota(n) {
    return n == null || isNaN(n) ? '—' : Number(n).toFixed(2).replace('.', ',');
  }
  function fmtProg(p) {
    return p == null || isNaN(p) ? '—' : `${Math.round(p)}%`;
  }
  function classeNota(n) {
    if (n == null || isNaN(n)) return '';
    if (n < 5) return 'nota-ruim';
    if (n < 7) return 'nota-medio';
    return 'nota-bom';
  }

  // ---------- Extração dos dados da turma exibida ----------

  function getTurmaNome() {
    const sel = document.getElementById('select-group');
    if (sel && sel.selectedOptions && sel.selectedOptions.length) {
      const opt = sel.selectedOptions[0];
      if (opt.value) return opt.textContent.trim();
    }
    return '';
  }

  function coletarTurmaAtual() {
    const display = document.getElementById('group-info-display');
    if (!display) return null;

    const legend = txt(document.getElementById('display-acessos-legend'));
    const mAcessaram = legend.match(/(\d+)\s*alunos?\s+acessaram/i);
    const mNunca = legend.match(/(\d+)\s*alunos?\s+nunca/i);

    const turma = {
      nome: getTurmaNome(),
      idSed: txt(document.getElementById('display-id')),
      diretoria: txt(document.getElementById('display-diretoria')),
      escola: txt(document.getElementById('display-escola')),
      cidade: txt(document.getElementById('display-cidade')),
      itinerario: txt(document.getElementById('display-itinerario')),
      mediaNota: parseNum(txt(document.getElementById('display-media-nota'))),
      mediaProgresso: parsePct(txt(document.getElementById('display-media-progresso'))),
      acessaram: mAcessaram ? parseInt(mAcessaram[1], 10) : null,
      nuncaAcessaram: mNunca ? parseInt(mNunca[1], 10) : null,
      alunos: [],
    };

    document.querySelectorAll('#students-table tbody tr.student-main-row').forEach((tr) => {
      const tds = tr.querySelectorAll(':scope > td');
      if (tds.length < 4) return;

      const aluno = {
        nome: txt(tds[0]),
        ultimoAcesso: txt(tds[1]),
        notaMedia: parseNum(txt(tds[2])),
        progresso: parsePct(txt(tds[3])),
        disciplinas: [],
      };

      // Linha de detalhes (notas por disciplina) — pode estar oculta (d-none).
      const btn = tr.querySelector('.toggle-details');
      const targetId = btn ? btn.getAttribute('data-target') : null;
      let detailRow = targetId ? document.getElementById(targetId) : null;
      if (!detailRow && tr.nextElementSibling && tr.nextElementSibling.classList.contains('student-details-row')) {
        detailRow = tr.nextElementSibling;
      }

      if (detailRow) {
        detailRow.querySelectorAll('table tbody tr').forEach((dr) => {
          const dcells = dr.querySelectorAll('td');
          if (dcells.length < 3) return;
          const disc = txt(dcells[0]);
          if (!disc || /MÉDIA\s+GERAL/i.test(disc)) return; // ignora linha de média
          aluno.disciplinas.push({
            disciplina: disc,
            nota: parseNum(txt(dcells[1])),
            progresso: parsePct(txt(dcells[2])),
          });
        });
      }

      turma.alunos.push(aluno);
    });

    turma.totalAlunos = turma.alunos.length;
    return turma;
  }

  // ---------- Espera a turma carregar (para exportar todas) ----------

  function waitForTurma(valorEsperado, timeoutMs = 10000) {
    return new Promise((resolve) => {
      const inicio = Date.now();
      const ok = () => {
        const id = txt(document.getElementById('display-id'));
        const temAlunos = document.querySelectorAll('#students-table tbody tr.student-main-row').length > 0;
        // O #display-id costuma ser igual ao value da opção (ID SED).
        return temAlunos && (!valorEsperado || id === String(valorEsperado));
      };
      if (ok()) return resolve(true);
      const timer = setInterval(() => {
        if (ok()) {
          clearInterval(timer);
          // pequeno respiro para o DOM terminar de renderizar
          setTimeout(() => resolve(true), 150);
        } else if (Date.now() - inicio > timeoutMs) {
          clearInterval(timer);
          resolve(false);
        }
      }, 200);
    });
  }

  // ---------- Estilos do Excel ----------

  const BRAND = '0A1970';        // azul institucional (cabeçalho)
  const ZEBRA = 'F4F6FB';        // listras das linhas pares
  const CINZA_BORDA = 'D5DAE5';
  const COR = { ruim: 'C0392B', medio: 'B9770E', bom: '1E8449', neutro: '7F8C8D' };

  const borda = { style: 'thin', color: { rgb: CINZA_BORDA } };
  const bordasTodas = { top: borda, bottom: borda, left: borda, right: borda };

  function corNota(n) {
    if (n == null || isNaN(n)) return COR.neutro;
    if (n < 5) return COR.ruim;
    if (n < 7) return COR.medio;
    return COR.bom;
  }
  function corProg(p) {
    if (p == null || isNaN(p)) return COR.neutro;
    if (p < 50) return COR.ruim;
    if (p < 80) return COR.medio;
    return COR.bom;
  }

  // Monta uma aba estilizada: título mesclado, cabeçalho colorido, zebra,
  // bordas, autofiltro, painel congelado e cores condicionais por coluna.
  // colDefs[c] = { w, align, numFmt, cor: 'nota' | 'prog' | null }
  function montarAba(wb, nomeAba, titulo, headers, linhas, colDefs) {
    const nCols = headers.length;
    const ultCol = nCols - 1;
    const ws = XLSX.utils.aoa_to_sheet([[titulo], headers, ...linhas]);

    ws['!cols'] = colDefs.map((c) => ({ wch: c.w }));
    ws['!rows'] = [{ hpt: 26 }, { hpt: 22 }];
    ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: ultCol } }];

    // Título
    const tAddr = XLSX.utils.encode_cell({ r: 0, c: 0 });
    ws[tAddr].s = {
      font: { bold: true, sz: 14, color: { rgb: BRAND } },
      alignment: { vertical: 'center', horizontal: 'left' },
    };

    // Cabeçalho
    for (let c = 0; c < nCols; c++) {
      const addr = XLSX.utils.encode_cell({ r: 1, c });
      if (!ws[addr]) ws[addr] = { t: 's', v: headers[c] };
      ws[addr].s = {
        font: { bold: true, sz: 11, color: { rgb: 'FFFFFF' } },
        fill: { fgColor: { rgb: BRAND } },
        alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
        border: {
          top: { style: 'thin', color: { rgb: BRAND } },
          bottom: { style: 'thin', color: { rgb: BRAND } },
          left: { style: 'thin', color: { rgb: 'FFFFFF' } },
          right: { style: 'thin', color: { rgb: 'FFFFFF' } },
        },
      };
    }

    // Linhas de dados
    for (let i = 0; i < linhas.length; i++) {
      const r = 2 + i;
      const ehZebra = i % 2 === 1;
      for (let c = 0; c < nCols; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr];
        if (!cell) continue;
        const def = colDefs[c] || {};
        const estilo = {
          alignment: { horizontal: def.align || 'left', vertical: 'center' },
          border: bordasTodas,
        };
        if (ehZebra) estilo.fill = { fgColor: { rgb: ZEBRA } };
        if (def.numFmt) {
          cell.z = def.numFmt;
          estilo.numFmt = def.numFmt;
        }
        if (def.cor === 'nota') estilo.font = { bold: true, color: { rgb: corNota(cell.v) } };
        else if (def.cor === 'prog') estilo.font = { color: { rgb: corProg(cell.v) } };
        cell.s = estilo;
      }
    }

    // Autofiltro na linha de cabeçalho
    ws['!autofilter'] = {
      ref: `${XLSX.utils.encode_cell({ r: 1, c: 0 })}:${XLSX.utils.encode_cell({ r: 1 + linhas.length, c: ultCol })}`,
    };

    XLSX.utils.book_append_sheet(wb, ws, nomeAba);
    return ws;
  }

  // ---------- Geração do arquivo Excel ----------

  function gerarExcel(turmas, nomeArquivo) {
    if (typeof XLSX === 'undefined') {
      alert('A biblioteca de Excel (SheetJS) não carregou. Verifique sua conexão e recarregue a página.');
      return;
    }
    const turmasValidas = turmas.filter((t) => t && (t.alunos.length || t.idSed));
    if (!turmasValidas.length) {
      alert('Nenhum dado de turma encontrado para exportar.');
      return;
    }

    const wb = XLSX.utils.book_new();
    const escola = turmasValidas[0].escola || 'Escola';
    const dataBR = new Date().toLocaleDateString('pt-BR');
    const subtitulo =
      turmasValidas.length === 1
        ? `${turmasValidas[0].nome || ''}`
        : `${turmasValidas.length} turmas`;

    // Aba 1: visão geral das turmas
    montarAba(
      wb,
      'Turmas',
      `Relatório de Turmas — ${escola}  •  ${dataBR}`,
      ['Turma', 'ID SED', 'Diretoria', 'Escola', 'Cidade', 'Itinerário',
        'Nota média', 'Progresso médio', 'Acessaram', 'Nunca acessaram', 'Total de alunos'],
      turmasValidas.map((t) => [
        t.nome, t.idSed, t.diretoria, t.escola, t.cidade, t.itinerario,
        t.mediaNota, t.mediaProgresso, t.acessaram, t.nuncaAcessaram, t.totalAlunos,
      ]),
      [
        { w: 28, align: 'left' }, { w: 12, align: 'center' }, { w: 16, align: 'left' },
        { w: 36, align: 'left' }, { w: 16, align: 'left' }, { w: 14, align: 'center' },
        { w: 12, align: 'center', numFmt: '0.00', cor: 'nota' },
        { w: 16, align: 'center', numFmt: '0"%"', cor: 'prog' },
        { w: 12, align: 'center' }, { w: 16, align: 'center' }, { w: 15, align: 'center' },
      ]
    );

    // Aba 2: alunos (resumo)
    const linhasAlunos = [];
    turmasValidas.forEach((t) => {
      t.alunos.forEach((a) => {
        linhasAlunos.push([t.nome, a.nome, a.ultimoAcesso, a.notaMedia, a.progresso]);
      });
    });
    montarAba(
      wb,
      'Alunos',
      `Alunos — ${escola}  •  ${subtitulo}`,
      ['Turma', 'Nome', 'Último acesso', 'Nota média', 'Progresso'],
      linhasAlunos,
      [
        { w: 28, align: 'left' }, { w: 38, align: 'left' }, { w: 16, align: 'center' },
        { w: 12, align: 'center', numFmt: '0.00', cor: 'nota' },
        { w: 14, align: 'center', numFmt: '0"%"', cor: 'prog' },
      ]
    );

    // Aba 3: notas por disciplina (formato longo, fácil de filtrar/dinamizar)
    const linhasNotas = [];
    turmasValidas.forEach((t) => {
      t.alunos.forEach((a) => {
        a.disciplinas.forEach((d) => {
          linhasNotas.push([t.nome, a.nome, d.disciplina, d.nota, d.progresso]);
        });
      });
    });
    montarAba(
      wb,
      'Notas por Disciplina',
      `Notas por disciplina — ${escola}  •  ${subtitulo}`,
      ['Turma', 'Nome', 'Disciplina', 'Nota', 'Progresso'],
      linhasNotas,
      [
        { w: 28, align: 'left' }, { w: 38, align: 'left' }, { w: 32, align: 'left' },
        { w: 10, align: 'center', numFmt: '0.00', cor: 'nota' },
        { w: 14, align: 'center', numFmt: '0"%"', cor: 'prog' },
      ]
    );

    XLSX.writeFile(wb, nomeArquivo);
  }

  // ---------- Lista de ciência dos responsáveis (para impressão) ----------

  function construirHtmlCiencia(turma) {
    const dataBR = new Date().toLocaleDateString('pt-BR');

    const meta = `
      <p class="meta">
        <b>Diretoria de Ensino:</b> ${escapeHtml(turma.diretoria)} &nbsp;•&nbsp;
        <b>Escola:</b> ${escapeHtml(turma.escola)}<br>
        <b>Turma:</b> ${escapeHtml(turma.nome)} &nbsp;•&nbsp;
        <b>Itinerário:</b> ${escapeHtml(turma.itinerario)} &nbsp;•&nbsp;
        <b>Cidade:</b> ${escapeHtml(turma.cidade)}<br>
        <b>Data de emissão:</b> ${dataBR} &nbsp;•&nbsp;
        <b>Bimestre/Período:</b> ____________________
      </p>`;

    // 1) Lista consolidada
    const linhasConsolidado = turma.alunos.map((a, i) => `
      <tr>
        <td class="c">${i + 1}</td>
        <td>${escapeHtml(a.nome)}</td>
        <td class="c ${classeNota(a.notaMedia)}">${fmtNota(a.notaMedia)}</td>
        <td class="c">${fmtProg(a.progresso)}</td>
        <td class="assinatura-cel"></td>
        <td class="assinatura-cel"></td>
      </tr>`).join('');

    const consolidado = `
      <h1>Ciência dos Responsáveis — Resultados dos Estudantes</h1>
      ${meta}
      <p class="declaracao">
        Declaramos, pela assinatura abaixo, estar cientes do desempenho escolar
        do(a) estudante sob nossa responsabilidade, referente ao período letivo informado.
      </p>
      <table>
        <thead>
          <tr>
            <th class="c" style="width:34px">#</th>
            <th>Estudante</th>
            <th class="c" style="width:70px">Nota média</th>
            <th class="c" style="width:80px">Progresso</th>
            <th style="width:32%">Nome do responsável</th>
            <th style="width:32%">Assinatura</th>
          </tr>
        </thead>
        <tbody>${linhasConsolidado}</tbody>
      </table>`;

    // 2) Comprovantes individuais
    const comprovantes = turma.alunos.map((a) => {
      const disc = a.disciplinas.map((d) => `
        <tr>
          <td>${escapeHtml(d.disciplina)}</td>
          <td class="c ${classeNota(d.nota)}">${fmtNota(d.nota)}</td>
          <td class="c">${fmtProg(d.progresso)}</td>
        </tr>`).join('');

      return `
        <div class="slip">
          <h3>${escapeHtml(a.nome)}</h3>
          <div class="slip-meta">
            ${escapeHtml(turma.escola)} • Turma ${escapeHtml(turma.nome)} • ${escapeHtml(turma.itinerario)}
          </div>
          <table>
            <thead>
              <tr><th>Disciplina</th><th class="c" style="width:80px">Nota</th><th class="c" style="width:90px">Progresso</th></tr>
            </thead>
            <tbody>
              ${disc}
              <tr class="linha-total">
                <td><b>Média geral / Progresso</b></td>
                <td class="c ${classeNota(a.notaMedia)}"><b>${fmtNota(a.notaMedia)}</b></td>
                <td class="c"><b>${fmtProg(a.progresso)}</b></td>
              </tr>
            </tbody>
          </table>
          <p class="declaracao">
            Declaro estar ciente do desempenho escolar acima, referente ao(à) estudante sob minha responsabilidade.
          </p>
          <div class="assina-area">
            <div class="campo"><div class="linha-assina"></div><div class="rotulo">Nome do responsável</div></div>
            <div class="campo" style="max-width:160px"><div class="linha-assina"></div><div class="rotulo">Data</div></div>
            <div class="campo"><div class="linha-assina"></div><div class="rotulo">Assinatura</div></div>
          </div>
        </div>`;
    }).join('');

    const css = `
      *{box-sizing:border-box}
      body{font-family:Arial,Helvetica,sans-serif;color:#222;margin:24px;font-size:12px}
      .toolbar{position:sticky;top:0;background:#0A1970;padding:10px 14px;margin:-24px -24px 18px;border-radius:0 0 8px 8px;display:flex;align-items:center;gap:12px}
      .toolbar button{background:#fff;color:#0A1970;border:0;padding:8px 14px;font-weight:bold;border-radius:6px;cursor:pointer;font-size:13px}
      .toolbar span{color:#cdd6ff;font-size:12px}
      h1{font-size:18px;color:#0A1970;margin:0 0 6px}
      h2{font-size:15px;color:#0A1970;margin:18px 0 10px}
      .meta{margin:0 0 10px;line-height:1.6}
      .declaracao{background:#f4f6fb;border:1px solid #d5dae5;padding:10px 12px;border-radius:6px;margin:10px 0 14px}
      table{width:100%;border-collapse:collapse;margin-bottom:12px}
      th,td{border:1px solid #b9c0d4;padding:6px 8px;vertical-align:middle}
      th{background:#0A1970;color:#fff;font-size:12px;text-align:left}
      td.c,th.c{text-align:center}
      .assinatura-cel{height:34px}
      .nota-ruim{color:#C0392B;font-weight:bold}
      .nota-medio{color:#B9770E;font-weight:bold}
      .nota-bom{color:#1E8449;font-weight:bold}
      .linha-total td{background:#f4f6fb}
      .slip{border:1px solid #0A1970;border-radius:8px;padding:14px 16px;margin-bottom:14px;page-break-inside:avoid}
      .slip h3{margin:0 0 4px;font-size:14px;color:#0A1970}
      .slip-meta{color:#555;font-size:11px;margin-bottom:8px}
      .assina-area{display:flex;gap:24px;margin-top:14px;flex-wrap:wrap}
      .campo{flex:1;min-width:200px}
      .linha-assina{border-bottom:1px solid #333;height:28px}
      .rotulo{font-size:11px;color:#555;margin-top:3px}
      .quebra{page-break-before:always}
      @media print{.no-print{display:none!important}body{margin:0}@page{size:A4;margin:14mm}}`;

    return `<!doctype html>
<html lang="pt-br"><head><meta charset="utf-8">
<title>Ciência dos Responsáveis — ${escapeHtml(turma.nome)}</title>
<style>${css}</style></head>
<body>
  <div class="toolbar no-print">
    <button onclick="window.print()">Imprimir / Salvar como PDF</button>
    <span>${turma.alunos.length} estudante(s) • a lista consolidada e os comprovantes individuais saem em páginas separadas.</span>
  </div>
  ${consolidado}
  <div class="quebra"></div>
  <h2>Comprovantes individuais (um por responsável)</h2>
  ${comprovantes}
</body></html>`;
  }

  function gerarListaCiencia(turma) {
    if (!turma || !turma.alunos.length) {
      alert('Selecione uma turma com alunos antes de gerar a lista.');
      return;
    }
    const html = construirHtmlCiencia(turma);
    const win = window.open('', '_blank');
    if (!win) {
      alert('Não foi possível abrir a janela de impressão. Permita pop-ups para este site e tente novamente.');
      return;
    }
    win.document.open();
    win.document.write(html);
    win.document.close();
  }

  // ---------- Ações dos botões ----------

  function exportarTurmaAtual() {
    const turma = coletarTurmaAtual();
    if (!turma || !turma.alunos.length) {
      alert('Selecione uma turma com alunos antes de exportar.');
      return;
    }
    const nome = `Relatorio_Turma_${sanitizeFile(turma.nome || turma.idSed)}_${dataHoje()}.xlsx`;
    gerarExcel([turma], nome);
  }

  function exportarListaCiencia() {
    const turma = coletarTurmaAtual();
    if (!turma || !turma.alunos.length) {
      alert('Selecione uma turma com alunos antes de gerar a lista.');
      return;
    }
    gerarListaCiencia(turma);
  }

  async function exportarTodasTurmas(botao) {
    const sel = document.getElementById('select-group');
    if (!sel) {
      alert('Seletor de turmas não encontrado.');
      return;
    }
    const opcoes = Array.from(sel.options).filter((o) => o.value);
    if (!opcoes.length) {
      alert('Nenhuma turma disponível no seletor.');
      return;
    }

    const valorOriginal = sel.value;
    const rotuloOriginal = botao.innerHTML;
    botao.classList.add('disabled');

    const turmas = [];
    const falhas = [];
    try {
      for (let i = 0; i < opcoes.length; i++) {
        const opt = opcoes[i];
        botao.innerHTML = `<i class="fa fa-spinner fa-spin"></i> Coletando ${i + 1}/${opcoes.length}…`;
        sel.value = opt.value;
        sel.dispatchEvent(new Event('change', { bubbles: true }));
        const carregou = await waitForTurma(opt.value);
        if (!carregou) {
          falhas.push(opt.textContent.trim());
          continue;
        }
        const turma = coletarTurmaAtual();
        if (turma) {
          if (!turma.nome) turma.nome = opt.textContent.trim();
          turmas.push(turma);
        }
      }
    } finally {
      // Restaura a turma originalmente selecionada
      if (valorOriginal) {
        sel.value = valorOriginal;
        sel.dispatchEvent(new Event('change', { bubbles: true }));
      }
      botao.innerHTML = rotuloOriginal;
      botao.classList.remove('disabled');
    }

    if (!turmas.length) {
      alert('Não foi possível coletar dados das turmas. Tente exportar a turma atual.');
      return;
    }

    const escola = turmas[0].escola || 'escola';
    const nome = `Relatorio_Turmas_${sanitizeFile(escola)}_${dataHoje()}.xlsx`;
    gerarExcel(turmas, nome);

    if (falhas.length) {
      alert(`Exportado com ${turmas.length} turma(s).\nNão carregaram a tempo: ${falhas.join(', ')}`);
    }
  }

  // ---------- Injeção dos botões na página ----------

  function criarBotao(id, texto, onClick) {
    const a = document.createElement('a');
    a.id = id;
    a.href = '#';
    a.className = 'btn btn-outline-primary ms-2';
    a.innerHTML = texto;
    a.addEventListener('click', (e) => {
      e.preventDefault();
      onClick(a);
    });
    return a;
  }

  function injetarBotoes() {
    if (document.getElementById('tm-export-turma-atual')) return; // já injetado
    const refBtn = document.getElementById('export-group-csv');
    const container = refBtn
      ? refBtn.parentElement
      : document.querySelector('.local_data_analytics_container_index .row');
    if (!container) return;

    const btnAtual = criarBotao(
      'tm-export-turma-atual',
      '<i class="fa fa-file-excel-o"></i> Excel (turma atual)',
      exportarTurmaAtual
    );
    const btnTodas = criarBotao(
      'tm-export-todas-turmas',
      '<i class="fa fa-file-excel-o"></i> Excel (todas as turmas)',
      exportarTodasTurmas
    );
    const btnCiencia = criarBotao(
      'tm-lista-ciencia',
      '<i class="fa fa-print"></i> Lista p/ responsáveis assinarem',
      exportarListaCiencia
    );
    btnCiencia.className = 'btn btn-outline-secondary ms-2';

    container.appendChild(btnAtual);
    container.appendChild(btnTodas);
    container.appendChild(btnCiencia);
  }

  // A página é dinâmica; garante que os botões permaneçam presentes.
  injetarBotoes();
  const obs = new MutationObserver(() => injetarBotoes());
  obs.observe(document.body, { childList: true, subtree: true });
})();
