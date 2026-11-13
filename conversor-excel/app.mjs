import { ajustarPlanilha } from './pgcc.mjs';

const arquivo = document.getElementById('file');
const area = document.getElementById('drop-area');
const resultado = document.getElementById('result');
const nome = document.getElementById('filename');
const estado = document.getElementById('status');
const erro = document.getElementById('error');
const detalhes = document.getElementById('details');
const baixar = document.getElementById('download');
const numero = new Intl.NumberFormat('pt-BR');
let selecao = 0;
let url;

function limparDownload() {
  if (url) URL.revokeObjectURL(url);
  url = undefined;
  baixar.removeAttribute('href');
  baixar.hidden = true;
}

function adicionarDetalhe(texto) {
  const item = document.createElement('li');
  item.textContent = texto;
  detalhes.append(item);
  detalhes.hidden = false;
}

async function processar(file) {
  if (!file) return;
  const atual = ++selecao;
  limparDownload();
  detalhes.replaceChildren();
  detalhes.hidden = true;
  erro.hidden = true;
  erro.textContent = '';
  resultado.hidden = false;
  nome.textContent = file.name;
  estado.textContent = 'Ajustando a planilha…';

  try {
    if (!/\.xlsx$/i.test(file.name)) throw new Error('Selecione um arquivo Excel no formato .xlsx.');
    await new Promise((resolver) => setTimeout(resolver, 0));
    if (atual !== selecao) return;
    const dados = await ajustarPlanilha(await file.arrayBuffer());
    if (atual !== selecao) return;
    url = URL.createObjectURL(dados.blob);
    baixar.href = url;
    baixar.download = file.name.replace(/\.xlsx$/i, '') + '_AJUSTADO.xlsx';
    const celulas = dados.ajustadas === 1 ? 'célula ajustada' : 'células ajustadas';
    const registros = dados.registros === 1 ? 'registro conferido' : 'registros conferidos';
    estado.textContent = dados.ajustadas
      ? numero.format(dados.ajustadas) + ' ' + celulas + ' em ' + numero.format(dados.registros) + ' ' + registros + '.'
      : 'Nenhum ajuste necessário nos textos conferidos.';
    if (dados.semDescricao === 1) adicionarDetalhe('1 conta analítica está sem descrição (descCta).');
    if (dados.semDescricao > 1) adicionarDetalhe(numero.format(dados.semDescricao) + ' contas analíticas estão sem descrição (descCta).');
    if (dados.formulas) adicionarDetalhe('Os campos de texto com fórmulas foram preservados.');
    baixar.hidden = false;
  } catch (falha) {
    if (atual !== selecao) return;
    estado.textContent = '';
    erro.hidden = false;
    erro.textContent = falha instanceof Error ? falha.message : 'Falha ao ler a planilha.';
  }
}

arquivo.addEventListener('change', () => {
  const file = arquivo.files[0];
  arquivo.value = '';
  processar(file);
});
area.addEventListener('dragover', (evento) => {
  evento.preventDefault();
  area.classList.add('is-dragging');
});
area.addEventListener('dragleave', (evento) => {
  if (!area.contains(evento.relatedTarget)) area.classList.remove('is-dragging');
});
area.addEventListener('drop', (evento) => {
  evento.preventDefault();
  area.classList.remove('is-dragging');
  processar(evento.dataTransfer.files[0]);
});
