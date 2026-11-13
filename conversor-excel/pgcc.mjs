const entidades = /&(amp|lt|gt|quot|apos);|&#(x[\da-f]+|\d+);/gi;
const entidadesNomeadas = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
const windows = new TextDecoder('windows-1252');
const utf8 = new TextDecoder('utf-8', { fatal: true });
const bytesWindows = new Map(Array.from({ length: 32 }, (_, i) => {
  const byte = i + 128;
  return [windows.decode(Uint8Array.of(byte)), byte];
}));
const bytes = (c) => c.charCodeAt(0) < 256 ? c.charCodeAt(0) : (bytesWindows.get(c) ?? -1);

function restaurarCodificacao(texto) {
  return texto.replace(/[ÃÂâïðÅÄ][\u0080-\u00BF\u0152\u0153\u0160\u0161\u0178\u017D\u017E\u0192\u02C6\u02DC\u2013-\u2022\u2026\u2030\u2039\u203A\u20AC\u2122]{1,3}/gu, (trecho) => {
    const tamanho = trecho[0] === 'ð' ? 4 : 'âï'.includes(trecho[0]) ? 3 : 2;
    if (trecho.length < tamanho) return trecho;
    const sequencia = Array.from(trecho.slice(0, tamanho), bytes);
    if (sequencia.some((byte, i) => i && (byte < 128 || byte > 191))) return trecho;
    try {
      const restaurado = utf8.decode(Uint8Array.from(sequencia));
      return /\p{C}/u.test(restaurado) && restaurado !== '\u00ad'
        ? trecho : restaurado + trecho.slice(tamanho);
    } catch {
      return trecho;
    }
  });
}

export function ajustarTexto(texto) {
  const interpretado = restaurarCodificacao(texto).replace(entidades, (_, nome, valor) => {
    if (nome) return entidadesNomeadas[nome.toLowerCase()];
    const numero = valor[0].toLowerCase() === 'x'
      ? Number.parseInt(valor.slice(1), 16)
      : Number(valor);
    if (numero <= 0x10ffff) {
      return String.fromCodePoint(numero);
    }
    return ' ';
  });
  return interpretado
    .split(/(\(\s*-\s*\))/)
    .map((parte, indice) => indice % 2 === 1 ? '( - )' : parte.normalize('NFD')
      .replace(/\p{M}/gu, '')
      .replace(/[ºª°]/g, '')
      .replace(/&/g, ' E ')
      .replace(/[^A-Za-z0-9 ]/g, ' '))
    .join('')
    .replace(/ +/g, ' ')
    .trim();
}

const filhos = (elemento, nome) => [...elemento.children].filter(
  (filho) => filho.localName === nome && filho.namespaceURI === elemento.namespaceURI,
);
const filho = (elemento, nome) => filhos(elemento, nome)[0];
const serializar = (xml) => new XMLSerializer().serializeToString(xml);

async function lerXml(zip, caminho) {
  const entrada = zip.file(caminho);
  if (!entrada) throw new Error('O Excel está incompleto ou corrompido.');
  const texto = await entrada.async('string');
  if (/<!DOCTYPE/i.test(texto)) throw new Error('Esta estrutura de Excel não é suportada.');
  const xml = new DOMParser().parseFromString(texto, 'application/xml');
  if (xml.getElementsByTagNameNS('*', 'parsererror').length) {
    throw new Error('O Excel contém uma estrutura XML inválida.');
  }
  return xml;
}

function resolver(caminho, alvo) {
  const base = alvo.startsWith('/') ? [] : caminho.split('/').slice(0, -1);
  for (const parte of alvo.split('/')) {
    if (!parte || parte === '.') continue;
    if (parte === '..') base.pop();
    else base.push(parte);
  }
  return base.join('/');
}

function textos(item) {
  return [...item.children].flatMap((elemento) => {
    if (elemento.namespaceURI !== item.namespaceURI) return [];
    if (elemento.localName === 't') return [elemento];
    if (elemento.localName === 'r') return filhos(elemento, 't');
    return [];
  });
}

function obterTexto(celula, itens) {
  const tipo = celula.getAttribute('t');
  const valor = filho(celula, 'v');
  if (tipo === 's' && valor) {
    const item = itens[Number(valor.textContent)];
    if (!item) throw new Error('O Excel contém uma referência de texto inválida.');
    return textos(item).map((texto) => texto.textContent).join('');
  }
  const inline = filho(celula, 'is');
  if (tipo === 'inlineStr' && inline) return textos(inline).map((texto) => texto.textContent).join('');
  if (tipo === 'str' && valor) return valor.textContent;
  return null;
}

function escreverTexto(item, ajustado) {
  const partes = textos(item);
  let posicao = 0;
  for (const [indice, parte] of partes.entries()) {
    const fim = indice === partes.length - 1
      ? ajustado.length
      : posicao + ajustarTexto(parte.textContent).length;
    parte.textContent = ajustado.slice(posicao, fim);
    parte.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
    posicao = fim;
  }
}

function coluna(celula) {
  return celula.getAttribute('r')?.match(/^[A-Z]+/)?.[0];
}

function localizarCampos(linhas, itens) {
  for (const [indice, linha] of linhas.slice(0, 30).entries()) {
    const campos = new Map(filhos(linha, 'c').map((c) => [obterTexto(c, itens)?.trim().toLowerCase(), coluna(c)]));
    if (campos.get('nomecta') && (campos.get('desccta') || campos.get('indcta') || campos.get('ccta'))) {
      return { campos, indice };
    }
  }
  return null;
}

export async function ajustarPlanilha(conteudo) {
  let zip;
  try {
    zip = await globalThis.JSZip.loadAsync(conteudo);
  } catch {
    throw new Error('Não foi possível abrir este Excel. Se estiver protegido por senha, salve uma cópia sem senha.');
  }
  const pacote = (await lerXml(zip, '_rels/.rels')).documentElement;
  const raiz = filhos(pacote, 'Relationship').find((r) => r.getAttribute('Type')?.endsWith('/officeDocument'));
  if (!raiz) throw new Error('O arquivo selecionado não contém uma pasta de trabalho Excel.');

  const caminhoWorkbook = resolver('', raiz.getAttribute('Target'));
  const [nomeWorkbook, ...diretorios] = caminhoWorkbook.split('/').reverse();
  const caminhoRelacoes = [...diretorios.reverse(), '_rels', nomeWorkbook + '.rels'].join('/');
  const workbook = (await lerXml(zip, caminhoWorkbook)).documentElement;
  const relacoes = (await lerXml(zip, caminhoRelacoes)).documentElement;
  const relacaoInterna = (r) => r.getAttribute('TargetMode') !== 'External';
  const recursos = filhos(relacoes, 'Relationship').filter(relacaoInterna);
  const compartilhadas = recursos.find((r) => r.getAttribute('Type')?.endsWith('/sharedStrings'));
  const caminhoStrings = compartilhadas && resolver(caminhoWorkbook, compartilhadas.getAttribute('Target'));
  const strings = caminhoStrings ? await lerXml(zip, caminhoStrings) : null;
  const itens = strings ? filhos(strings.documentElement, 'si') : [];
  const novosIndices = new Map();
  const guias = filho(workbook, 'sheets');
  let guiasPGCC = 0;
  let registros = 0;
  let ajustadas = 0;
  let semDescricao = 0;
  let formulas = 0;

  for (const guia of guias ? filhos(guias, 'sheet') : []) {
    const id = [...guia.attributes].find((a) => a.localName === 'id' && a.namespaceURI)?.value;
    const relacao = recursos.find((r) => r.getAttribute('Id') === id && r.getAttribute('Type')?.endsWith('/worksheet'));
    if (!relacao) continue;
    const caminho = resolver(caminhoWorkbook, relacao.getAttribute('Target'));
    const planilha = await lerXml(zip, caminho);
    const dados = filho(planilha.documentElement, 'sheetData');
    if (!dados) continue;
    const linhas = filhos(dados, 'row');
    const cabecalho = localizarCampos(linhas, itens);
    if (!cabecalho) continue;
    guiasPGCC++;
    const { campos, indice } = cabecalho;
    const alvo = [campos.get('nomecta'), campos.get('desccta')].filter(Boolean);
    let alterada = false;

    for (const linha of linhas.slice(indice + 1)) {
      const celulas = new Map(filhos(linha, 'c').map((c) => [coluna(c), c]));
      const nome = celulas.get(campos.get('nomecta'));
      const codigo = celulas.get(campos.get('ccta'));
      if ((nome && (filho(nome, 'v') || filho(nome, 'is') || filho(nome, 'f'))) ||
          (codigo && (filho(codigo, 'v') || filho(codigo, 'is')))) registros++;
      const tipo = celulas.get(campos.get('indcta'));
      const descricao = celulas.get(campos.get('desccta'));
      const analitica = tipo && obterTexto(tipo, itens)?.trim().toUpperCase() === 'A';
      for (const letra of alvo) {
        const celula = celulas.get(letra);
        if (!celula) continue;
        if (filho(celula, 'f')) {
          formulas++;
          continue;
        }
        const original = obterTexto(celula, itens);
        if (original === null) continue;
        const ajustado = ajustarTexto(original);
        if (ajustado === original) continue;
        const valor = filho(celula, 'v');
        switch (celula.getAttribute('t')) {
          case 's': {
            const anterior = Number(valor.textContent);
            if (!novosIndices.has(anterior)) {
              const item = itens[anterior].cloneNode(true);
              escreverTexto(item, ajustado);
              const sst = strings.documentElement;
              sst.insertBefore(item, filho(sst, 'extLst') ?? null);
              novosIndices.set(anterior, itens.length);
              itens.push(item);
            }
            valor.textContent = novosIndices.get(anterior);
            break;
          }
          case 'inlineStr':
            escreverTexto(filho(celula, 'is'), ajustado);
            break;
          default:
            valor.textContent = ajustado;
        }
        ajustadas++;
        alterada = true;
      }
      if (analitica && (!descricao || !obterTexto(descricao, itens)?.trim())) semDescricao++;
    }
    if (alterada) zip.file(caminho, serializar(planilha));
  }
  if (!guiasPGCC) throw new Error('Não encontrei uma aba com os campos do PGCC. Selecione o Excel exportado para a DeRE.');
  if (novosIndices.size) {
    strings.documentElement.setAttribute('uniqueCount', String(itens.length));
    zip.file(caminhoStrings, serializar(strings));
  }

  const blob = ajustadas
    ? await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
    : new Blob([conteudo], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  return { blob, registros, ajustadas, semDescricao, formulas };
}
