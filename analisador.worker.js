/* Web Worker: lê as planilhas com Python (Pyodide) e devolve JSON para a página. */
const VERSAO = "v3-20261112";
const PYODIDE_VERSAO = "0.29.5";
const PYODIDE_URL = `https://cdn.jsdelivr.net/pyodide/v${PYODIDE_VERSAO}/full/`;
const ARQUIVOS_PY = ["pacote/__init__.py", "pacote/leitura.py", "pacote/regras.json"];
importScripts(PYODIDE_URL + "pyodide.js");

let pronto = null;
async function prepara() {
  const py = await loadPyodide({ indexURL: PYODIDE_URL });
  self.postMessage({ tipo: "status", texto: "Carregando pandas…" });
  await py.loadPackage(["pandas", "micropip", "xlrd"]);
  await py.pyimport("micropip").install("openpyxl");
  for (const caminho of ARQUIVOS_PY) {
    const resp = await fetch(caminho + "?v=" + VERSAO);
    if (!resp.ok) throw new Error(`Arquivo do painel não encontrado no site: ${caminho} (HTTP ${resp.status})`);
    const txt = await resp.text();
    py.FS.mkdirTree(caminho.split("/").slice(0, -1).join("/") || ".");
    py.FS.writeFile(caminho, txt);
  }
  py.runPython("import sys; sys.path.insert(0, '.'); from pacote.leitura import le_json");
  self.postMessage({ tipo: "status", texto: "Python pronto" });
  return py;
}

self.onmessage = async (e) => {
  const msg = e.data;
  try {
    if (!pronto) pronto = prepara();
    const py = await pronto;
    if (msg.fase !== "analisar") return;
    py.FS.mkdirTree("/dados");
    const leJson = py.globals.get("le_json");
    let n = 0;
    for (const a of msg.arquivos) {
      n++;
      self.postMessage({ tipo: "status", texto: `Lendo ${n}/${msg.arquivos.length}: ${a.nome}` });
      const ext = a.nome.split(".").pop().toLowerCase();
      const caminho = `/dados/arq${n}.${ext}`;
      py.FS.writeFile(caminho, new Uint8Array(await a.arquivo.arrayBuffer()));
      const dados = JSON.parse(leJson(caminho, a.nome));
      py.FS.unlink(caminho);
      self.postMessage({ tipo: "arquivo", chave: a.chave, dados });
    }
    leJson.destroy();
    self.postMessage({ tipo: "fim" });
  } catch (err) {
    self.postMessage({ tipo: "erro", texto: String(err && err.message || err) });
  }
};
