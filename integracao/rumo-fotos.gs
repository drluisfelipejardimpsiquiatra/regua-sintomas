/**
 * Rumo Livre · fotos dos testes
 * -----------------------------
 * Recebe a foto da fita reagente enviada pela página do paciente (rumo.html)
 * e guarda no Google Drive do médico, na pasta "Rumo Livre · fotos dos testes".
 * Nome do arquivo: RL_<código do paciente>_<data do teste>_<hora do envio>.jpg
 * O arquivo não tem nome nem CPF do paciente: só o código pessoal do link.
 * O painel Rumo Livre lê a pasta pelo conector do Google Drive.
 *
 * Instalação: veja integracao/COMO-INSTALAR-RUMO-FOTOS.md
 */

const PASTA_NOME = "Rumo Livre · fotos dos testes";
const MAX_BASE64 = 4 * 1024 * 1024; // ~3 MB de imagem

/** Rodar UMA vez pelo menu Executar: cria a pasta e guarda o id. */
function instalar() {
  const pasta = pasta_();
  Logger.log("Instalado. Pasta: " + pasta.getName() + " (" + pasta.getId() + "). Agora publique como app da Web.");
}

function doPost(e) {
  try {
    const j = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const c = String(j.c || "").toUpperCase(), d = String(j.d || ""), img = String(j.img || "");
    if (!/^[A-Z0-9]{6,16}$/.test(c)) return resp_({ ok: false, erro: "codigo" });
    if (!/^\d{4}-\d\d-\d\d$/.test(d)) return resp_({ ok: false, erro: "data" });
    if (!img || img.length > MAX_BASE64 || !/^[A-Za-z0-9+/=]+$/.test(img.slice(0, 200))) return resp_({ ok: false, erro: "imagem" });
    const bytes = Utilities.base64Decode(img);
    if (!(bytes[0] === -1 && bytes[1] === -40)) return resp_({ ok: false, erro: "formato" }); // JPEG começa com FF D8
    const hora = Utilities.formatDate(new Date(), "America/Sao_Paulo", "HHmmss");
    const nome = "RL_" + c + "_" + d + "_" + hora + ".jpg";
    pasta_().createFile(Utilities.newBlob(bytes, "image/jpeg", nome));
    return resp_({ ok: true });
  } catch (err) {
    return resp_({ ok: false, erro: "falha" });
  }
}

function doGet() { return resp_({ ok: true, servico: "rumo-fotos" }); }

function pasta_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty("PASTA_ID");
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const it = DriveApp.getFoldersByName(PASTA_NOME);
  const pasta = it.hasNext() ? it.next() : DriveApp.createFolder(PASTA_NOME);
  props.setProperty("PASTA_ID", pasta.getId());
  return pasta;
}

function resp_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
