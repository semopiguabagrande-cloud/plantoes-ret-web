// =========================================================================
// PLANTÕES RET - BACKEND (Google Apps Script)
// =========================================================================

const ID_PLANILHA = "1vIsQ1JCh6lNJ9z5iazZLbt34e_wljL0IW80VePcjp4o";

// =====================================
// TEMPO DE EXPIRAÇÃO DA SESSÃO
// =====================================
const TEMPO_EXPIRACAO_SESSAO = 1800000;

// =====================================
// NOMES DE MESES (PT-BR, SEM ACENTO + abreviações)
// =====================================
const MESES_NOMES = {
  janeiro: "01", jan: "01",
  fevereiro: "02", fev: "02",
  marco: "03", mar: "03",
  abril: "04", abr: "04",
  maio: "05", mai: "05",
  junho: "06", jun: "06",
  julho: "07", jul: "07",
  agosto: "08", ago: "08",
  setembro: "09", set: "09",
  outubro: "10", out: "10",
  novembro: "11", nov: "11",
  dezembro: "12", dez: "12"
};

// =====================================
// NORMALIZA CÓDIGO
// =====================================
function normalizarCodigo(valor) {
  return String(valor || "").trim().replace(/^0+/, "");
}

function pad2(valor) {
  return String(Number(valor)).padStart(2, "0");
}

// =====================================
// CONFIG — MÊS ABERTO
// =====================================

let _cacheMesAberto = null;

function obterMesAberto() {
  if (_cacheMesAberto !== null) return _cacheMesAberto;

  let valor = "";
  let celulaUsada = "";

  try {
    const planilha = SpreadsheetApp.openById(ID_PLANILHA);
    const aba = planilha.getSheetByName("CONFIG");

    if (aba) {
      const candidatos = ["B1", "B2", "A2", "C1"];

      for (let i = 0; i < candidatos.length; i++) {
        const ref = candidatos[i];
        const celula = aba.getRange(ref);
        const bruto = celula.getValue();

        if (bruto instanceof Date) {
          _cacheMesAberto = Utilities.formatDate(
            bruto,
            Session.getScriptTimeZone(),
            "yyyy-MM"
          );

          Logger.log(
            "MÊS ABERTO (" + ref + ", data): '" + _cacheMesAberto + "'"
          );

          return _cacheMesAberto;
        }

        const texto = String(
          celula.getDisplayValue() || bruto || ""
        ).trim();

        if (texto && normalizarMesAberto(texto)) {
          valor = texto;
          celulaUsada = ref;
          break;
        }
      }
    }
  } catch (e) {
    Logger.log("Erro ao ler CONFIG: " + e);
  }

  _cacheMesAberto = normalizarMesAberto(valor);

  if (celulaUsada) {
    Logger.log(
      "MÊS ABERTO (" + celulaUsada + "): '" + _cacheMesAberto + "'"
    );
  } else {
    Logger.log(
      "MÊS ABERTO (final): '" + _cacheMesAberto + "'"
    );
  }

  return _cacheMesAberto;
}

function normalizarMesAberto(valor) {
  let texto = String(valor || "").trim().toLowerCase();

  if (!texto) return "";

  texto = texto
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");

  let m = texto.match(/^(\d{4})[-\/.](\d{1,2})$/);

  if (m) {
    return m[1] + "-" + pad2(m[2]);
  }

  m = texto.match(/^(\d{1,2})[-\/.](\d{4})$/);

  if (m) {
    return m[2] + "-" + pad2(m[1]);
  }

  m = texto.match(/^([a-z]+)[\s\/\-.]*(?:de\s*)?(\d{4})$/);

  if (m && MESES_NOMES[m[1]]) {
    return m[2] + "-" + MESES_NOMES[m[1]];
  }

  return "";
}

// =====================================
// MÊS DE UMA DATA "dd/MM/yyyy" → "yyyy-MM"
// =====================================
function mesDaData(dataTexto) {
  const partes = String(dataTexto || "").split("/");

  if (partes.length !== 3) return "";

  const mes = partes[1];
  const ano = partes[2];

  if (String(ano).length !== 4) return "";

  return ano + "-" + pad2(mes);
}

function pertenceAoMesAberto(dataTexto) {
  const aberto = obterMesAberto();

  if (!aberto) return true;

  return mesDaData(dataTexto) === aberto;
}

// =====================================
// doGet
// =====================================
function doGet(e) {
  try {

    const parametros = e && e.parameter ? e.parameter : {};

    const planilha = SpreadsheetApp.openById(ID_PLANILHA);

    const abaAgentes = planilha.getSheetByName("AGENTES");
    const abaVagas = planilha.getSheetByName("VAGAS");
    const abaInscricoes = planilha.getSheetByName("INSCRICOES");

    if (!abaAgentes || !abaVagas || !abaInscricoes) {
      return resposta({
        success: false,
        mensagem: "Uma ou mais abas não foram encontradas."
      });
    }

    const AGENTES = lerAgentes(abaAgentes);
    const tipo = String(parametros.tipo || "").trim();

    // =====================================
    // VALIDAÇÃO DA VERSÃO
    // =====================================
    const cliente = String(parametros.cliente || "").trim();
    const versao = String(parametros.versao || "").trim();

    const clientesPermitidos = [
      "flutter_web",
      "flutter_app"
    ];

    const versaoMinima = "2.0.0+2";

    if (
      !clientesPermitidos.includes(cliente) ||
      versao !== versaoMinima
    ) {
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          atualizacaoObrigatoria: true,
          mensagem:
            "Esta versão do aplicativo está desatualizada. Atualize o Plantões RET para continuar."
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // =====================================
    // LOGIN
    // =====================================
    if (tipo === "agente") {

      const codigo = normalizarCodigo(parametros.codigo);
      const deviceId = String(parametros.deviceId || "").trim();
      const forcar =
        String(parametros.forcar || "").trim() === "1";
      const tentativaId =
        String(parametros.tentativaId || "").trim();

      const dados = AGENTES;

      for (let i = 1; i < dados.length; i++) {

        const codigoLinha =
          normalizarCodigo(dados[i][0]);

        if (codigoLinha !== codigo) continue;

        const resultadoSessao =
          criarSessao(
            codigo,
            tentativaId,
            deviceId,
            forcar
          );

        if (!resultadoSessao.success) {
          return resposta(resultadoSessao);
        }

        return resposta({
          success: true,
          codigo: dados[i][0],
          matricula: dados[i][1],
          nome: dados[i][2],
          tipo: dados[i][3],
          ferias:
            String(dados[i][4] || "")
              .trim()
              .toUpperCase() === "SIM",
          sessionId: resultadoSessao.sessionId
        });
      }

      return resposta({
        success: false,
        mensagem: "Agente não encontrado."
      });
    }

    // =====================================
    // RECUPERAR SESSÃO
    // =====================================
    if (tipo === "recuperarSessao") {

      const codigo =
        normalizarCodigo(parametros.codigo);

      const sessionId =
        String(parametros.sessionId || "").trim();

      if (codigo === "" || sessionId === "") {
        return resposta({
          success: false,
          sessaoExpirada: true,
          sessaoRecuperada: false,
          mensagem:
            "Sessão inválida. Faça login novamente."
        });
      }

      if (!validarSessao(codigo, sessionId)) {
        return resposta({
          success: false,
          sessaoExpirada: true,
          sessaoRecuperada: false,
          mensagem:
            "Sua sessão foi encerrada. Faça login novamente."
        });
      }

      atualizarSessao(codigo, sessionId);

      let agenteEncontrado = null;

      for (let i = 1; i < AGENTES.length; i++) {

        if (
          normalizarCodigo(AGENTES[i][0]) === codigo
        ) {
          agenteEncontrado = AGENTES[i];
          break;
        }
      }

      if (!agenteEncontrado) {
        return resposta({
          success: false,
          sessaoExpirada: true,
          sessaoRecuperada: false,
          mensagem:
            "Agente não encontrado."
        });
      }

      return resposta({
        success: true,
        sessaoRecuperada: true,
        agente: {
          codigo: agenteEncontrado[0],
          matricula: agenteEncontrado[1],
          nome: agenteEncontrado[2],
          tipo: agenteEncontrado[3],
          ferias:
            String(agenteEncontrado[4] || "")
              .trim()
              .toUpperCase() === "SIM"
        },
        sessionId: sessionId
      });
    }

    // =====================================
    // DADOS INICIAIS
    // =====================================
    if (tipo === "inicial") {

      const codigo =
        normalizarCodigo(parametros.codigo);

      const sessionId =
        String(parametros.sessionId || "").trim();

      if (!codigo || sessionId === "") {
        return resposta({
          success: false,
          sessaoExpirada: true,
          mensagem:
            "Sessão inválida. Faça login novamente."
        });
      }

      if (!validarSessao(codigo, sessionId)) {
        return resposta({
          success: false,
          sessaoExpirada: true,
          mensagem:
            "Sua sessão foi encerrada. Faça login novamente."
        });
      }

      const resAtual =
        atualizarSessao(codigo, sessionId);

      if (!resAtual || !resAtual.success) {
        return resposta({
          success: false,
          sessaoExpirada: true,
          mensagem:
            "Sua sessão foi encerrada. Faça login novamente."
        });
      }

      let agente = null;

      for (let i = 1; i < AGENTES.length; i++) {

        if (
          normalizarCodigo(AGENTES[i][0]) === codigo
        ) {
          agente = {
            codigo: AGENTES[i][0],
            matricula: AGENTES[i][1],
            nome: AGENTES[i][2],
            tipo: AGENTES[i][3],
            ferias:
              String(AGENTES[i][4] || "")
                .trim()
                .toUpperCase() === "SIM"
          };

          break;
        }
      }

      if (!agente) {
        return resposta({
          success: false,
          mensagem: "Agente não encontrado."
        });
      }

      return resposta({
        success: true,
        agente: agente,
        mesAberto: obterMesAberto(),
        vagas:
          montarListaVagas(
            abaVagas,
            abaInscricoes
          ),
        minhas:
          montarMinhasInscricoes(
            codigo,
            abaInscricoes
          )
      });
    }

    // =====================================
    // MINHAS INSCRIÇÕES
    // =====================================
    if (tipo === "minhas") {

      const codigo =
        normalizarCodigo(parametros.codigo);

      const sessionId =
        String(parametros.sessionId || "").trim();

      if (!validarSessao(codigo, sessionId)) {
        return resposta({
          success: false,
          sessaoExpirada: true,
          mensagem:
            "Sua sessão foi encerrada. Faça login novamente."
        });
      }

      const r =
        atualizarSessao(codigo, sessionId);

      if (!r || !r.success) {
        return resposta({
          success: false,
          sessaoExpirada: true,
          mensagem:
            "Sua sessão foi encerrada. Faça login novamente."
        });
      }

      return resposta(
        montarMinhasInscricoes(
          codigo,
          abaInscricoes
        )
      );
    }

    // =====================================
    // CANCELAR INSCRIÇÃO
    // =====================================
    if (tipo === "cancelarInscricao") {

      const codigo =
        normalizarCodigo(parametros.codigo);

      const sessionId =
        String(parametros.sessionId || "").trim();

      if (!validarSessao(codigo, sessionId)) {
        return resposta({
          success: false,
          sessaoExpirada: true,
          mensagem:
            "Sua sessão foi encerrada. Faça login novamente."
        });
      }

      const lock =
        LockService.getScriptLock();

      try {

        lock.waitLock(30000);

        const listaDatas =
          JSON.parse(
            parametros.datas || "[]"
          );

        const dados =
          lerInscricoes(abaInscricoes);

        for (
          let i = dados.length - 1;
          i >= 1;
          i--
        ) {

          const codigoLinha =
            normalizarCodigo(dados[i][1]);

          const dataLinha =
            normalizarData(dados[i][4]);

          const turnoLinha =
            String(
              dados[i][5] || "UNICO"
            )
              .trim()
              .toUpperCase();

          const deveExcluir =
            listaDatas.some(item => {

              const dataItem =
                String(item.data || "").trim();

              const turnoItem =
                String(item.turno || "")
                  .trim()
                  .toUpperCase();

              return (
                dataItem === dataLinha &&
                (
                  turnoItem === "" ||
                  turnoItem === turnoLinha ||
                  (
                    turnoItem === "DIA" &&
                    turnoLinha === "UNICO"
                  )
                )
              );
            });

          if (
            codigoLinha === codigo &&
            deveExcluir
          ) {
            abaInscricoes.deleteRow(i + 1);
          }
        }

        SpreadsheetApp.flush();

        // O ScriptLock já está adquirido: atualiza a sessão sem
        // tentar adquirir o mesmo lock novamente.
        atualizarSessaoSemLock(codigo, sessionId);

        return resposta({
          success: true,
          mensagem: "Inscrição cancelada."
        });

      } finally {

        try {
          lock.releaseLock();
        } catch (e) {}
      }
    }

    // =====================================
    // BUSCAR VAGAS
    // =====================================
    if (tipo === "vagas") {

      const codigo =
        normalizarCodigo(parametros.codigo);

      const sessionId =
        String(parametros.sessionId || "").trim();

      if (
        codigo === "" ||
        sessionId === ""
      ) {
        return resposta({
          success: false,
          sessaoExpirada: true,
          mensagem:
            "Sessão inválida. Faça login novamente."
        });
      }

      if (!validarSessao(codigo, sessionId)) {
        return resposta({
          success: false,
          sessaoExpirada: true,
          mensagem:
            "Sua sessão foi encerrada. Faça login novamente."
        });
      }

      atualizarSessao(
        codigo,
        sessionId
      );

      return resposta(
        montarListaVagas(
          abaVagas,
          abaInscricoes
        )
      );
    }

    // =====================================
    // SALVAR INSCRIÇÃO
    // =====================================
    if (tipo === "salvarInscricao") {

      const codigo =
        normalizarCodigo(parametros.codigo);

      const sessionId =
        String(parametros.sessionId || "").trim();

      if (!validarSessao(codigo, sessionId)) {
        return resposta({
          success: false,
          sessaoExpirada: true,
          mensagem:
            "Sua sessão foi encerrada. Faça login novamente."
        });
      }

      const lock =
        LockService.getScriptLock();

      try {

        lock.waitLock(30000);

        // =====================================
        // LOCALIZA AGENTE
        // =====================================
        const agenteAutenticado =
          AGENTES.find((linha, indice) => {

            if (indice === 0) return false;

            return (
              normalizarCodigo(linha[0]) === codigo
            );
          });

        if (!agenteAutenticado) {
          return resposta({
            success: false,
            mensagem:
              "Agente não encontrado na planilha."
          });
        }

        const matricula =
          String(
            agenteAutenticado[1] || ""
          ).trim();

        const nome =
          String(
            agenteAutenticado[2] || ""
          ).trim();

        if (!matricula || !nome) {
          return resposta({
            success: false,
            mensagem:
              "Matrícula ou nome do agente não cadastrado corretamente."
          });
        }

        // A sessão será renovada sem novo LockService dentro
        // da seção crítica, evitando espera/deadlock e lentidão.

        // =====================================
        // LÊ SOLICITAÇÕES
        // =====================================
        let listaDatas = [];

        try {

          listaDatas =
            JSON.parse(
              parametros.datas || "[]"
            );

        } catch (erro) {

          return resposta({
            success: false,
            mensagem:
              "Formato de datas inválido."
          });
        }

        if (!Array.isArray(listaDatas)) {
          return resposta({
            success: false,
            mensagem:
              "Lista de datas inválida."
          });
        }

        if (listaDatas.length === 0) {
          return resposta({
            success: false,
            mensagem:
              "Selecione pelo menos uma data."
          });
        }

        // =====================================
        // VALIDAÇÃO DAS DATAS
        // =====================================
        const mesAberto =
          obterMesAberto();

        const chavesSolicitadas =
          new Set();

        for (
          let i = 0;
          i < listaDatas.length;
          i++
        ) {

          const item =
            listaDatas[i];

          if (
            !item ||
            typeof item !== "object"
          ) {
            return resposta({
              success: false,
              mensagem:
                "Uma das preferências está em formato inválido."
            });
          }

          const dataSolicitada =
            String(
              item.data || ""
            ).trim();

          const turnoSolicitado =
            String(
              item.turno || "DIA"
            )
              .trim()
              .toUpperCase();

          if (
            !/^\d{2}\/\d{2}\/\d{4}$/.test(
              dataSolicitada
            )
          ) {
            return resposta({
              success: false,
              mensagem:
                "Data inválida: " +
                dataSolicitada
            });
          }

          const partes =
            dataSolicitada.split("/");

          const diaNum =
            Number(partes[0]);

          const mesNum =
            Number(partes[1]);

          const anoNum =
            Number(partes[2]);

          const dataTeste =
            new Date(
              anoNum,
              mesNum - 1,
              diaNum
            );

          if (
            anoNum < 2000 ||
            mesNum < 1 ||
            mesNum > 12 ||
            diaNum < 1 ||
            dataTeste.getFullYear() !== anoNum ||
            dataTeste.getMonth() !== mesNum - 1 ||
            dataTeste.getDate() !== diaNum
          ) {
            return resposta({
              success: false,
              mensagem:
                "Data inexistente: " +
                dataSolicitada
            });
          }

          if (
            mesAberto &&
            mesDaData(dataSolicitada) !== mesAberto
          ) {
            return resposta({
              success: false,
              mensagem:
                "As vagas abertas são de outro mês. Atualize a tela e escolha novamente."
            });
          }

          if (
            ![
              "DIA",
              "NOITE",
              "UNICO"
            ].includes(turnoSolicitado)
          ) {
            return resposta({
              success: false,
              mensagem:
                "Turno inválido para " +
                dataSolicitada +
                "."
            });
          }

          const chave =
            dataSolicitada +
            "|" +
            turnoSolicitado;

          if (
            chavesSolicitadas.has(chave)
          ) {
            return resposta({
              success: false,
              mensagem:
                "A mesma data e turno foram selecionados mais de uma vez."
            });
          }

          chavesSolicitadas.add(chave);
        }

        // =====================================
        // LÊ VAGAS E INSCRIÇÕES
        // =====================================
        const vagas =
          lerVagas(abaVagas);

        const inscricoes =
          lerInscricoes(abaInscricoes);

        // =====================================
        // LIMITE MENSAL DO AGENTE
        // =====================================
        let limite = 8;

        const agente =
          AGENTES.find(
            (linha, indice) => {

              if (indice === 0) return false;

              return (
                normalizarCodigo(linha[0]) === codigo
              );
            }
          );

        if (agente) {

          const ferias =
            String(
              agente[4] || ""
            )
              .trim()
              .toUpperCase();

          if (ferias === "SIM") {
            limite = 10;
          }
        }

        let totalInscricoesMes = 0;

        const primeiraData =
          listaDatas.length > 0
            ? String(
                listaDatas[0].data || ""
              ).trim()
            : "";

        let mesAtual = "";
        let anoAtual = "";

        if (primeiraData !== "") {

          const partes =
            primeiraData.split("/");

          if (partes.length === 3) {

            mesAtual = partes[1];
            anoAtual = partes[2];
          }
        }

        for (
          let i = 1;
          i < inscricoes.length;
          i++
        ) {

          if (
            normalizarCodigo(
              inscricoes[i][1]
            ) !== codigo
          ) {
            continue;
          }

          const partesData =
            normalizarData(
              inscricoes[i][4]
            ).split("/");

          if (
            partesData.length === 3 &&
            partesData[1] === mesAtual &&
            partesData[2] === anoAtual
          ) {
            totalInscricoesMes++;
          }
        }

        // =====================================
        // PASSAGEM 1
        // VALIDA TUDO ANTES DE GRAVAR
        // =====================================
        const paraGravar = [];

        const ocupadasPorData = {};

        // =====================================
        // PRIMEIRO: CONTA TODAS AS INSCRIÇÕES
        // EXISTENTES POR DATA E TURNO
        // =====================================
        for (
          let x = 1;
          x < inscricoes.length;
          x++
        ) {

          const dataExistente =
            normalizarData(
              inscricoes[x][4]
            );

          if (!dataExistente) continue;

          if (!ocupadasPorData[dataExistente]) {

            ocupadasPorData[dataExistente] = {
              dia: 0,
              noite: 0,
              total: 0
            };
          }

          const turnoExistente =
            String(
              inscricoes[x][5] || "UNICO"
            )
              .trim()
              .toUpperCase();

          // =====================================
          // DIA / UNICO
          // =====================================
          if (
            turnoExistente === "DIA" ||
            turnoExistente === "UNICO" ||
            turnoExistente === ""
          ) {

            ocupadasPorData[dataExistente].dia++;
            ocupadasPorData[dataExistente].total++;

          }

          // =====================================
          // NOITE
          // =====================================
          else if (
            turnoExistente === "NOITE"
          ) {

            ocupadasPorData[dataExistente].noite++;
            ocupadasPorData[dataExistente].total++;
          }

          // =====================================
          // QUALQUER OUTRO VALOR DE TURNO
          // TAMBÉM CONTA PARA O LIMITE TOTAL
          // =====================================
          else {

            ocupadasPorData[dataExistente].total++;
          }
        }

        // =====================================
        // SEGUNDA: VALIDA CADA SOLICITAÇÃO
        // =====================================
        for (
          let i = 0;
          i < listaDatas.length;
          i++
        ) {

          const item =
            listaDatas[i];

          const data =
            String(
              item.data || ""
            ).trim();

          const turno =
            String(
              item.turno || "DIA"
            )
              .trim()
              .toUpperCase();

          if (data === "") continue;

          // =====================================
          // NÃO PERMITE O MESMO AGENTE
          // EM DOIS TURNOS DO MESMO DIA
          // =====================================
          let jaTemNoDia = false;

          for (
            let j = 1;
            j < inscricoes.length;
            j++
          ) {

            if (
              normalizarCodigo(
                inscricoes[j][1]
              ) !== codigo
            ) {
              continue;
            }

            if (
              normalizarData(
                inscricoes[j][4]
              ) === data
            ) {
              jaTemNoDia = true;
              break;
            }
          }

          if (jaTemNoDia) {

            return resposta({
              success: false,
              mensagem:
                "Você já possui um plantão neste dia."
            });
          }

          // =====================================
          // LIMITE MENSAL
          // =====================================
          if (
            totalInscricoesMes >= limite
          ) {

            return resposta({
              success: false,
              tipoLimite: "AGENTE",
              limite: limite,
              ocupadas: totalInscricoesMes,
              mensagem:
                "Você já atingiu o limite mensal de " + limite +
                " plantões. Cancele um plantão confirmado para escolher outra data."
            });
          }

          // =====================================
          // LOCALIZA VAGAS DA DATA
          // =====================================
          let totalDia = 0;
          let totalNoite = 0;
          let encontrouData = false;

          for (
            let x = 1;
            x < vagas.length;
            x++
          ) {

            if (
              normalizarData(
                vagas[x][2]
              ) !== data
            ) {
              continue;
            }

            encontrouData = true;

            totalDia = Math.max(0, Math.floor(Number(vagas[x][3]) || 0));
            totalNoite = Math.max(0, Math.floor(Number(vagas[x][4]) || 0));

            break;
          }

          if (!encontrouData) {

            return resposta({
              success: false,
              mensagem:
                "Não existem vagas cadastradas para a data " +
                data +
                "."
            });
          }

          totalDia =
            Math.max(
              totalDia,
              0
            );

          totalNoite =
            Math.max(
              totalNoite,
              0
            );

          // UNICO só é válido para datas sem turno noturno.
          // Em datas com DIA e NOITE, a inscrição deve informar o turno exato.
          if (turno === "UNICO" && totalNoite > 0) {
            return resposta({
              success: false,
              mensagem: "Selecione DIA ou NOITE para " + data + ".",
              data: data,
              turno: turno
            });
          }

          const limiteTotalData =
            totalDia + totalNoite;

          if (
            limiteTotalData <= 0
          ) {

            return resposta({
              success: false,
              mensagem:
                "Não existem vagas cadastradas para a data " +
                data +
                "."
            });
          }

          if (
            !ocupadasPorData[data]
          ) {

            ocupadasPorData[data] = {
              dia: 0,
              noite: 0,
              total: 0
            };
          }

          // =====================================
          // QUANTIDADE ATUAL DA DATA
          // =====================================
          const ocupadasData =
            ocupadasPorData[data].total;

          // =====================================
          // NOVA BARREIRA:
          // LIMITE TOTAL DA DATA
          //
          // EXEMPLO:
          // 6 DIA + 2 NOITE = 8
          //
          // SE JÁ EXISTEM 8:
          // NÃO ACEITA MAIS NENHUMA.
          // =====================================
          if (
            ocupadasData >=
            limiteTotalData
          ) {

            return resposta({
              success: false,
              tipoLimite: "DIA",
              limite: limiteTotalData,
              ocupadas: ocupadasData,
              data: data,
              turno: turno,
              mensagem:
                "Todas as vagas de " + data + " foram preenchidas. " +
                "O limite deste dia é de " + limiteTotalData +
                " plantões. Escolha outra data."
            });
          }

          // =====================================
          // VERIFICA LIMITE DO TURNO
          // =====================================
          let ocupadasTurno = 0;
          let limiteTurno = 0;

          if (turno === "NOITE") {

            ocupadasTurno =
              ocupadasPorData[data].noite;

            limiteTurno =
              totalNoite;

          } else {

            ocupadasTurno =
              ocupadasPorData[data].dia;

            limiteTurno =
              totalDia;
          }

          // =====================================
          // LIMITE DO TURNO
          // =====================================
          if (
            ocupadasTurno >=
            limiteTurno
          ) {

            return resposta({
              success: false,
              tipoLimite: "TURNO",
              limite: limiteTurno,
              ocupadas: ocupadasTurno,
              data: data,
              turno: turno,
              mensagem:
                "As vagas do turno " + turno + " em " + data +
                " foram preenchidas. Escolha outro turno ou outra data."
            });
          }

          // =====================================
          // RESERVA TEMPORÁRIA
          // PARA VÁRIAS DATAS NA MESMA REQUISIÇÃO
          // =====================================
          if (turno === "NOITE") {

            ocupadasPorData[data].noite++;

          } else {

            ocupadasPorData[data].dia++;
          }

          // =====================================
          // TOTAL DA DATA TAMBÉM É INCREMENTADO
          // =====================================
          ocupadasPorData[data].total++;

          paraGravar.push({
            data: data,
            turno: turno
          });

          totalInscricoesMes++;
        }

        // =====================================
        // NADA PARA GRAVAR
        // =====================================
        if (paraGravar.length === 0) {

          return resposta({
            success: false,
            mensagem:
              "Selecione pelo menos uma data."
          });
        }

        // =====================================
        // PASSAGEM 2
        // GRAVA TUDO
        // =====================================
        for (
          let i = 0;
          i < paraGravar.length;
          i++
        ) {

          abaInscricoes.appendRow([

            new Date().getTime(),

            codigo,

            matricula,

            nome,

            paraGravar[i].data,

            paraGravar[i].turno,

            new Date()
          ]);
        }

        SpreadsheetApp.flush();

        atualizarSessaoSemLock(codigo, sessionId);

        return resposta({
          success: true,
          mensagem:
            "Inscrição salva."
        });

      } finally {

        try {
          lock.releaseLock();
        } catch (e) {}
      }
    }

    // =====================================
    // RELATÓRIO — ADMIN
    // =====================================
    if (tipo === "relatorio") {

      const codigo =
        normalizarCodigo(parametros.codigo);

      const sessionId =
        String(parametros.sessionId || "").trim();

      if (
        codigo === "" ||
        sessionId === ""
      ) {
        return resposta({
          success: false,
          sessaoExpirada: true,
          mensagem:
            "Sessão inválida. Faça login novamente."
        });
      }

      if (!validarSessao(codigo, sessionId)) {
        return resposta({
          success: false,
          sessaoExpirada: true,
          mensagem:
            "Sua sessão foi encerrada. Faça login novamente."
        });
      }

      let usuarioAdmin = false;

      for (
        let i = 1;
        i < AGENTES.length;
        i++
      ) {

        const codigoLinha =
          normalizarCodigo(
            AGENTES[i][0]
          );

        if (codigoLinha !== codigo) continue;

        const tipoAgente =
          String(
            AGENTES[i][3] || ""
          )
            .trim()
            .toUpperCase();

        if (tipoAgente === "ADMIN") {
          usuarioAdmin = true;
        }

        break;
      }

      if (!usuarioAdmin) {
        return resposta({
          success: false,
          mensagem:
            "Acesso negado. Somente administradores podem consultar o relatório."
        });
      }

      const resultadoAtualizacao =
        atualizarSessao(
          codigo,
          sessionId
        );

      if (
        !resultadoAtualizacao ||
        !resultadoAtualizacao.success
      ) {
        return resposta({
          success: false,
          sessaoExpirada: true,
          mensagem:
            "Não foi possível validar sua sessão. Faça login novamente."
        });
      }

      const dados =
        lerInscricoes(
          abaInscricoes
        );

      const lista = [];

      for (
        let i = 1;
        i < dados.length;
        i++
      ) {

        lista.push({
          codigo: dados[i][1],
          matricula: dados[i][2],
          nome: dados[i][3],
          data: normalizarData(
            dados[i][4]
          ),
          turno:
            String(
              dados[i][5] || "UNICO"
            )
              .trim()
              .toUpperCase()
        });
      }

      lista.sort((a, b) => {

        const comparaNome =
          String(a.nome || "")
            .localeCompare(
              String(b.nome || ""),
              "pt-BR"
            );

        if (comparaNome !== 0) {
          return comparaNome;
        }

        const dataA =
          new Date(
            a.data
              .split("/")
              .reverse()
              .join("-")
          );

        const dataB =
          new Date(
            b.data
              .split("/")
              .reverse()
              .join("-")
          );

        return dataA - dataB;
      });

      return resposta(lista);
    }

    // =====================================
    // HEARTBEAT
    // =====================================
    if (tipo === "heartbeat") {

      const codigo =
        normalizarCodigo(
          parametros.codigo
        );

      const sessionId =
        String(
          parametros.sessionId || ""
        ).trim();

      return resposta(
        atualizarSessao(
          codigo,
          sessionId
        )
      );
    }

    // =====================================
    // LOGOUT
    // =====================================
    if (tipo === "logout") {

      const codigo =
        normalizarCodigo(
          parametros.codigo
        );

      const sessionId =
        String(
          parametros.sessionId || ""
        ).trim();

      return resposta(
        encerrarSessao(
          codigo,
          sessionId
        )
      );
    }

    return resposta({
      success: false,
      mensagem: "Tipo inválido."
    });

  } catch (erro) {

    Logger.log(erro);

    return resposta({
      success: false,
      mensagem: erro.toString()
    });
  }
}

// =====================================
// VALIDAR SESSÃO
// =====================================
function validarSessao(codigo, sessionId) {

  const aba =
    obterAbaSessoes();

  const dados =
    aba.getDataRange().getValues();

  const agora =
    Date.now();

  for (
    let i = 1;
    i < dados.length;
    i++
  ) {

    const codigoLinha =
      normalizarCodigo(
        dados[i][0]
      );

    const sessaoLinha =
      String(
        dados[i][1] || ""
      ).trim();

    if (
      codigoLinha === codigo &&
      sessaoLinha === sessionId
    ) {

      const ultimoAcesso =
        dados[i][2]
          ? new Date(
              dados[i][2]
            ).getTime()
          : 0;

      if (
        ultimoAcesso > 0 &&
        agora - ultimoAcesso <
          TEMPO_EXPIRACAO_SESSAO
      ) {
        return true;
      }

      return false;
    }
  }

  return false;
}

// =====================================
// NORMALIZA DATA
// =====================================
function normalizarData(valor) {

  if (valor instanceof Date) {

    return Utilities.formatDate(
      valor,
      Session.getScriptTimeZone(),
      "dd/MM/yyyy"
    );
  }

  const texto =
    String(valor).trim();

  if (texto.indexOf("T") > -1) {

    return Utilities.formatDate(
      new Date(texto),
      Session.getScriptTimeZone(),
      "dd/MM/yyyy"
    );
  }

  return texto;
}

// =====================================
// MONTA LISTA DE VAGAS
// =====================================
function montarListaVagas(
  abaVagas,
  abaInscricoes
) {

  const vagas =
    lerVagas(abaVagas);

  const inscricoes =
    lerInscricoes(
      abaInscricoes
    );

  const ocupacao = {};

  for (
    let i = 1;
    i < inscricoes.length;
    i++
  ) {

    const data =
      normalizarData(
        inscricoes[i][4]
      );

    const turno =
      String(
        inscricoes[i][5] || "UNICO"
      )
        .trim()
        .toUpperCase();

    if (!ocupacao[data]) {

      ocupacao[data] = {
        dia: 0,
        noite: 0
      };
    }

    if (
      turno === "UNICO" ||
      turno === "DIA" ||
      turno === ""
    ) {
      ocupacao[data].dia++;
    }

    if (
      turno === "NOITE"
    ) {
      ocupacao[data].noite++;
    }
  }

  const resultado = [];

  for (
    let i = 1;
    i < vagas.length;
    i++
  ) {

    const dataVaga =
      normalizarData(
        vagas[i][2]
      );

    if (
      !pertenceAoMesAberto(
        dataVaga
      )
    ) {
      continue;
    }

    const vagasDia =
      Number(
        vagas[i][3] || 0
      );

    const vagasNoite =
      Number(
        vagas[i][4] || 0
      );

    const info =
      ocupacao[dataVaga] || {
        dia: 0,
        noite: 0
      };

    resultado.push({

      data: dataVaga,

      vagasDia: vagasDia,

      vagasNoite: vagasNoite,

      possuiNoite:
        vagasNoite > 0,

      ocupadasDia:
        info.dia,

      ocupadasNoite:
        info.noite,

      restantesDia:
        Math.max(
          vagasDia - info.dia,
          0
        ),

      restantesNoite:
        Math.max(
          vagasNoite - info.noite,
          0
        )
    });
  }

  resultado.sort((a, b) => {

    const da =
      a.data
        .split("/")
        .reverse()
        .join("-");

    const db =
      b.data
        .split("/")
        .reverse()
        .join("-");

    return new Date(da) -
      new Date(db);
  });

  return resultado;
}

// =====================================
// MONTA MINHAS INSCRIÇÕES
// =====================================
function montarMinhasInscricoes(
  codigo,
  abaInscricoes
) {

  const dados =
    lerInscricoes(
      abaInscricoes
    );

  const lista = [];

  const codigoNormalizado =
    normalizarCodigo(
      codigo
    );

  for (
    let i = 1;
    i < dados.length;
    i++
  ) {

    if (
      normalizarCodigo(
        dados[i][1]
      ) !== codigoNormalizado
    ) {
      continue;
    }

    const dataNormalizada =
      normalizarData(
        dados[i][4]
      );

    if (
      !pertenceAoMesAberto(
        dataNormalizada
      )
    ) {
      continue;
    }

    lista.push({

      data:
        dataNormalizada,

      turno:
        String(
          dados[i][5] || "UNICO"
        )
          .trim()
          .toUpperCase()
    });
  }

  return lista;
}

// =====================================
// JSON
// =====================================
function resposta(obj) {

  return ContentService
    .createTextOutput(
      JSON.stringify(obj)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}

// =====================================
// LEITURA DAS ABAS
// =====================================
function lerAgentes(abaAgentes) {
  return abaAgentes
    .getDataRange()
    .getValues();
}

function lerVagas(abaVagas) {
  return abaVagas
    .getDataRange()
    .getValues();
}

function lerInscricoes(abaInscricoes) {
  return abaInscricoes
    .getDataRange()
    .getValues();
}

// =====================================
// ABA SESSOES
// CODIGO | SESSION_ID | ULTIMO_ACESSO |
// TENTATIVA_ID | DEVICE_ID | CONFIRMADO
// =====================================
function obterAbaSessoes() {

  const planilha =
    SpreadsheetApp.openById(
      ID_PLANILHA
    );

  let aba =
    planilha.getSheetByName(
      "SESSOES"
    );

  if (!aba) {

    aba =
      planilha.insertSheet(
        "SESSOES"
      );

    aba
      .getRange(
        1,
        1,
        1,
        6
      )
      .setValues([[
        "CODIGO",
        "SESSION_ID",
        "ULTIMO_ACESSO",
        "TENTATIVA_ID",
        "DEVICE_ID",
        "CONFIRMADO"
      ]]);
  }

  garantirCabecalhoSessoes(
    aba
  );

  return aba;
}

// =====================================
// GARANTIR CABEÇALHO SESSÕES
// =====================================
function garantirCabecalhoSessoes(
  aba
) {

  if (
    aba.getMaxColumns() < 6
  ) {

    aba.insertColumnsAfter(
      aba.getMaxColumns(),
      6 - aba.getMaxColumns()
    );
  }

  const cab =
    aba
      .getRange(
        1,
        1,
        1,
        6
      )
      .getValues()[0];

  const nomes = [
    "CODIGO",
    "SESSION_ID",
    "ULTIMO_ACESSO",
    "TENTATIVA_ID",
    "DEVICE_ID",
    "CONFIRMADO"
  ];

  for (
    let c = 0;
    c < 6;
    c++
  ) {

    if (
      String(
        cab[c] || ""
      ).trim() === ""
    ) {

      aba
        .getRange(
          1,
          c + 1
        )
        .setValue(
          nomes[c]
        );
    }
  }
}

// =====================================
// CRIAR SESSÃO
// =====================================
function criarSessao(
  codigo,
  tentativaId,
  deviceId,
  forcar
) {

  const lock =
    LockService.getScriptLock();

  try {

    lock.waitLock(30000);

    const aba =
      obterAbaSessoes();

    const dados =
      aba
        .getDataRange()
        .getValues();

    const agora =
      Date.now();

    const tentativa =
      String(
        tentativaId || ""
      ).trim();

    const device =
      String(
        deviceId || ""
      ).trim();

    const forcarLogin =
      forcar === true;

    if (!tentativa) {

      return {
        success: false,
        mensagem:
          "Identificador da tentativa de login ausente."
      };
    }

    for (
      let i = 1;
      i < dados.length;
      i++
    ) {

      const codigoLinha =
        normalizarCodigo(
          dados[i][0]
        );

      if (
        codigoLinha !== codigo
      ) {
        continue;
      }

      const sessionId =
        String(
          dados[i][1] || ""
        ).trim();

      const ultimoAcesso =
        dados[i][2]
          ? new Date(
              dados[i][2]
            ).getTime()
          : 0;

      const tentativaAnterior =
        String(
          dados[i][3] || ""
        ).trim();

      const deviceAnterior =
        String(
          dados[i][4] || ""
        ).trim();

      const confirmado =
        String(
          dados[i][5] || ""
        )
          .trim()
          .toUpperCase();

      const ativa =
        ultimoAcesso > 0 &&
        agora - ultimoAcesso <
          TEMPO_EXPIRACAO_SESSAO;

      const mesmaTentativa =
        tentativaAnterior !== "" &&
        tentativaAnterior === tentativa;

      const mesmoDispositivo =
        device !== "" &&
        deviceAnterior !== "" &&
        deviceAnterior === device;

      // MESMO APARELHO / MESMA TENTATIVA
      if (
        ativa &&
        sessionId &&
        (
          mesmaTentativa ||
          mesmoDispositivo
        )
      ) {

        const row = i + 1;

        aba
          .getRange(row, 3)
          .setValue(
            new Date(agora)
          );

        if (
          !deviceAnterior &&
          device
        ) {

          aba
            .getRange(row, 5)
            .setValue(device);
        }

        SpreadsheetApp.flush();

        return {
          success: true,
          sessionId:
            sessionId,
          adotada: true
        };
      }

      // SESSÃO FANTASMA
      const fantasma =
        confirmado !== "SIM";

      if (
        ativa &&
        !forcarLogin &&
        !fantasma
      ) {

        return {
          success: false,
          codigoConflito: true,
          mensagem:
            "Este agente já está conectado em outro dispositivo."
        };
      }

      // CRIA NOVA
      const novaSessao =
        Utilities.getUuid();

      aba
        .getRange(
          i + 1,
          1,
          1,
          6
        )
        .setValues([[
          codigo,
          novaSessao,
          new Date(agora),
          tentativa,
          device,
          "NAO"
        ]]);

      SpreadsheetApp.flush();

      return {
        success: true,
        sessionId:
          novaSessao,
        adotada: false
      };
    }

    // PRIMEIRA SESSÃO
    const novaSessao =
      Utilities.getUuid();

    aba.appendRow([
      codigo,
      novaSessao,
      new Date(agora),
      tentativa,
      device,
      "NAO"
    ]);

    SpreadsheetApp.flush();

    return {
      success: true,
      sessionId:
        novaSessao,
      adotada: false
    };

  } finally {

    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

// =====================================
// ATUALIZAR SESSÃO / HEARTBEAT
// =====================================
// Atualiza sessão quando o chamador JÁ possui ScriptLock.
// Não chama LockService novamente: evita bloqueio recursivo e lentidão.
function atualizarSessaoSemLock(codigo, sessionId) {
  const codigoNormalizado = normalizarCodigo(codigo);
  const sessaoNormalizada = String(sessionId || "").trim();
  if (!codigoNormalizado || !sessaoNormalizada) return false;

  const aba = obterAbaSessoes();
  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha < 2) return false;

  const dados = aba.getRange(2, 1, ultimaLinha - 1, 6).getValues();
  for (let i = 0; i < dados.length; i++) {
    if (
      normalizarCodigo(dados[i][0]) === codigoNormalizado &&
      String(dados[i][1] || "").trim() === sessaoNormalizada
    ) {
      const linha = i + 2;
      aba.getRange(linha, 3, 1, 1).setValue(new Date());
      aba.getRange(linha, 6, 1, 1).setValue("SIM");
      return true;
    }
  }
  return false;
}

function atualizarSessao(
  codigo,
  sessionId
) {

  const codigoNormalizado =
    normalizarCodigo(codigo);

  const sessaoNormalizada =
    String(
      sessionId || ""
    ).trim();

  if (
    codigoNormalizado === "" ||
    sessaoNormalizada === ""
  ) {

    return {
      success: false,
      sessaoExpirada: true,
      mensagem:
        "Sessão inválida ou encerrada."
    };
  }

  const lock =
    LockService.getScriptLock();

  try {

    lock.waitLock(30000);

    const aba =
      obterAbaSessoes();

    const dados =
      aba
        .getDataRange()
        .getValues();

    const agora =
      Date.now();

    for (
      let i = 1;
      i < dados.length;
      i++
    ) {

      const codigoLinha =
        normalizarCodigo(
          dados[i][0]
        );

      const sessaoLinha =
        String(
          dados[i][1] || ""
        ).trim();

      if (
        codigoLinha !==
          codigoNormalizado ||
        sessaoLinha !==
          sessaoNormalizada
      ) {
        continue;
      }

      const ultimoAcesso =
        dados[i][2]
          ? new Date(
              dados[i][2]
            ).getTime()
          : 0;

      if (
        !ultimoAcesso ||
        isNaN(ultimoAcesso) ||
        agora - ultimoAcesso >=
          TEMPO_EXPIRACAO_SESSAO
      ) {

        aba.deleteRow(
          i + 1
        );

        SpreadsheetApp.flush();

        return {
          success: false,
          sessaoExpirada: true,
          mensagem:
            "Sua sessão expirou. Faça login novamente."
        };
      }

      aba
        .getRange(
          i + 1,
          3
        )
        .setValue(
          new Date(agora)
        );

      aba
        .getRange(
          i + 1,
          6
        )
        .setValue("SIM");

      SpreadsheetApp.flush();

      return {
        success: true
      };
    }

    return {
      success: false,
      sessaoExpirada: true,
      mensagem:
        "Sessão inválida ou encerrada."
    };

  } finally {

    try {
      lock.releaseLock();
    } catch (e) {}
  }
}

// =====================================
// ENCERRAR SESSÃO / LOGOUT
// =====================================
function encerrarSessao(
  codigo,
  sessionId
) {

  const codigoNormalizado =
    normalizarCodigo(codigo);

  const sessaoNormalizada =
    String(
      sessionId || ""
    ).trim();

  const lock =
    LockService.getScriptLock();

  try {

    lock.waitLock(30000);

    const aba =
      obterAbaSessoes();

    const dados =
      aba
        .getDataRange()
        .getValues();

    for (
      let i = dados.length - 1;
      i >= 1;
      i--
    ) {

      const codigoLinha =
        normalizarCodigo(
          dados[i][0]
        );

      const sessaoLinha =
        String(
          dados[i][1] || ""
        ).trim();

      if (
        codigoLinha ===
          codigoNormalizado &&
        sessaoLinha ===
          sessaoNormalizada
      ) {

        aba.deleteRow(
          i + 1
        );

        SpreadsheetApp.flush();

        return {
          success: true
        };
      }
    }

    return {
      success: true
    };

  } finally {

    try {
      lock.releaseLock();
    } catch (e) {}
  }
}
