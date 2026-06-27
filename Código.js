// =========================================================================
// BACKEND - REGISTROS TUBERCULOSE
// =========================================================================

const NOME_ABA_PROFISSIONAIS = 'Login_Profissional';
const ID_PLANILHA_TB = '1Qonf3_b91auVPxe41EoRMr-x6ZyB3MVWo8aJVyoLCR8';

function obterPlanilha() {
  const ativa = SpreadsheetApp.getActiveSpreadsheet();
  if (ativa) return ativa;
  return SpreadsheetApp.openById(ID_PLANILHA_TB);
}

// 1. Função principal que renderiza o Web App
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Registros Tuberculose - SMS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 2. Validação de Acesso (Login via CPF)
function validarProfissional(cpfDigitado) {
  const ss = obterPlanilha();
  const sheet = ss.getSheetByName(NOME_ABA_PROFISSIONAIS);

  if (!sheet) {
    throw new Error("Aba '" + NOME_ABA_PROFISSIONAIS + "' não encontrada no banco de dados.");
  }

  const dados = sheet.getDataRange().getValues();
  const cpfLimpo = cpfDigitado.replace(/\D/g, "");

  // O loop começa em 1 para pular a linha de cabeçalho
  for (let i = 1; i < dados.length; i++) {
    let cpfPlanilha = String(dados[i][0]).replace(/\D/g, "");

    if (cpfPlanilha === cpfLimpo) {
      return {
        sucesso: true,
        nome: dados[i][1],     // Coluna B: Nome do profissional
        unidade: dados[i][2],  // Coluna C: Unidade
        funcao: dados[i][3],   // Coluna D: Função (ex: Enfermeiro)
        cpf: cpfDigitado       // CPF formatado
      };
    }
  }

  return { sucesso: false, erro: "CPF não localizado na base de profissionais." };
}

function obterDadosProfissionalPorCpf(cpf) {
  const ss = obterPlanilha();
  const sheet = ss.getSheetByName(NOME_ABA_PROFISSIONAIS);

  if (!sheet) {
    throw new Error("Aba '" + NOME_ABA_PROFISSIONAIS + "' não encontrada no banco de dados.");
  }

  const cpfLimpo = String(cpf).replace(/\D/g, "");
  const dados = sheet.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]).replace(/\D/g, "") === cpfLimpo) {
      return {
        nome: dados[i][1],
        unidade: dados[i][2]
      };
    }
  }

  throw new Error("Profissional não encontrado para o CPF informado.");
}

const NOME_ABA_LIVRO_VERDE = 'Livro Verde - Tratamento';
const COL_INICIO_DADOS_LIVRO_VERDE = 4; // após as 4 colunas automáticas

const CHAVES_CAMPOS_LIVRO_VERDE = [
  'sinan', 'prontuario', 'nomePaciente', 'idade', 'sexo',
  'bac1', 'bac2', 'trm', 'cultura', 'culturaTS', 'rx', 'hiv',
  'formaClinica', 'tipoEntrada', 'esquema', 'dataInicio', 'tdo', 'tarv'
];

const CHAVES_ENCERRAMENTO_LIVRO_VERDE = [
  'motivoEnc', 'dataEnc', 'contIdent', 'contExam', 'obs'
];

function celulaVazia(valor) {
  return valor === null || valor === undefined || String(valor).trim() === '';
}

function formatarValorCelula(valor) {
  if (celulaVazia(valor)) return '';
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(valor).trim();
}

function montarDadosLivroVerdeDaLinha(linha) {
  const dados = {};
  const preenchidos = {};

  CHAVES_CAMPOS_LIVRO_VERDE.forEach(function(chave, indice) {
    const valor = formatarValorCelula(linha[COL_INICIO_DADOS_LIVRO_VERDE + indice]);
    dados[chave] = valor;
    preenchidos[chave] = valor !== '';
  });

  const meses = [];
  const inicioMeses = COL_INICIO_DADOS_LIVRO_VERDE + CHAVES_CAMPOS_LIVRO_VERDE.length;
  for (let i = 0; i < 12; i++) {
    const valor = formatarValorCelula(linha[inicioMeses + i]);
    meses.push(valor);
    preenchidos['mes' + (i + 1)] = valor !== '';
  }
  dados.meses = meses;

  const inicioEncerramento = inicioMeses + 12;
  CHAVES_ENCERRAMENTO_LIVRO_VERDE.forEach(function(chave, indice) {
    const valor = formatarValorCelula(linha[inicioEncerramento + indice]);
    dados[chave] = valor;
    preenchidos[chave] = valor !== '';
  });

  return { dados: dados, preenchidos: preenchidos };
}

function buscarRegistroPorSinan(sinan) {
  const ss = obterPlanilha();
  const sheet = ss.getSheetByName(NOME_ABA_LIVRO_VERDE);

  if (!sheet) {
    throw new Error("Aba '" + NOME_ABA_LIVRO_VERDE + "' não encontrada no banco de dados.");
  }

  const sinanBusca = String(sinan).trim();
  if (!sinanBusca) {
    throw new Error('Informe o número do SINAN para buscar.');
  }

  const dados = sheet.getDataRange().getValues();
  let linhaEncontrada = -1;

  for (let i = dados.length - 1; i >= 1; i--) {
    const sinanPlanilha = formatarValorCelula(dados[i][COL_INICIO_DADOS_LIVRO_VERDE]);
    if (sinanPlanilha === sinanBusca) {
      linhaEncontrada = i;
      break;
    }
  }

  if (linhaEncontrada === -1) {
    return { encontrado: false, sinan: sinanBusca };
  }

  const resultado = montarDadosLivroVerdeDaLinha(dados[linhaEncontrada]);
  return {
    encontrado: true,
    sinan: sinanBusca,
    linha: linhaEncontrada + 1,
    dados: resultado.dados,
    preenchidos: resultado.preenchidos
  };
}

function atualizarRegistroParcial(sheet, numeroLinha, dadosFormulario, profissional, emailAgente) {
  const linhaAtual = sheet.getRange(numeroLinha, 1, 1, sheet.getLastColumn()).getValues()[0];
  let alteracoes = 0;

  CHAVES_CAMPOS_LIVRO_VERDE.forEach(function(chave, indice) {
    const coluna = COL_INICIO_DADOS_LIVRO_VERDE + indice + 1;
    const valorAtual = linhaAtual[COL_INICIO_DADOS_LIVRO_VERDE + indice];
    const valorNovo = dadosFormulario[chave] || '';

    if (celulaVazia(valorAtual) && !celulaVazia(valorNovo)) {
      sheet.getRange(numeroLinha, coluna).setValue(valorNovo);
      alteracoes++;
    }
  });

  const inicioMeses = COL_INICIO_DADOS_LIVRO_VERDE + CHAVES_CAMPOS_LIVRO_VERDE.length;
  const meses = dadosFormulario.meses || [];
  for (let i = 0; i < 12; i++) {
    const coluna = inicioMeses + i + 1;
    const valorAtual = linhaAtual[inicioMeses + i];
    const valorNovo = meses[i] || '';

    if (celulaVazia(valorAtual) && !celulaVazia(valorNovo)) {
      sheet.getRange(numeroLinha, coluna).setValue(valorNovo);
      alteracoes++;
    }
  }

  const inicioEncerramento = inicioMeses + 12;
  CHAVES_ENCERRAMENTO_LIVRO_VERDE.forEach(function(chave, indice) {
    if (chave === 'obs') return;

    const coluna = inicioEncerramento + indice + 1;
    const valorAtual = linhaAtual[inicioEncerramento + indice];
    const valorNovo = dadosFormulario[chave] || '';

    if (celulaVazia(valorAtual) && !celulaVazia(valorNovo)) {
      sheet.getRange(numeroLinha, coluna).setValue(valorNovo);
      alteracoes++;
    }
  });

  const indiceObs = CHAVES_ENCERRAMENTO_LIVRO_VERDE.indexOf('obs');
  const colunaObs = inicioEncerramento + indiceObs + 1;
  const valorNovoObs = dadosFormulario.obs || '';
  const valorAtualObs = linhaAtual[inicioEncerramento + indiceObs];

  if (!celulaVazia(valorNovoObs) && String(valorAtualObs || '').trim() !== String(valorNovoObs).trim()) {
    sheet.getRange(numeroLinha, colunaObs).setValue(valorNovoObs);
    alteracoes++;
  }

  if (alteracoes > 0) {
    atualizarAuditoriaUltimaModificacao(sheet, numeroLinha, profissional, emailAgente);
  }

  return alteracoes;
}

function atualizarAuditoriaUltimaModificacao(sheet, numeroLinha, profissional, emailAgente) {
  sheet.getRange(numeroLinha, 1).setValue(new Date());
  sheet.getRange(numeroLinha, 2).setValue(emailAgente);
  sheet.getRange(numeroLinha, 3).setValue(profissional.nome);
  sheet.getRange(numeroLinha, 4).setValue(profissional.unidade);
}

function contarCamposPreenchidosFormulario(dadosFormulario) {
  let total = 0;

  CHAVES_CAMPOS_LIVRO_VERDE.forEach(function(chave) {
    if (!celulaVazia(dadosFormulario[chave])) total++;
  });

  (dadosFormulario.meses || []).forEach(function(mes) {
    if (!celulaVazia(mes)) total++;
  });

  CHAVES_ENCERRAMENTO_LIVRO_VERDE.forEach(function(chave) {
    if (!celulaVazia(dadosFormulario[chave])) total++;
  });

  return total;
}

function respostaSalvamento(tipo, mensagem, camposPreenchidos) {
  return {
    sucesso: true,
    tipo: tipo,
    mensagem: mensagem,
    camposPreenchidos: camposPreenchidos
  };
}

// 3. Motor de Inserção — colunas 1–4 preenchidas automaticamente no Livro Verde
function salvarRegistro(nomeAbaDestino, dadosFormulario, dadosSessao, numeroLinhaAtualizacao) {
  const ss = obterPlanilha();
  const sheet = ss.getSheetByName(nomeAbaDestino);

  if (!sheet) throw new Error("Aba de destino não encontrada: " + nomeAbaDestino);
  if (!dadosSessao || !dadosSessao.cpf) {
    throw new Error("Sessão inválida. Faça login novamente.");
  }

  const profissional = obterDadosProfissionalPorCpf(dadosSessao.cpf);

  let emailAgente = "";
  try {
    emailAgente = Session.getActiveUser().getEmail();
  } catch (e) {}

  if (numeroLinhaAtualizacao) {
    const trava = LockService.getScriptLock();
    trava.waitLock(10000);

    try {
      const alteracoes = atualizarRegistroParcial(
        sheet,
        numeroLinhaAtualizacao,
        dadosFormulario,
        profissional,
        emailAgente
      );

      if (alteracoes === 0) {
        return {
          sucesso: false,
          tipo: 'atualizado',
          mensagem: 'Nenhum campo foi alterado.',
          camposPreenchidos: 0
        };
      }

      return respostaSalvamento('atualizado', 'Registro atualizado com sucesso', alteracoes);
    } finally {
      trava.releaseLock();
    }
  }

  if (!dadosFormulario.sinan || String(dadosFormulario.sinan).trim() === '') {
    throw new Error('Informe o número do SINAN antes de salvar.');
  }

  const busca = buscarRegistroPorSinan(dadosFormulario.sinan);
  if (busca.encontrado) {
    throw new Error('Este SINAN já está cadastrado. Busque o registro para completar campos em branco.');
  }

  const dataCriacao = new Date();
  const meses = dadosFormulario.meses || [];
  while (meses.length < 12) meses.push('');

  const novaLinha = [
    // ── COLUNAS AUTOMÁTICAS (1–4) ──────────────────────────────
    dataCriacao,              // data_criação
    emailAgente,              // email_agente
    profissional.nome,        // nome_profissional_registro
    profissional.unidade,     // unidade_profissional

    // ── DADOS DO FORMULÁRIO (5 em diante) ──────────────────────
    dadosFormulario.sinan         || '',
    dadosFormulario.prontuario    || '',
    dadosFormulario.nomePaciente  || '',
    dadosFormulario.idade         || '',
    dadosFormulario.sexo          || '',

    // ── DIAGNÓSTICO E EXAMES ───────────────────────────────────
    dadosFormulario.bac1          || '',
    dadosFormulario.bac2          || '',
    dadosFormulario.trm           || '',
    dadosFormulario.cultura       || '',
    dadosFormulario.culturaTS     || '',
    dadosFormulario.rx            || '',
    dadosFormulario.hiv           || '',

    // ── TRATAMENTO ─────────────────────────────────────────────
    dadosFormulario.formaClinica  || '',
    dadosFormulario.tipoEntrada   || '',
    dadosFormulario.esquema       || '',
    dadosFormulario.dataInicio    || '',
    dadosFormulario.tdo           || '',
    dadosFormulario.tarv          || '',

    // ── ACOMPANHAMENTO MENSAL (12 meses) ───────────────────────
    meses[0], meses[1], meses[2],  meses[3],
    meses[4], meses[5], meses[6],  meses[7],
    meses[8], meses[9], meses[10], meses[11],

    // ── ENCERRAMENTO ───────────────────────────────────────────
    dadosFormulario.motivoEnc     || '',
    dadosFormulario.dataEnc       || '',
    dadosFormulario.contIdent     || '',
    dadosFormulario.contExam      || '',
    dadosFormulario.obs           || ''
  ];

  // LockService evita colisão se dois usuários salvarem ao mesmo tempo
  const trava = LockService.getScriptLock();
  trava.waitLock(10000);

  try {
    sheet.appendRow(novaLinha);
  } finally {
    trava.releaseLock();
  }

  return respostaSalvamento(
    'novo',
    'Registro salvo com sucesso',
    contarCamposPreenchidosFormulario(dadosFormulario)
  );
}
