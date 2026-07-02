// =========================================================================
// BACKEND - REGISTROS TUBERCULOSE
// =========================================================================

const NOME_ABA_PROFISSIONAIS = 'Login_Profissional';
const ID_PLANILHA_TB = '1Qonf3_b91auVPxe41EoRMr-x6ZyB3MVWo8aJVyoLCR8';
const COL_INICIO_DADOS = 0; // coluna A — início dos campos do formulário
const QTD_COLS_METADADOS = 4; // data_edicao, email_agente, nome_profissional_registro, unidade_profissional (ao final)

const CONFIG_LIVROS = {
  'Livro Verde - Tratamento': {
    chaveBusca: 'cpf',
    campoSempreEditavel: 'obs',
    campos: [
      'cpf', 'sinan', 'prontuario', 'nomePaciente', 'idade', 'sexo',
      'bac1', 'trm', 'cultura', 'culturaTS', 'rx', 'hiv',
      'formaClinica', 'tipoEntrada', 'esquema', 'dataInicio', 'tdo', 'tarv'
    ],
    meses: 12,
    encerramento: ['motivoEnc', 'dataEnc', 'contIdent', 'contExam', 'obs']
  },
  'Livro Amarelo - ILTB': {
    chaveBusca: 'cpf',
    campoSempreEditavel: null,
    campos: [
      'cpf', 'sinan', 'prontuario', 'notificacao', 'identificacaoIndice', 'nomePaciente', 'idade', 'sexo',
      'baciloscopiaTrm', 'ppd', 'rx', 'hiv', 'esquema', 'dataInicio', 'tdo', 'tarv',
      'motivoEnc', 'dataEnc'
    ],
    meses: 0,
    encerramento: []
  },
  'Livro Azul - Sintomático Respiratório': {
    chaveBusca: 'cpf',
    campoSempreEditavel: null,
    campos: [
      'cpf', 'sinan', 'sequencial', 'entregaAmostra', 'nomeCompleto', 'idade', 'sexo', 'endereco',
      'data1aAmostra', 'resultado1aBaar', 'data2aAmostra', 'resultado2aBaar',
      'dataTrm', 'resultadoTrm', 'numeroGal'
    ],
    meses: 0,
    encerramento: []
  },
  'Livro Amarelo - Contatos': {
    chaveBusca: 'cpf',
    campoSempreEditavel: null,
    campos: [
      'cpf', 'sinanIndice', 'prontuario', 'nomeContato', 'idade', 'sexo', 'grauParentesco',
      'ppdAplicado', 'ppdResultado', 'dataRaioX', 'resultadoRaioX', 'data1aBaar', 'resultado1aBaar',
      'dataAntiHiv', 'resultadoHiv', 'destino', 'data'
    ],
    meses: 0,
    encerramento: []
  }
};

function obterPlanilha() {
  const ativa = SpreadsheetApp.getActiveSpreadsheet();
  if (ativa) return ativa;
  return SpreadsheetApp.openById(ID_PLANILHA_TB);
}

function obterConfigLivro(nomeAba) {
  const config = CONFIG_LIVROS[nomeAba];
  if (!config) throw new Error('Configuração não encontrada para a aba: ' + nomeAba);
  return config;
}

function obterChavesFormulario(config) {
  const chaves = config.campos.slice();
  for (let i = 1; i <= (config.meses || 0); i++) chaves.push('mes' + i);
  (config.encerramento || []).forEach(function(chave) { chaves.push(chave); });
  return chaves;
}

function obterEmailAgente() {
  try {
    const email = Session.getActiveUser().getEmail();
    return email || '';
  } catch (e) {
    return '';
  }
}

function obterQtdColunasFormulario(config) {
  return config.campos.length + (config.meses || 0) + (config.encerramento || []).length;
}

function obterIndiceInicioMetadados(config) {
  return COL_INICIO_DADOS + obterQtdColunasFormulario(config);
}

function gravarMetadadosProfissional(sheet, numeroLinha, profissional, emailAgente, config) {
  const inicio = obterIndiceInicioMetadados(config);
  sheet.getRange(numeroLinha, inicio + 1).setValue(new Date());
  sheet.getRange(numeroLinha, inicio + 2).setValue(emailAgente);
  sheet.getRange(numeroLinha, inicio + 3).setValue(profissional.nome);
  sheet.getRange(numeroLinha, inicio + 4).setValue(profissional.unidade);
}

function montarMetadadosProfissional(profissional, emailAgente) {
  return [new Date(), emailAgente, profissional.nome, profissional.unidade];
}

function inserirNovoRegistro(sheet, dadosFormulario, profissional, emailAgente, config) {
  const valoresForm = montarValoresFormulario(dadosFormulario, config);
  const metadados = montarMetadadosProfissional(profissional, emailAgente);
  const novaLinha = valoresForm.concat(metadados);
  const proximaLinha = sheet.getLastRow() + 1;
  sheet.getRange(proximaLinha, 1, 1, novaLinha.length).setValues([novaLinha]);
}

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Registros Tuberculose - SMS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function validarProfissional(cpfDigitado) {
  const ss = obterPlanilha();
  const sheet = ss.getSheetByName(NOME_ABA_PROFISSIONAIS);
  if (!sheet) throw new Error("Aba '" + NOME_ABA_PROFISSIONAIS + "' não encontrada.");

  const dados = sheet.getDataRange().getValues();
  const cpfLimpo = cpfDigitado.replace(/\D/g, '');

  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]).replace(/\D/g, '') === cpfLimpo) {
      return {
        sucesso: true,
        nome: dados[i][1],
        unidade: dados[i][2],
        funcao: dados[i][3],
        cpf: cpfDigitado
      };
    }
  }
  return { sucesso: false, erro: 'CPF não localizado na base de profissionais.' };
}

function obterDadosProfissionalPorCpf(cpf) {
  const ss = obterPlanilha();
  const sheet = ss.getSheetByName(NOME_ABA_PROFISSIONAIS);
  if (!sheet) throw new Error("Aba '" + NOME_ABA_PROFISSIONAIS + "' não encontrada.");

  const cpfLimpo = String(cpf).replace(/\D/g, '');
  const dados = sheet.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]).replace(/\D/g, '') === cpfLimpo) {
      return { nome: dados[i][1], unidade: dados[i][2] };
    }
  }
  throw new Error('Profissional não encontrado para o CPF informado.');
}

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

function montarDadosDaLinha(linha, config) {
  const dados = {};
  const preenchidos = {};
  let indice = COL_INICIO_DADOS;

  config.campos.forEach(function(chave) {
    const valor = formatarValorCelula(linha[indice++]);
    dados[chave] = valor;
    preenchidos[chave] = valor !== '';
  });

  const meses = [];
  for (let i = 0; i < (config.meses || 0); i++) {
    const valor = formatarValorCelula(linha[indice++]);
    meses.push(valor);
    preenchidos['mes' + (i + 1)] = valor !== '';
  }
  if (config.meses) dados.meses = meses;

  (config.encerramento || []).forEach(function(chave) {
    const valor = formatarValorCelula(linha[indice++]);
    dados[chave] = valor;
    preenchidos[chave] = valor !== '';
  });

  return { dados: dados, preenchidos: preenchidos };
}

function normalizarCpf(cpf) {
  return String(cpf).replace(/\D/g, '');
}

function obterIndiceColunaCampo(config, chave) {
  const indiceCampos = config.campos.indexOf(chave);
  if (indiceCampos >= 0) return COL_INICIO_DADOS + indiceCampos;

  const indiceEncerramento = (config.encerramento || []).indexOf(chave);
  if (indiceEncerramento >= 0) {
    return COL_INICIO_DADOS + config.campos.length + (config.meses || 0) + indiceEncerramento;
  }

  return -1;
}

function obterIndiceColunaBusca(config) {
  return obterIndiceColunaCampo(config, config.chaveBusca);
}

function verificarRegistroEncerrado(linha, config, nomeAba) {
  if (nomeAba !== 'Livro Verde - Tratamento') return false;
  const indiceMotivoEnc = obterIndiceColunaCampo(config, 'motivoEnc');
  if (indiceMotivoEnc < 0) return false;
  return !celulaVazia(linha[indiceMotivoEnc]);
}

function buscarRegistroPorCpf(cpf, nomeAba) {
  const config = obterConfigLivro(nomeAba);
  const ss = obterPlanilha();
  const sheet = ss.getSheetByName(nomeAba);
  if (!sheet) throw new Error("Aba '" + nomeAba + "' não encontrada.");

  const cpfBusca = normalizarCpf(cpf);
  if (!cpfBusca) throw new Error('Informe o CPF para buscar.');

  const dados = sheet.getDataRange().getValues();
  const colCpf = obterIndiceColunaBusca(config);
  let linhaEncontrada = -1;

  for (let i = dados.length - 1; i >= 1; i--) {
    if (normalizarCpf(dados[i][colCpf]) === cpfBusca) {
      linhaEncontrada = i;
      break;
    }
  }

  if (linhaEncontrada === -1) {
    return { encontrado: false, cpf: cpfBusca };
  }

  const resultado = montarDadosDaLinha(dados[linhaEncontrada], config);
  return {
    encontrado: true,
    cpf: cpfBusca,
    linha: linhaEncontrada + 1,
    encerrado: verificarRegistroEncerrado(dados[linhaEncontrada], config, nomeAba),
    dados: resultado.dados,
    preenchidos: resultado.preenchidos
  };
}

function buscarRegistroPorSinan(sinan, nomeAba) {
  return buscarRegistroPorCpf(sinan, nomeAba);
}

function atualizarRegistroParcial(sheet, numeroLinha, dadosFormulario, profissional, emailAgente, config) {
  const linhaAtual = sheet.getRange(numeroLinha, 1, 1, sheet.getLastColumn()).getValues()[0];
  let alteracoes = 0;
  let indice = COL_INICIO_DADOS;

  config.campos.forEach(function(chave) {
    const valorAtual = linhaAtual[indice];
    const valorNovo = dadosFormulario[chave] || '';
    if (celulaVazia(valorAtual) && !celulaVazia(valorNovo)) {
      sheet.getRange(numeroLinha, indice + 1).setValue(valorNovo);
      alteracoes++;
    }
    indice++;
  });

  const meses = dadosFormulario.meses || [];
  for (let i = 0; i < (config.meses || 0); i++) {
    const valorAtual = linhaAtual[indice];
    const valorNovo = meses[i] || '';
    if (celulaVazia(valorAtual) && !celulaVazia(valorNovo)) {
      sheet.getRange(numeroLinha, indice + 1).setValue(valorNovo);
      alteracoes++;
    }
    indice++;
  }

  (config.encerramento || []).forEach(function(chave) {
    const valorAtual = linhaAtual[indice];
    const valorNovo = dadosFormulario[chave] || '';
    const sempreEditavel = config.campoSempreEditavel === chave;

    if (sempreEditavel) {
      if (!celulaVazia(valorNovo) && String(valorAtual || '').trim() !== String(valorNovo).trim()) {
        sheet.getRange(numeroLinha, indice + 1).setValue(valorNovo);
        alteracoes++;
      }
    } else if (celulaVazia(valorAtual) && !celulaVazia(valorNovo)) {
      sheet.getRange(numeroLinha, indice + 1).setValue(valorNovo);
      alteracoes++;
    }
    indice++;
  });

  if (alteracoes > 0) {
    gravarMetadadosProfissional(sheet, numeroLinha, profissional, emailAgente, config);
  }

  return alteracoes;
}

function montarValoresFormulario(dadosFormulario, config) {
  const valores = [];
  config.campos.forEach(function(chave) {
    valores.push(dadosFormulario[chave] || '');
  });

  const meses = dadosFormulario.meses || [];
  for (let i = 0; i < (config.meses || 0); i++) {
    valores.push(meses[i] || '');
  }

  (config.encerramento || []).forEach(function(chave) {
    valores.push(dadosFormulario[chave] || '');
  });

  return valores;
}

function contarCamposPreenchidosFormulario(dadosFormulario, config) {
  let total = 0;
  obterChavesFormulario(config).forEach(function(chave) {
    if (chave.indexOf('mes') === 0 && dadosFormulario.meses) {
      const idx = parseInt(chave.replace('mes', ''), 10) - 1;
      if (!celulaVazia(dadosFormulario.meses[idx])) total++;
      return;
    }
    if (!celulaVazia(dadosFormulario[chave])) total++;
  });
  return total;
}

function respostaSalvamento(tipo, mensagem, camposPreenchidos) {
  return { sucesso: true, tipo: tipo, mensagem: mensagem, camposPreenchidos: camposPreenchidos };
}

function salvarRegistro(nomeAbaDestino, dadosFormulario, dadosSessao, numeroLinhaAtualizacao) {
  const config = obterConfigLivro(nomeAbaDestino);
  const ss = obterPlanilha();
  const sheet = ss.getSheetByName(nomeAbaDestino);

  if (!sheet) throw new Error('Aba de destino não encontrada: ' + nomeAbaDestino);
  if (!dadosSessao || !dadosSessao.cpf) throw new Error('Sessão inválida. Faça login novamente.');

  const profissional = obterDadosProfissionalPorCpf(dadosSessao.cpf);
  const emailAgente = obterEmailAgente();

  if (numeroLinhaAtualizacao) {
    const trava = LockService.getScriptLock();
    trava.waitLock(10000);
    try {
      const alteracoes = atualizarRegistroParcial(
        sheet, numeroLinhaAtualizacao, dadosFormulario, profissional, emailAgente, config
      );
      if (alteracoes === 0) {
        return { sucesso: false, tipo: 'atualizado', mensagem: 'Nenhum campo foi alterado.', camposPreenchidos: 0 };
      }
      return respostaSalvamento('atualizado', 'Registro atualizado com sucesso', alteracoes);
    } finally {
      trava.releaseLock();
    }
  }

  const cpfValor = dadosFormulario[config.chaveBusca];
  if (celulaVazia(cpfValor)) throw new Error('Informe o CPF antes de salvar.');

  const busca = buscarRegistroPorCpf(cpfValor, nomeAbaDestino);
  if (busca.encontrado) {
    const podeNovoComCpfEncerrado = nomeAbaDestino === 'Livro Verde - Tratamento' && busca.encerrado;
    if (!podeNovoComCpfEncerrado) {
      throw new Error('Este CPF já está cadastrado. Busque o registro para completar campos em branco.');
    }
  }

  const trava = LockService.getScriptLock();
  trava.waitLock(10000);
  try {
    inserirNovoRegistro(sheet, dadosFormulario, profissional, emailAgente, config);
  } finally {
    trava.releaseLock();
  }

  return respostaSalvamento('novo', 'Registro salvo com sucesso', contarCamposPreenchidosFormulario(dadosFormulario, config));
}
