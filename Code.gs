const CONFIG = {
  SPREADSHEET_ID: 'COLE_AQUI_O_ID_DA_PLANILHA',
  PDF_FOLDER_ID: 'COLE_AQUI_O_ID_DA_PASTA_DE_PDFS',
  MUNICIPIO_PADRAO: 'Almirante Tamandaré do Sul',
  FUSO_HORARIO: 'America/Sao_Paulo',
  ABA_DECLARACOES: 'Declarações',
  ABA_CONFIGURACAO: 'Configuração',
  ABA_USUARIOS: 'Usuários',
  ABA_HISTORICO_USUARIOS: 'Histórico de Usuários',
  USUARIOS_PROTEGIDOS: [
    'administrador1@exemplo.com',
    'administrador2@exemplo.com'
  ]
};


/**
 * Abre a página principal do aplicativo.
 */
function doGet() {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Declaração Complementar de Rebanho')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/**
 * Permite incluir outros arquivos HTML.
 */
function include(nome) {
  return HtmlService
    .createHtmlOutputFromFile(nome)
    .getContent();
}


/**
 * Retorna os valores apresentados inicialmente no formulário.
 *
 * O número mostrado aqui é apenas uma previsão.
 * O número definitivo é atribuído no momento do salvamento.
 */
function dadosIniciais(token) {
  const agora = new Date();
  const usuario = token ? obterUsuarioPorToken_(token) : null;

  return {
    proximoNumero: proximoNumero_(),
    ano: Number(
      Utilities.formatDate(
        agora,
        CONFIG.FUSO_HORARIO,
        'yyyy'
      )
    ),
    data: Utilities.formatDate(
      agora,
      CONFIG.FUSO_HORARIO,
      'yyyy-MM-dd'
    ),
    municipio: CONFIG.MUNICIPIO_PADRAO,
    usuario: usuario
  };
}


/**
 * Salva uma declaração.
 *
 * Número, ano e data são sempre determinados no servidor.
 */
function salvarDeclaracao(token, dados) {
  const usuarioExecutor = exigirPermissao_(token, 'emitir');
  const bloqueio = LockService.getScriptLock();

  bloqueio.waitLock(30000);

  try {
    const agora = new Date();

    const anoAtual = Number(
      Utilities.formatDate(
        agora,
        CONFIG.FUSO_HORARIO,
        'yyyy'
      )
    );

    const dataFormulario = Utilities.formatDate(
      agora,
      CONFIG.FUSO_HORARIO,
      'yyyy-MM-dd'
    );

    /*
     * O número é consultado somente depois de o bloqueio
     * ser obtido, impedindo números repetidos.
     */
    const numeroAutomatico = proximoNumero_();

    /*
     * Ignora completamente número, ano e data que possam
     * ter sido enviados pelo navegador.
     */
    dados.numero = numeroAutomatico;
    dados.ano = anoAtual;
    dados.data = dataFormulario;
    dados.id = Utilities.getUuid();

    validar_(dados);

    const planilha = SpreadsheetApp.openById(
      CONFIG.SPREADSHEET_ID
    );

    const aba = planilha.getSheetByName(
      CONFIG.ABA_DECLARACOES
    );

    if (!aba) {
      throw new Error(
        'A aba "Declarações" não foi encontrada.'
      );
    }

    const pdf = gerarPdf_(dados);

    const especies = Object.keys(dados.especies || {})
      .filter(function(chave) {
        return dados.especies[chave].selecionada;
      })
      .map(function(chave) {
        return dados.especies[chave].tipo;
      })
      .join(', ');

    const total = Object.values(dados.especies || {})
      .reduce(function(totalGeral, especie) {
        const evolucao = Object.values(
          especie.evolucao || {}
        ).reduce(function(subtotal, valor) {
          return subtotal + (Number(valor) || 0);
        }, 0);

        return totalGeral + evolucao;
      }, 0);

    const boletins = Object.values(dados.especies || {})
      .map(function(especie) {
        return especie.bo;
      })
      .filter(Boolean)
      .join('; ');

    aba.appendRow([
      dados.id,
      dados.numero,
      dados.ano,
      agora,
      dados.nome,
      dados.cpf,
      dados.municipio,
      dados.local,
      especies,
      total,
      boletins,
      pdf.getUrl(),
      agora,
      usuarioExecutor.email,
      JSON.stringify(dados)
    ]);

    /*
     * Atualiza a configuração somente depois que o registro
     * é gravado na planilha.
     */
    atualizarProximoNumero_(
      numeroAutomatico + 1,
      anoAtual
    );

    SpreadsheetApp.flush();

    return {
      ok: true,
      id: dados.id,
      numero: numeroAutomatico,
      ano: anoAtual,
      data: dataFormulario,
      pdfUrl: pdf.getUrl(),
      mensagem:
        'Declaração ' +
        numeroAutomatico +
        '/' +
        anoAtual +
        ' salva com sucesso.'
    };

  } finally {
    bloqueio.releaseLock();
  }
}


/**
 * Consulta declarações já emitidas.
 */
function listarDeclaracoes(token, termo) {
  const planilha = SpreadsheetApp.openById(
    CONFIG.SPREADSHEET_ID
  );

  const aba = planilha.getSheetByName(
    CONFIG.ABA_DECLARACOES
  );

  if (!aba || aba.getLastRow() < 3) {
    return [];
  }

  const quantidadeColunas = Math.max(
    aba.getLastColumn(),
    19
  );

  const valores = aba
    .getRange(
      3,
      1,
      aba.getLastRow() - 2,
      quantidadeColunas
    )
    .getValues();

  const pesquisa = String(termo || '')
    .trim()
    .toLowerCase();

  const usuario = exigirUsuarioAtivo_(token);

  return valores
    .filter(function(linha) {
      if (!pesquisa) {
        return true;
      }

      const conteudo =
        linha[1] + '/' + linha[2] + ' ' +
        linha[4] + ' ' +
        linha[5];

      return conteudo
        .toLowerCase()
        .includes(pesquisa);
    })
    .slice(-100)
    .reverse()
    .map(function(linha) {
      return {
        id: linha[0],
        numero: linha[1],
        ano: linha[2],

        data: linha[3]
          ? Utilities.formatDate(
              new Date(linha[3]),
              CONFIG.FUSO_HORARIO,
              'dd/MM/yyyy'
            )
          : '',

        nome: linha[4],
        cpf: linha[5],
        especies: linha[8],
        total: linha[9],
        pdfUrl: linha[11],

        // Colunas P, Q, R e S.
        status: linha[15] || 'ATIVA',
        motivoCancelamento: linha[16] || '',
        dataCancelamento: linha[17]
          ? Utilities.formatDate(
              new Date(linha[17]),
              CONFIG.FUSO_HORARIO,
              'dd/MM/yyyy HH:mm'
            )
          : '',
        usuarioCancelamento: linha[18] || '',

        podeCancelar: usuario.cancelar,
        podeExcluir: usuario.excluir
      };
    });
}


/**
 * Obtém o próximo número do ano atual.
 *
 * Confere tanto a aba Configuração quanto os registros
 * existentes. Isso acrescenta proteção caso a configuração
 * seja alterada ou não seja atualizada corretamente.
 */
function proximoNumero_() {
  const planilha = SpreadsheetApp.openById(
    CONFIG.SPREADSHEET_ID
  );

  const agora = new Date();

  const anoAtual = Number(
    Utilities.formatDate(
      agora,
      CONFIG.FUSO_HORARIO,
      'yyyy'
    )
  );

  let proximoConfiguracao = 1;
  let maiorNumeroRegistrado = 0;

  const abaConfiguracao = planilha.getSheetByName(
    CONFIG.ABA_CONFIGURACAO
  );

  if (abaConfiguracao) {
    const anoConfigurado = Number(
      abaConfiguracao.getRange('B4').getValue()
    );

    if (anoConfigurado === anoAtual) {
      proximoConfiguracao = Number(
        abaConfiguracao.getRange('B5').getValue()
      ) || 1;
    }
  }

  const abaDeclaracoes = planilha.getSheetByName(
    CONFIG.ABA_DECLARACOES
  );

  if (abaDeclaracoes && abaDeclaracoes.getLastRow() >= 3) {
    const registros = abaDeclaracoes
      .getRange(
        3,
        2,
        abaDeclaracoes.getLastRow() - 2,
        2
      )
      .getValues();

    registros.forEach(function(linha) {
      const numero = Number(linha[0]);
      const ano = Number(linha[1]);

      if (
        ano === anoAtual &&
        numero > maiorNumeroRegistrado
      ) {
        maiorNumeroRegistrado = numero;
      }
    });
  }

  return Math.max(
    proximoConfiguracao,
    maiorNumeroRegistrado + 1
  );
}


/**
 * Salva o próximo número na aba Configuração.
 */
function atualizarProximoNumero_(numero, ano) {
  const planilha = SpreadsheetApp.openById(
    CONFIG.SPREADSHEET_ID
  );

  const aba = planilha.getSheetByName(
    CONFIG.ABA_CONFIGURACAO
  );

  if (!aba) {
    throw new Error(
      'A aba "Configuração" não foi encontrada.'
    );
  }

  aba.getRange('B4').setValue(ano);
  aba.getRange('B5').setValue(numero);
}


/**
 * Valida os dados informados pelo usuário.
 */
function validar_(dados) {
  if (!dados) {
    throw new Error(
      'Os dados da declaração não foram recebidos.'
    );
  }

  if (!dados.nome) {
    throw new Error(
      'Informe o nome do declarante.'
    );
  }

  if (!dados.cpf) {
    throw new Error(
      'Informe o CPF/CNPJ.'
    );
  }
if (!validarCpfCnpjServidor_(dados.cpf)) {
  throw new Error('Informe um CPF ou CNPJ válido.');
}
  if (
    !Object.values(dados.especies || {})
      .some(function(especie) {
        return especie.selecionada;
      })
  ) {
    throw new Error(
      'Selecione ao menos uma espécie.'
    );
  }
}


/**
 * Gera e salva o PDF da declaração.
 */
function gerarPdf_(dados) {
  const modelo = HtmlService.createTemplateFromFile(
    'Print'
  );

  modelo.d = dados;

  const html = modelo
    .evaluate()
    .getContent();

  const nomeSeguro = String(dados.nome || 'Declarante')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/gi, '_');

  const nomeArquivo =
    'Declaracao_Rebanho_' +
    dados.numero +
    '_' +
    dados.ano +
    '_' +
    nomeSeguro +
    '.pdf';

  const arquivoTemporario = Utilities.newBlob(
    html,
    'text/html',
    'declaracao.html'
  );

  const pdf = arquivoTemporario
    .getAs(MimeType.PDF)
    .setName(nomeArquivo);

  const pasta = DriveApp.getFolderById(
    CONFIG.PDF_FOLDER_ID
  );

  return pasta.createFile(pdf);
}
function obterEmailUsuario_() {
  return String(
    Session.getActiveUser().getEmail() || ''
  )
    .trim()
    .toLowerCase();
}


function localizarDeclaracao_(id) {
  const planilha = SpreadsheetApp.openById(
    CONFIG.SPREADSHEET_ID
  );

  const aba = planilha.getSheetByName(
    CONFIG.ABA_DECLARACOES
  );

  if (!aba || aba.getLastRow() < 3) {
    throw new Error(
      'Nenhuma declaração foi encontrada.'
    );
  }

  const ids = aba
    .getRange(
      3,
      1,
      aba.getLastRow() - 2,
      1
    )
    .getValues();

  for (let indice = 0; indice < ids.length; indice++) {
    if (String(ids[indice][0]) === String(id)) {
      return {
        planilha: planilha,
        aba: aba,
        linha: indice + 3
      };
    }
  }

  throw new Error(
    'A declaração solicitada não foi encontrada.'
  );
}


function cancelarDeclaracao(token, id, motivo) {
  const usuarioExecutor = exigirPermissao_(token, 'cancelar');
  motivo = String(motivo || '').trim();

  if (!motivo) {
    throw new Error(
      'Informe o motivo do cancelamento.'
    );
  }

  const bloqueio = LockService.getScriptLock();
  bloqueio.waitLock(30000);

  try {
    const registro = localizarDeclaracao_(id);
    const aba = registro.aba;
    const linha = registro.linha;

    const statusAtual = String(
      aba.getRange(linha, 16).getValue() || 'ATIVA'
    ).toUpperCase();

    if (statusAtual === 'CANCELADA') {
      throw new Error(
        'Esta declaração já está cancelada.'
      );
    }

    const emailUsuario = usuarioExecutor.email;

    aba.getRange(linha, 16, 1, 4).setValues([[
      'CANCELADA',
      motivo,
      new Date(),
      emailUsuario || 'Usuário não identificado'
    ]]);

    SpreadsheetApp.flush();

    const numero = aba.getRange(linha, 2).getValue();
    const ano = aba.getRange(linha, 3).getValue();

    return {
      ok: true,
      mensagem:
        'Declaração ' +
        numero +
        '/' +
        ano +
        ' cancelada com sucesso.'
    };

  } finally {
    bloqueio.releaseLock();
  }
}


function excluirDeclaracao(token, id) {
  exigirPermissao_(token, 'excluir');

  const bloqueio = LockService.getScriptLock();
  bloqueio.waitLock(30000);

  try {
    const registro = localizarDeclaracao_(id);
    const aba = registro.aba;
    const linha = registro.linha;

    const numero = aba.getRange(linha, 2).getValue();
    const ano = aba.getRange(linha, 3).getValue();
    const pdfUrl = String(
      aba.getRange(linha, 12).getValue() || ''
    );

    /*
     * Envia o PDF para a lixeira antes de remover
     * a linha da planilha.
     */
    if (pdfUrl) {
      try {
        const idPdf = extrairIdDrive_(pdfUrl);

        if (idPdf) {
          DriveApp
            .getFileById(idPdf)
            .setTrashed(true);
        }
      } catch (erroPdf) {
        throw new Error(
          'Não foi possível enviar o PDF para a lixeira: ' +
          erroPdf.message
        );
      }
    }

    aba.deleteRow(linha);
    SpreadsheetApp.flush();

    const proximoNumero = recalcularProximoNumero_();

    return {
      ok: true,
      proximoNumero: proximoNumero,
      mensagem:
        'Declaração ' +
        numero +
        '/' +
        ano +
        ' excluída definitivamente. O PDF foi enviado para a lixeira. ' +
        'Próximo número: ' +
        proximoNumero +
        '/' +
        ano +
        '.'
    };

  } finally {
    bloqueio.releaseLock();
  }
}


function extrairIdDrive_(url) {
  const texto = String(url || '');

  const padroes = [
    /\/d\/([a-zA-Z0-9_-]+)/,
    /[?&]id=([a-zA-Z0-9_-]+)/,
    /^([a-zA-Z0-9_-]{20,})$/
  ];

  for (let indice = 0; indice < padroes.length; indice++) {
    const resultado = texto.match(padroes[indice]);

    if (resultado && resultado[1]) {
      return resultado[1];
    }
  }

  return '';
}
function formatarDataPrint_(data) {
  if (!data) {
    return '';
  }

  const partes = String(data).split('-');

  if (partes.length !== 3) {
    return data;
  }

  return partes[2] + '/' + partes[1] + '/' + partes[0];
}


function formatarDataExtensoPrint_(data) {
  if (!data) {
    return '';
  }

  const partes = String(data).split('-');

  if (partes.length !== 3) {
    return data;
  }

  const meses = [
    'janeiro',
    'fevereiro',
    'março',
    'abril',
    'maio',
    'junho',
    'julho',
    'agosto',
    'setembro',
    'outubro',
    'novembro',
    'dezembro'
  ];

  return (
    partes[2] +
    ' de ' +
    meses[Number(partes[1]) - 1] +
    ' de ' +
    partes[0]
  );
}
function validarCpfCnpjServidor_(valor) {
  const numeros = String(valor || '').replace(/\D/g, '');

  if (numeros.length === 11) {
    if (/^(\d)\1{10}$/.test(numeros)) {
      return false;
    }

    let soma = 0;

    for (let i = 0; i < 9; i++) {
      soma += Number(numeros[i]) * (10 - i);
    }

    let digito1 = 11 - (soma % 11);
    digito1 = digito1 >= 10 ? 0 : digito1;

    soma = 0;

    for (let i = 0; i < 10; i++) {
      soma += Number(numeros[i]) * (11 - i);
    }

    let digito2 = 11 - (soma % 11);
    digito2 = digito2 >= 10 ? 0 : digito2;

    return (
      digito1 === Number(numeros[9]) &&
      digito2 === Number(numeros[10])
    );
  }

  if (numeros.length === 14) {
    if (/^(\d)\1{13}$/.test(numeros)) {
      return false;
    }

    function calcular(base, pesos) {
      let soma = 0;

      for (let i = 0; i < pesos.length; i++) {
        soma += Number(base[i]) * pesos[i];
      }

      const resto = soma % 11;
      return resto < 2 ? 0 : 11 - resto;
    }

    const primeiro = calcular(
      numeros.slice(0, 12),
      [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    );

    const segundo = calcular(
      numeros.slice(0, 12) + primeiro,
      [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    );

    return (
      primeiro === Number(numeros[12]) &&
      segundo === Number(numeros[13])
    );
  }

  return false;
}
function recalcularProximoNumero_() {
  const planilha = SpreadsheetApp.openById(
    CONFIG.SPREADSHEET_ID
  );

  const abaDeclaracoes = planilha.getSheetByName(
    CONFIG.ABA_DECLARACOES
  );

  const abaConfiguracao = planilha.getSheetByName(
    CONFIG.ABA_CONFIGURACAO
  );

  if (!abaConfiguracao) {
    throw new Error(
      'A aba "Configuração" não foi encontrada.'
    );
  }

  const anoAtual = Number(
    Utilities.formatDate(
      new Date(),
      CONFIG.FUSO_HORARIO,
      'yyyy'
    )
  );

  let maiorNumero = 0;

  if (
    abaDeclaracoes &&
    abaDeclaracoes.getLastRow() >= 3
  ) {
    const registros = abaDeclaracoes
      .getRange(
        3,
        2,
        abaDeclaracoes.getLastRow() - 2,
        2
      )
      .getValues();

    registros.forEach(function(linha) {
      const numero = Number(linha[0]);
      const ano = Number(linha[1]);

      if (
        ano === anoAtual &&
        numero > maiorNumero
      ) {
        maiorNumero = numero;
      }
    });
  }

  abaConfiguracao.getRange('B4').setValue(anoAtual);
  abaConfiguracao.getRange('B5').setValue(maiorNumero + 1);

  SpreadsheetApp.flush();

  return maiorNumero + 1;
}

function usuarioDaLinha_(r) {
  const email = String(r[2] || '').trim().toLowerCase();
  const ultimoAcesso = r[13]
    ? Utilities.formatDate(new Date(r[13]), CONFIG.FUSO_HORARIO, 'dd/MM/yyyy HH:mm')
    : '';
  return {
    id: String(r[0] || ''), nome: String(r[1] || ''), email: email,
    status: String(r[3] || ''), ativo: String(r[3]).toUpperCase() === 'ATIVO',
    emitir: r[4] === true, cancelar: r[5] === true, excluir: r[6] === true,
    protegido: CONFIG.USUARIOS_PROTEGIDOS.indexOf(email) >= 0,
    ultimoAcesso: ultimoAcesso
  };
}

function localizarUsuarioPorEmail_(email) {
  email = String(email || '').trim().toLowerCase();
  const sh = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.ABA_USUARIOS);
  if (!sh || sh.getLastRow() < 3) return null;
  const valores = sh.getRange(3, 1, sh.getLastRow() - 2, 14).getValues();
  for (let i = 0; i < valores.length; i++) {
    if (String(valores[i][2]).trim().toLowerCase() === email) {
      return {aba: sh, linha: i + 3, valores: valores[i], usuario: usuarioDaLinha_(valores[i])};
    }
  }
  return null;
}

function hashSenha_(senha, salt) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(salt) + String(senha),
    Utilities.Charset.UTF_8
  );
  return bytes.map(function(b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
}

function validarSenha_(senha) {
  if (String(senha || '').length < 8) throw new Error('A senha deve possuir pelo menos 8 caracteres.');
}

function criarSessao_(email) {
  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('sessao_' + token, email, 21600);
  return token;
}

function obterUsuarioPorToken_(token) {
  token = String(token || '').trim();
  if (!token) return null;
  const email = CacheService.getScriptCache().get('sessao_' + token);
  if (!email) return null;
  const registro = localizarUsuarioPorEmail_(email);
  return registro ? registro.usuario : null;
}

function exigirUsuarioAtivo_(token) {
  const usuario = obterUsuarioPorToken_(token);
  if (!usuario || !usuario.ativo) throw new Error('Sua sessão expirou ou seu usuário está inativo. Entre novamente.');
  return usuario;
}

function exigirPermissao_(token, permissao) {
  const usuario = exigirUsuarioAtivo_(token);
  if (!usuario[permissao]) throw new Error('Seu usuário não possui permissão para esta operação.');
  return usuario;
}

function login(email, senha) {
  email = String(email || '').trim().toLowerCase();
  const registro = localizarUsuarioPorEmail_(email);
  if (!registro || !registro.valores[11] || hashSenha_(senha, registro.valores[12]) !== registro.valores[11]) {
    throw new Error('E-mail ou senha inválidos.');
  }
  if (!registro.usuario.ativo) throw new Error('Seu cadastro ainda não foi aprovado ou está inativo.');
  const token = criarSessao_(email);
  registro.aba.getRange(registro.linha, 14).setValue(new Date());
  registrarHistoricoUsuario_('ACESSO', registro.valores, registro.valores, email, 'Login realizado');
  return {ok: true, token: token, usuario: registro.usuario};
}

function logout(token) {
  const usuario = obterUsuarioPorToken_(token);
  CacheService.getScriptCache().remove('sessao_' + String(token || ''));
  if (usuario) registrarHistoricoUsuario_('SAÍDA', null, [usuario.id, usuario.nome, usuario.email], usuario.email, 'Logout realizado');
  return {ok: true};
}

function listarUsuarios(token) {
  exigirPermissao_(token, 'excluir');
  const sh = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.ABA_USUARIOS);
  if (!sh || sh.getLastRow() < 3) return [];
  return sh.getRange(3, 1, sh.getLastRow() - 2, 14).getValues()
    .filter(function(r) { return String(r[0] || '').trim() !== ''; })
    .map(usuarioDaLinha_);
}

function proximaLinhaLivreUsuario_(sh) {
  const ultimaLinha = Math.max(sh.getLastRow(), 3);
  const ids = sh.getRange(3, 1, ultimaLinha - 2, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (!String(ids[i][0] || '').trim()) return i + 3;
  }
  return ultimaLinha + 1;
}

function concederAcessoPastaPdf_(email) {
  email = String(email || '').trim().toLowerCase();
  if (!email || CONFIG.USUARIOS_PROTEGIDOS.indexOf(email) >= 0) return;
  try {
    DriveApp.getFolderById(CONFIG.PDF_FOLDER_ID).addViewer(email);
  } catch (erro) {
    throw new Error('Não foi possível compartilhar a pasta de PDFs com ' + email + '. Verifique se o e-mail possui uma Conta Google. Detalhes: ' + erro.message);
  }
}

function removerAcessoPastaPdf_(email) {
  email = String(email || '').trim().toLowerCase();
  if (!email || CONFIG.USUARIOS_PROTEGIDOS.indexOf(email) >= 0) return;
  try {
    DriveApp.getFolderById(CONFIG.PDF_FOLDER_ID).removeViewer(email);
  } catch (erro) {
    throw new Error('Não foi possível remover o acesso de ' + email + ' à pasta de PDFs. Detalhes: ' + erro.message);
  }
}

function salvarUsuario(token, dados) {
  const executor = exigirPermissao_(token, 'excluir');
  const nome = String(dados.nome || '').trim().toLocaleUpperCase('pt-BR');
  const email = String(dados.email || '').trim().toLowerCase();
  if (!nome) throw new Error('Informe o nome do usuário.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Informe um e-mail válido.');
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sh = ss.getSheetByName(CONFIG.ABA_USUARIOS);
    const valores = sh.getLastRow() >= 3 ? sh.getRange(3,1,sh.getLastRow()-2,14).getValues() : [];
    let linha = -1, anterior = null;
    for (let i=0;i<valores.length;i++) {
      if (String(valores[i][0]) === String(dados.id || '') || String(valores[i][2]).trim().toLowerCase() === email) {
        linha = i + 3; anterior = valores[i]; break;
      }
    }
    const agora = new Date();
    const id = linha > 0 ? String(anterior[0]) : Utilities.getUuid();
    const emailAnterior = linha > 0 ? String(anterior[2]).trim().toLowerCase() : '';
    const protegido = CONFIG.USUARIOS_PROTEGIDOS.indexOf(email) >= 0 || CONFIG.USUARIOS_PROTEGIDOS.indexOf(emailAnterior) >= 0;
    if (protegido && linha > 0 && email !== emailAnterior) throw new Error('O e-mail deste usuário institucional é protegido e não pode ser alterado.');
    const statusSolicitado = protegido ? 'ATIVO' : (dados.status === 'INATIVO' ? 'INATIVO' : (dados.status === 'PENDENTE' ? 'PENDENTE' : 'ATIVO'));
    if (statusSolicitado === 'ATIVO') concederAcessoPastaPdf_(email);
    if (emailAnterior && emailAnterior !== email) removerAcessoPastaPdf_(emailAnterior);
    if (statusSolicitado !== 'ATIVO' && emailAnterior) removerAcessoPastaPdf_(emailAnterior);
    const novo = [id,nome,email,statusSolicitado,dados.emitir !== false,dados.cancelar !== false,dados.excluir !== false,
      linha > 0 ? anterior[7] : agora, linha > 0 ? anterior[8] : executor.email, agora, executor.email,
      linha > 0 ? anterior[11] : '', linha > 0 ? anterior[12] : '', linha > 0 ? anterior[13] : ''];
    if (linha > 0) {
      sh.getRange(linha,1,1,14).setValues([novo]);
    } else {
      sh.getRange(proximaLinhaLivreUsuario_(sh),1,1,14).setValues([novo]);
    }
    registrarHistoricoUsuario_(linha > 0 ? 'ALTERAÇÃO' : 'CADASTRO', anterior, novo, executor.email, 'Cadastro atualizado pelo aplicativo');
    SpreadsheetApp.flush();
    return {ok:true,mensagem:'Usuário salvo com sucesso.'};
  } finally { lock.releaseLock(); }
}

function excluirUsuario(token, id, confirmacao) {
  const executor = exigirPermissao_(token, 'excluir');
  if (String(confirmacao || '').trim().toUpperCase() !== 'EXCLUIR') throw new Error('Confirmação de exclusão inválida.');
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const sh = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.ABA_USUARIOS);
    const valores = sh.getRange(3,1,Math.max(sh.getLastRow()-2,1),14).getValues();
    let linha = -1, anterior = null;
    for (let i=0;i<valores.length;i++) if (String(valores[i][0]) === String(id)) {linha=i+3; anterior=valores[i]; break;}
    if (linha < 0) throw new Error('Usuário não encontrado.');
    if (String(anterior[2]).trim().toLowerCase() === executor.email) throw new Error('Você não pode excluir o próprio usuário.');
    if (CONFIG.USUARIOS_PROTEGIDOS.indexOf(String(anterior[2]).trim().toLowerCase()) >= 0) {
      throw new Error('Este usuário institucional é protegido e não pode ser excluído.');
    }
    removerAcessoPastaPdf_(String(anterior[2]).trim().toLowerCase());
    registrarHistoricoUsuario_('EXCLUSÃO', anterior, null, executor.email, 'Usuário removido do cadastro');
    sh.deleteRow(linha); SpreadsheetApp.flush();
    return {ok:true,mensagem:'Usuário excluído e operação registrada no histórico.'};
  } finally { lock.releaseLock(); }
}

function registrarHistoricoUsuario_(acao, anterior, novo, executor, detalhes) {
  const sh = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.ABA_HISTORICO_USUARIOS);
  const base = novo || anterior || [];
  sh.appendRow([
    new Date(), acao, base[0] || '', base[1] || '', base[2] || '',
    anterior ? JSON.stringify(sanitizarUsuarioHistorico_(anterior)) : '',
    novo ? JSON.stringify(sanitizarUsuarioHistorico_(novo)) : '',
    executor, detalhes || ''
  ]);
}

function sanitizarUsuarioHistorico_(linha) {
  if (!Array.isArray(linha)) return linha;
  return {
    id: linha[0] || '', nome: linha[1] || '', email: linha[2] || '',
    status: linha[3] || '', emitir: linha[4] === true,
    cancelar: linha[5] === true, excluir: linha[6] === true,
    cadastradoEm: linha[7] || '', cadastradoPor: linha[8] || '',
    atualizadoEm: linha[9] || '', atualizadoPor: linha[10] || '',
    senhaConfigurada: Boolean(linha[11]),
    ultimoAcesso: linha[13] || ''
  };
}

function limparCredenciaisDoHistorico() {
  const sh = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.ABA_HISTORICO_USUARIOS);
  if (!sh || sh.getLastRow() < 3) return;
  const faixa = sh.getRange(3, 6, sh.getLastRow() - 2, 2);
  const valores = faixa.getValues();
  valores.forEach(function(linha) {
    for (let i = 0; i < 2; i++) {
      if (!linha[i]) continue;
      try {
        const conteudo = JSON.parse(linha[i]);
        if (Array.isArray(conteudo)) linha[i] = JSON.stringify(sanitizarUsuarioHistorico_(conteudo));
      } catch (erro) {
        // Mantém registros antigos que não estejam em JSON.
      }
    }
  });
  faixa.setValues(valores);
  SpreadsheetApp.flush();
}

function solicitarCadastro(dados) {
  dados = dados || {};
  const email = String(dados.email || '').trim().toLowerCase();
  const nome = String(dados.nome || '').trim().toLocaleUpperCase('pt-BR');
  const senha = String(dados.senha || '');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Informe um e-mail válido.');
  if (!nome) throw new Error('Informe seu nome completo.');
  validarSenha_(senha);
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const sh = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.ABA_USUARIOS);
    const valores = sh.getLastRow() >= 3 ? sh.getRange(3,1,sh.getLastRow()-2,14).getValues() : [];
    for (let i=0;i<valores.length;i++) {
      if (String(valores[i][2]).trim().toLowerCase() === email) {
        const status = String(valores[i][3]).toUpperCase();
        if (status === 'ATIVO') throw new Error('Seu usuário já está ativo.');
        throw new Error('Já existe uma solicitação para este e-mail com status ' + status + '.');
      }
    }
    const agora = new Date();
    const salt = Utilities.getUuid().replace(/-/g, '');
    const novo = [Utilities.getUuid(),nome,email,'PENDENTE',false,false,false,agora,email,agora,email,hashSenha_(senha,salt),salt,''];
    sh.getRange(proximaLinhaLivreUsuario_(sh),1,1,14).setValues([novo]);
    registrarHistoricoUsuario_('SOLICITAÇÃO DE CADASTRO', null, novo, email, 'Cadastro solicitado na tela inicial');
    SpreadsheetApp.flush();
    return {ok:true,mensagem:'Cadastro solicitado. Aguarde a liberação no Gerencial.'};
  } finally { lock.releaseLock(); }
}

function gerarCodigosPrimeiroAcesso() {
  const propriedades = PropertiesService.getScriptProperties();
  const validade = Date.now() + 20 * 60 * 1000;
  CONFIG.USUARIOS_PROTEGIDOS.forEach(function(email) {
    const registro = localizarUsuarioPorEmail_(email);
    if (!registro) throw new Error('Usuário protegido não encontrado: ' + email);
    if (registro.valores[11]) {
      console.log(email + ': senha já definida');
      return;
    }
    const codigo = String(Math.floor(100000 + Math.random() * 900000));
    propriedades.setProperty('PRIMEIRO_ACESSO_' + email, JSON.stringify({codigo: codigo, validade: validade}));
    console.log(email + ': ' + codigo + ' (válido por 20 minutos)');
  });
}

function solicitarCodigoPrimeiroAcesso(email) {
  email = String(email || '').trim().toLowerCase();
  const resposta = {
    ok: true,
    mensagem: 'Se o e-mail estiver autorizado e ainda não possuir senha, o código será enviado.'
  };
  if (CONFIG.USUARIOS_PROTEGIDOS.indexOf(email) < 0) return resposta;

  const cache = CacheService.getScriptCache();
  const chaveLimite = 'limite_primeiro_' + hashSenha_(email, 'limite');
  if (cache.get(chaveLimite)) return resposta;
  cache.put(chaveLimite, '1', 60);

  const registro = localizarUsuarioPorEmail_(email);
  if (!registro || registro.valores[11]) return resposta;

  const codigo = String(Math.floor(100000 + Math.random() * 900000));
  PropertiesService.getScriptProperties().setProperty(
    'PRIMEIRO_ACESSO_' + email,
    JSON.stringify({codigo: codigo, validade: Date.now() + 20 * 60 * 1000})
  );

  MailApp.sendEmail({
    to: email,
    subject: 'Código de primeiro acesso',
    name: 'Declaração de Rebanho',
    body:
      'Olá, ' + registro.usuario.nome + '.\n\n' +
      'Seu código de primeiro acesso é: ' + codigo + '\n\n' +
      'O código é válido por 20 minutos.\n\n' +
      'Sistema de Declaração Complementar de Rebanho'
  });
  registrarHistoricoUsuario_('CÓDIGO DE PRIMEIRO ACESSO', registro.valores, registro.valores, email, 'Código enviado ao e-mail cadastrado');
  return resposta;
}

function criarSenhaInicial(email, codigo, senha) {
  email = String(email || '').trim().toLowerCase();
  if (CONFIG.USUARIOS_PROTEGIDOS.indexOf(email) < 0) throw new Error('Primeiro acesso não autorizado para este e-mail.');
  validarSenha_(senha);
  const chave = 'PRIMEIRO_ACESSO_' + email;
  const propriedades = PropertiesService.getScriptProperties();
  const texto = propriedades.getProperty(chave);
  if (!texto) throw new Error('Gere um novo código de primeiro acesso no editor do Apps Script.');
  const acesso = JSON.parse(texto);
  if (Date.now() > Number(acesso.validade) || String(codigo || '').trim() !== String(acesso.codigo)) throw new Error('Código inválido ou expirado.');
  const registro = localizarUsuarioPorEmail_(email);
  if (!registro) throw new Error('Usuário não encontrado.');
  const salt = Utilities.getUuid().replace(/-/g, '');
  registro.aba.getRange(registro.linha, 12, 1, 3).setValues([[hashSenha_(senha,salt), salt, '']]);
  propriedades.deleteProperty(chave);
  registrarHistoricoUsuario_('PRIMEIRO ACESSO', registro.valores, registro.valores, email, 'Senha inicial definida');
  return {ok:true,mensagem:'Senha criada. Agora você já pode entrar no sistema.'};
}

function solicitarRedefinicaoSenha(email) {
  email = String(email || '').trim().toLowerCase();
  const resposta = {
    ok: true,
    mensagem: 'Se o e-mail estiver cadastrado e ativo, você receberá um código para redefinir a senha.'
  };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return resposta;

  const cache = CacheService.getScriptCache();
  const chaveLimite = 'limite_reset_' + hashSenha_(email, 'limite');
  if (cache.get(chaveLimite)) return resposta;
  cache.put(chaveLimite, '1', 60);

  const registro = localizarUsuarioPorEmail_(email);
  if (!registro || !registro.usuario.ativo || !registro.valores[11]) return resposta;

  const codigo = String(Math.floor(100000 + Math.random() * 900000));
  const saltCodigo = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty(
    'RESET_SENHA_' + email,
    JSON.stringify({
      hash: hashSenha_(codigo, saltCodigo),
      salt: saltCodigo,
      validade: Date.now() + 15 * 60 * 1000,
      tentativas: 0
    })
  );

  MailApp.sendEmail({
    to: email,
    subject: 'Código para redefinir sua senha',
    name: 'Declaração de Rebanho',
    body:
      'Olá, ' + registro.usuario.nome + '.\n\n' +
      'Seu código para redefinir a senha é: ' + codigo + '\n\n' +
      'O código é válido por 15 minutos. Se você não solicitou esta alteração, ignore esta mensagem.\n\n' +
      'Sistema de Declaração Complementar de Rebanho'
  });
  registrarHistoricoUsuario_('SOLICITAÇÃO DE NOVA SENHA', registro.valores, registro.valores, email, 'Código enviado ao e-mail cadastrado');
  return resposta;
}

function confirmarRedefinicaoSenha(email, codigo, novaSenha) {
  email = String(email || '').trim().toLowerCase();
  codigo = String(codigo || '').trim();
  validarSenha_(novaSenha);
  const chave = 'RESET_SENHA_' + email;
  const propriedades = PropertiesService.getScriptProperties();
  const texto = propriedades.getProperty(chave);
  if (!texto) throw new Error('Código inválido ou expirado. Solicite um novo código.');

  const redefinicao = JSON.parse(texto);
  if (Date.now() > Number(redefinicao.validade)) {
    propriedades.deleteProperty(chave);
    throw new Error('Código expirado. Solicite um novo código.');
  }
  redefinicao.tentativas = Number(redefinicao.tentativas || 0) + 1;
  if (redefinicao.tentativas > 5) {
    propriedades.deleteProperty(chave);
    throw new Error('Número máximo de tentativas excedido. Solicite outro código.');
  }
  if (hashSenha_(codigo, redefinicao.salt) !== redefinicao.hash) {
    propriedades.setProperty(chave, JSON.stringify(redefinicao));
    throw new Error('Código inválido.');
  }

  const registro = localizarUsuarioPorEmail_(email);
  if (!registro || !registro.usuario.ativo) throw new Error('Não foi possível redefinir a senha.');
  const saltSenha = Utilities.getUuid().replace(/-/g, '');
  registro.aba.getRange(registro.linha, 12, 1, 2).setValues([[
    hashSenha_(novaSenha, saltSenha), saltSenha
  ]]);
  propriedades.deleteProperty(chave);
  registrarHistoricoUsuario_('REDEFINIÇÃO DE SENHA', registro.valores, registro.valores, email, 'Senha redefinida pelo próprio usuário');
  return {ok:true,mensagem:'Senha redefinida com sucesso. Você já pode entrar.'};
}
