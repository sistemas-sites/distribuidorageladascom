const express = require('express');
const bcrypt = require('bcrypt');
const fs = require('fs');
const { Parser } = require('json2csv');
const path = require('path');
const ExcelJS = require('exceljs');
const mysql = require('mysql'); // Importa a versão mysql
const cors = require('cors');
const bodyParser = require('body-parser');
require('dotenv').config();
const session = require('express-session');
const saltRounds = 10;
const app = express();
app.use(cors());
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));


app.use(session({
    secret: 'seuSegredoAqui',
    resave: false,
    saveUninitialized: true,
    cookie: { secure: false, maxAge: 30 * 60 * 1000 } // Em produção, mude para true e use HTTPS
}));

const port = 3001;

// Configuração de conexão com o MySQL
const db = mysql.createPool({
    host: process.env.DB_HOST,
    user: process.env.DB_USER,
    password: process.env.DB_PASSWORD,
    database: process.env.DB_DATABASE,
    port: process.env.DB_PORT || 3306,
    // Datas voltam como string 'YYYY-MM-DD' para não sofrerem conversão de fuso
    dateStrings: ['DATE']
});

// Função de consulta que retorna uma Promise
function query(sql, params) {
    return new Promise((resolve, reject) => {
        db.query(sql, params, (err, result) => {
            if (err) return reject(err);
            resolve(result);
        });
    });
}

/* =======================================================================
   REGRAS DE JORNADA
   -----------------------------------------------------------------------
   A jornada é noturna e virou o dia: começa por volta das 17:00 e termina
   por volta das 02:00 do dia seguinte. Por isso todo horário posterior à
   entrada que aparecer "menor" que ela recebe +24h antes de qualquer conta.
   ======================================================================= */

const JORNADA_DIARIA_MIN = 8 * 60;        // 8h por dia trabalhado
const JANTA_PADRAO_MIN = 30;              // intervalo de janta de 30 minutos
const NOTURNO_INICIO_MIN = 22 * 60;       // adicional noturno a partir das 22:00
const NOTURNO_FIM_MIN = 29 * 60;          // ... até as 05:00 do dia seguinte
const META_MENSAL_HORAS = 176;            // carga horária mensal do funcionário
const TIPOS_DIA_VALIDOS = ['FOLGA', 'FALTA', 'ATESTADO', 'FERIADO', 'FERIAS', 'NORMAL'];

// Garante a coluna que marca manualmente o tipo do dia (folga, atestado, etc.)
async function garantirColunaTipoDia() {
    try {
        const [existe] = await query(
            `SELECT COUNT(*) AS total
               FROM information_schema.COLUMNS
              WHERE TABLE_SCHEMA = ? AND TABLE_NAME = 'pontos' AND COLUMN_NAME = 'tipo_dia'`,
            [process.env.DB_DATABASE]
        );
        if (existe && existe.total === 0) {
            await query('ALTER TABLE pontos ADD COLUMN tipo_dia VARCHAR(20) DEFAULT NULL');
            console.log('✅ Coluna pontos.tipo_dia criada');
        }
    } catch (err) {
        console.error('⚠️ Não foi possível verificar/criar a coluna tipo_dia:', err.message);
    }
}
garantirColunaTipoDia();

// 'HH:MM[:SS]' -> minutos desde a meia-noite (com fração de segundos)
function horaParaMinutos(hora) {
    if (hora === null || hora === undefined || hora === '' || hora === 'N/A') return null;
    const partes = String(hora).split(':').map(Number);
    const [h, m, s] = partes;
    if (!Number.isFinite(h) || !Number.isFinite(m)) return null;
    return h * 60 + m + (Number.isFinite(s) ? s / 60 : 0);
}

function minutosParaHoraTexto(minutos) {
    if (minutos === null || minutos === undefined) return '--';
    const sinal = minutos < 0 ? '-' : '';
    const total = Math.round(Math.abs(minutos));
    return `${sinal}${Math.floor(total / 60)}h ${String(total % 60).padStart(2, '0')}m`;
}

// Minutos do intervalo [inicio, fim] que caem na faixa noturna (22:00 às 05:00).
// A faixa é testada com deslocamento de -24h/0/+24h porque a jornada atravessa dias.
function minutosNaFaixaNoturna(inicio, fim) {
    if (inicio === null || fim === null || fim <= inicio) return 0;
    let total = 0;
    for (const deslocamento of [-1440, 0, 1440]) {
        const faixaInicio = NOTURNO_INICIO_MIN + deslocamento;
        const faixaFim = NOTURNO_FIM_MIN + deslocamento;
        total += Math.max(0, Math.min(fim, faixaFim) - Math.max(inicio, faixaInicio));
    }
    return total;
}

/**
 * Calcula um dia de trabalho já tratando a virada de meia-noite.
 * Retorna minutos trabalhados (líquidos), minutos de intervalo e minutos
 * de adicional noturno (que NÃO entram no total trabalhado — vão para a
 * tabela de adicional noturno).
 */
function calcularDiaTrabalhado(ponto) {
    const entrada = horaParaMinutos(ponto.entrada);
    const saidaBruta = horaParaMinutos(ponto.saida);

    if (entrada === null || saidaBruta === null) {
        return {
            minutosTrabalhados: 0,
            minutosIntervalo: 0,
            minutosNoturnos: 0,
            inconsistente: true,
            observacao: entrada === null ? 'Entrada não registrada' : 'Saída não registrada'
        };
    }

    // Saída no dia seguinte (ex.: entrada 17:00, saída 02:00)
    let saida = saidaBruta;
    while (saida < entrada) saida += 1440;

    const minutosBrutos = saida - entrada;

    // Intervalo de janta: usa o registrado; sem registro, desconta os 30 min padrão
    let jantaInicio = horaParaMinutos(ponto.saida_almoco);
    let jantaFim = horaParaMinutos(ponto.volta_almoco);
    let minutosIntervalo;
    let observacao = '';

    if (jantaInicio !== null && jantaFim !== null) {
        while (jantaInicio < entrada) jantaInicio += 1440;
        while (jantaFim < jantaInicio) jantaFim += 1440;

        if (jantaFim > saida) {
            jantaInicio = null;
            jantaFim = null;
            minutosIntervalo = JANTA_PADRAO_MIN;
            observacao = 'Janta fora da jornada — descontados 30 min padrão';
        } else {
            minutosIntervalo = jantaFim - jantaInicio;
        }
    } else {
        jantaInicio = null;
        jantaFim = null;
        minutosIntervalo = JANTA_PADRAO_MIN;
        observacao = 'Janta não registrada — descontados 30 min padrão';
    }

    const minutosTrabalhados = Math.max(0, minutosBrutos - minutosIntervalo);

    // Adicional noturno: faixa noturna da jornada menos a faixa noturna da janta
    let minutosNoturnos = minutosNaFaixaNoturna(entrada, saida);
    if (jantaInicio !== null) {
        minutosNoturnos -= minutosNaFaixaNoturna(jantaInicio, jantaFim);
    }
    minutosNoturnos = Math.max(0, Math.min(minutosNoturnos, minutosTrabalhados));

    return {
        minutosTrabalhados,
        minutosIntervalo,
        minutosNoturnos,
        inconsistente: !!observacao,
        observacao
    };
}

// --- Helpers de data (sem passar por fuso horário) ---
function parseDataISO(texto) {
    const [ano, mes, dia] = String(texto).split('-').map(Number);
    return new Date(ano, mes - 1, dia);
}

function formatarDataISO(data) {
    return `${data.getFullYear()}-${String(data.getMonth() + 1).padStart(2, '0')}-${String(data.getDate()).padStart(2, '0')}`;
}

function formatarDataBR(texto) {
    const [ano, mes, dia] = String(texto).split('-');
    return `${dia}/${mes}/${ano}`;
}

const NOMES_DIA_SEMANA = ['DOMINGO', 'SEGUNDA-FEIRA', 'TERCA-FEIRA', 'QUARTA-FEIRA', 'QUINTA-FEIRA', 'SEXTA-FEIRA', 'SABADO'];

// Remove acentos/pontuação para comparar 'terça-feir' (truncado no banco) com 'TERCA-FEIRA'
function normalizarTexto(texto) {
    return String(texto || '')
        .normalize('NFD')
        .replace(/[̀-ͯ]/g, '')
        .toUpperCase()
        .replace(/[^A-Z]/g, '');
}

function ehDiaDeFolgaFixa(diaSemanaNome, diaFolgaCadastrado) {
    const folga = normalizarTexto(diaFolgaCadastrado);
    if (folga.length < 3) return false;
    const dia = normalizarTexto(diaSemanaNome);
    return dia.startsWith(folga) || folga.startsWith(dia);
}

// Rotas
app.get('/', verificarAutenticacao, async (req, res) => {
    try {
        const sql = 'SELECT id, nome FROM funcionarios'
        console.log("sql", sql)
        const result = await query(sql);
        console.log("result", result);
        res.render('index', { funcionarios: result });
    } catch (err) {
        console.error('Erro ao buscar funcionários:', err);
        return res.status(500).send('Erro ao buscar funcionários');
    }
});

const adjustToBrasiliaTime = (date) => {
    // Converte o horário UTC para o horário de Brasília (GMT-3)
    const brDate = new Date(date.getTime() - 3 * 3600000);
    return brDate;
};

app.post('/ponto/entrada', async (req, res) => {
    const { funcionario_id } = req.body;
    const dataAtual = adjustToBrasiliaTime(new Date());
    const horaEntrada = dataAtual.toTimeString().slice(0, 8);

    // A jornada começa por volta das 17:00 e termina no dia seguinte. O
    // registro sempre pertence ao dia em que a jornada COMEÇOU, então só
    // vira para o dia seguinte quem entra depois da meia-noite.
    const dataRegistro = new Date(dataAtual.getTime());
    const dataRegistroFormatada = formatarDataISO(dataRegistro);

    try {
        // Bloqueia entrada duplicada e também entrada com jornada ainda aberta
        const abertos = await query(
            `SELECT id, data FROM pontos
              WHERE funcionario_id = ? AND (data = ? OR (saida IS NULL AND entrada IS NOT NULL))
              ORDER BY data DESC`,
            [funcionario_id, dataRegistroFormatada]
        );

        if (abertos.some(p => String(p.data) === dataRegistroFormatada && p.entrada !== null)) {
            return res.status(400).json({ message: 'Já existe uma entrada registrada para hoje.' });
        }
        if (abertos.some(p => p.saida === null && p.entrada !== null)) {
            return res.status(400).json({ message: 'Existe uma jornada anterior sem saída registrada. Procure o supervisor.' });
        }

        // O dia pode já existir marcado como folga/atestado: nesse caso completa o registro
        const existente = abertos.find(p => String(p.data) === dataRegistroFormatada);
        if (existente) {
            await query('UPDATE pontos SET entrada = ? WHERE id = ?', [horaEntrada, existente.id]);
        } else {
            await query(
                'INSERT INTO pontos (funcionario_id, entrada, data) VALUES (?, ?, ?)',
                [funcionario_id, horaEntrada, dataRegistroFormatada]
            );
        }

        return res.status(200).json({ message: 'Entrada registrada com sucesso!' });
    } catch (err) {
        console.error('Erro ao registrar entrada:', err);
        return res.status(500).json({ error: 'Erro ao registrar entrada' });
    }
});

// Busca a jornada aberta (entrada batida, saída ainda não) — indispensável
// para a janta e a saída que acontecem depois da meia-noite.
async function buscarJornadaAberta(funcionario_id) {
    const [aberta] = await query(
        `SELECT id, data, entrada, saida_almoco, volta_almoco
           FROM pontos
          WHERE funcionario_id = ? AND entrada IS NOT NULL AND saida IS NULL
          ORDER BY data DESC
          LIMIT 1`,
        [funcionario_id]
    );
    return aberta;
}

app.post('/ponto/saida-almoco', async (req, res) => {
    const { funcionario_id } = req.body;
    const horaSaidaAlmoco = adjustToBrasiliaTime(new Date()).toTimeString().slice(0, 8);

    try {
        const jornada = await buscarJornadaAberta(funcionario_id);

        if (!jornada) {
            return res.status(400).json({ message: 'Nenhuma jornada aberta. Bata a entrada primeiro.' });
        }
        if (jornada.saida_almoco) {
            return res.status(400).json({ message: 'Já existe uma saída para a janta registrada nesta jornada.' });
        }

        await query('UPDATE pontos SET saida_almoco = ? WHERE id = ?', [horaSaidaAlmoco, jornada.id]);
        return res.status(200).json({ message: 'Ponto de saída para a janta registrado com sucesso!' });
    } catch (err) {
        console.error('Erro ao registrar saída para a janta:', err);
        return res.status(500).json({ error: 'Erro ao registrar saída para a janta' });
    }
});

app.post('/ponto/volta-almoco', async (req, res) => {
    const { funcionario_id } = req.body;
    const horaVoltaAlmoco = adjustToBrasiliaTime(new Date()).toTimeString().slice(0, 8);

    try {
        const jornada = await buscarJornadaAberta(funcionario_id);

        if (!jornada) {
            return res.status(400).json({ message: 'Nenhuma jornada aberta. Bata a entrada primeiro.' });
        }
        if (jornada.volta_almoco) {
            return res.status(400).json({ message: 'Já existe uma volta da janta registrada nesta jornada.' });
        }

        await query('UPDATE pontos SET volta_almoco = ? WHERE id = ?', [horaVoltaAlmoco, jornada.id]);
        return res.status(200).json({ message: 'Ponto de volta da janta registrado com sucesso!' });
    } catch (err) {
        console.error('Erro ao registrar volta da janta:', err);
        return res.status(500).json({ error: 'Erro ao registrar volta da janta' });
    }
});

app.post('/ponto/saida', async (req, res) => {
    const { funcionario_id } = req.body;
    const horaSaidaCompleta = adjustToBrasiliaTime(new Date()).toTimeString().slice(0, 8);

    try {
        const jornada = await buscarJornadaAberta(funcionario_id);

        if (!jornada) {
            return res.status(400).json({ message: 'Saída já registrada para hoje.' });
        }

        await query('UPDATE pontos SET saida = ? WHERE id = ?', [horaSaidaCompleta, jornada.id]);
        return res.json({ message: 'Saída registrada com sucesso.' });
    } catch (err) {
        console.error('Erro ao registrar saída:', err);
        return res.status(500).json({ error: 'Erro ao registrar saída.' });
    }
});


/**
 * Monta o relatório do período. Usada tanto pela rota JSON (tela) quanto
 * pela exportação em Excel, para os dois nunca divergirem.
 * Lança Error com .status quando a entrada é inválida.
 */
async function montarRelatorio({ funcionario_id, mes, data_inicio, data_fim, meta_horas }) {
  const erro = (mensagem, status) => Object.assign(new Error(mensagem), { status });
  const metaHoras = Number(meta_horas) > 0 ? Number(meta_horas) : META_MENSAL_HORAS;

  if (!funcionario_id) throw erro('Informe o funcionário', 400);

  // Mês inteiro: do dia 1 ao último dia do mês
  if (mes) {
    const [ano, mesNumero] = String(mes).split('-').map(Number);
    if (!Number.isFinite(ano) || !Number.isFinite(mesNumero)) {
      throw erro('Mês inválido. Use o formato AAAA-MM.', 400);
    }
    data_inicio = formatarDataISO(new Date(ano, mesNumero - 1, 1));
    data_fim = formatarDataISO(new Date(ano, mesNumero, 0));
  }

  if (!data_inicio || !data_fim) {
    throw erro('Informe o mês ou o período (data início e data fim)', 400);
  }

  const [funcionario] = await query(
    'SELECT id, nome, dia_folga FROM funcionarios WHERE id = ?',
    [funcionario_id]
  );
  if (!funcionario) throw erro('Funcionário não encontrado', 404);

  const pontos = await query(
    `SELECT data, entrada, saida_almoco, volta_almoco, saida, tipo_dia
       FROM pontos
      WHERE funcionario_id = ? AND data BETWEEN ? AND ?
      ORDER BY data`,
    [funcionario_id, data_inicio, data_fim]
  );

  const mapPontos = {};
  for (const p of pontos) mapPontos[String(p.data)] = p;

  const hojeISO = formatarDataISO(new Date());
  const dias = [];
  const noturnos = [];

  const totais = {
    minutosTrabalhados: 0,
    minutosNoturnos: 0,
    minutosIntervalo: 0,
    diasTrabalhados: 0,
    faltas: 0,
    folgas: 0,
    abonados: 0,
    previstos: 0,
    inconsistencias: 0
  };

  const fim = parseDataISO(data_fim);
  for (let d = parseDataISO(data_inicio); d <= fim; d.setDate(d.getDate() + 1)) {
    const dataISO = formatarDataISO(d);
    const diaSemana = NOMES_DIA_SEMANA[d.getDay()];
    const ponto = mapPontos[dataISO];
    const tipoManual = ponto && ponto.tipo_dia ? String(ponto.tipo_dia).toUpperCase() : null;
    const temBatidas = !!(ponto && (ponto.entrada || ponto.saida));

    const linha = {
      data: dataISO,
      dataBR: formatarDataBR(dataISO),
      diaSemana,
      entrada: (ponto && ponto.entrada) || 'N/A',
      saida_almoco: (ponto && ponto.saida_almoco) || 'N/A',
      volta_almoco: (ponto && ponto.volta_almoco) || 'N/A',
      saida: (ponto && ponto.saida) || 'N/A',
      status: 'FALTA',
      tipoManual,
      minutosTrabalhados: 0,
      minutosIntervalo: 0,
      minutosNoturnos: 0,
      saldoMinutos: 0,
      inconsistente: false,
      observacao: ''
    };

    // 1) Tipo marcado à mão vence tudo (é assim que a folga variável é corrigida)
    if (tipoManual && tipoManual !== 'NORMAL') {
      linha.status = tipoManual;
      linha.saldoMinutos = 0;
      if (tipoManual === 'FOLGA') totais.folgas++;
      else if (tipoManual === 'FALTA') { totais.faltas++; linha.saldoMinutos = -JORNADA_DIARIA_MIN; }
      else totais.abonados++;

      // Um dia de folga/abonado ainda pode ter batidas: se tiver, as horas contam
      if (temBatidas) {
        const calculo = calcularDiaTrabalhado(ponto);
        linha.minutosTrabalhados = calculo.minutosTrabalhados;
        linha.minutosIntervalo = calculo.minutosIntervalo;
        linha.minutosNoturnos = calculo.minutosNoturnos;
        linha.observacao = calculo.observacao;
        linha.saldoMinutos = calculo.minutosTrabalhados; // trabalhou na folga: tudo vira saldo
      }
    }
    // 2) Dia com batidas de ponto
    else if (temBatidas) {
      const calculo = calcularDiaTrabalhado(ponto);
      linha.status = 'REGISTRADO';
      linha.minutosTrabalhados = calculo.minutosTrabalhados;
      linha.minutosIntervalo = calculo.minutosIntervalo;
      linha.minutosNoturnos = calculo.minutosNoturnos;
      linha.inconsistente = calculo.inconsistente;
      linha.observacao = calculo.observacao;
      linha.saldoMinutos = calculo.minutosTrabalhados - JORNADA_DIARIA_MIN;
      totais.diasTrabalhados++;
      if (calculo.inconsistente) totais.inconsistencias++;
    }
    // 3) Dia futuro: ainda não é falta
    else if (dataISO > hojeISO) {
      linha.status = 'PREVISTO';
      totais.previstos++;
    }
    // 4) Folga fixa cadastrada
    else if (ehDiaDeFolgaFixa(diaSemana, funcionario.dia_folga)) {
      linha.status = 'FOLGA';
      totais.folgas++;
    }
    // 5) Sem registro e sem justificativa: falta (-8h)
    else {
      linha.status = 'FALTA';
      linha.saldoMinutos = -JORNADA_DIARIA_MIN;
      totais.faltas++;
    }

    totais.minutosTrabalhados += linha.minutosTrabalhados;
    totais.minutosIntervalo += linha.minutosIntervalo;
    totais.minutosNoturnos += linha.minutosNoturnos;

    dias.push(linha);

    if (linha.minutosNoturnos > 0) {
      noturnos.push({
        data: linha.data,
        dataBR: linha.dataBR,
        diaSemana: linha.diaSemana,
        entrada: linha.entrada,
        saida: linha.saida,
        minutosNoturnos: linha.minutosNoturnos
      });
    }
  }

  // Arredonda os minutos (as batidas têm segundos) e fecha os totais
  for (const linha of dias) {
    linha.minutosTrabalhados = Math.round(linha.minutosTrabalhados);
    linha.minutosIntervalo = Math.round(linha.minutosIntervalo);
    linha.minutosNoturnos = Math.round(linha.minutosNoturnos);
    linha.saldoMinutos = Math.round(linha.saldoMinutos);
  }
  for (const n of noturnos) n.minutosNoturnos = Math.round(n.minutosNoturnos);

  totais.minutosTrabalhados = Math.round(totais.minutosTrabalhados);
  totais.minutosIntervalo = Math.round(totais.minutosIntervalo);
  totais.minutosNoturnos = Math.round(totais.minutosNoturnos);

  const metaMinutos = Math.round(metaHoras * 60);
  const saldoMinutos = totais.minutosTrabalhados - metaMinutos;
  // Visão alternativa: soma dia a dia (cada falta pesa -8h, folga pesa 0)
  const somaSaldosDiarios = dias.reduce((acc, l) => acc + (l.status === 'PREVISTO' ? 0 : l.saldoMinutos), 0);

  return {
    funcionario: { id: funcionario.id, nome: funcionario.nome, dia_folga: funcionario.dia_folga },
    periodo: { data_inicio, data_fim, mes: mes || null },
    regras: {
      jornadaDiariaMinutos: JORNADA_DIARIA_MIN,
      jantaPadraoMinutos: JANTA_PADRAO_MIN,
      noturnoInicio: '22:00',
      noturnoFim: '05:00',
      metaHoras
    },
    pontos: dias,
    noturnos,
    totais: Object.assign({}, totais, {
      metaMinutos,
      saldoMinutos,
      somaSaldosDiarios,
      horasExtrasMinutos: Math.max(0, saldoMinutos),
      horasDevidasMinutos: Math.max(0, -saldoMinutos),
      // total efetivo somando a hora dobrada da faixa noturna
      minutosComAdicionalNoturno: totais.minutosTrabalhados + totais.minutosNoturnos
    })
  };
}

/* =======================================================================
   RELATÓRIO MENSAL (JSON para a tela)
   -----------------------------------------------------------------------
   Aceita ?mes=YYYY-MM (mês inteiro, com faltas e folgas) ou
   ?data_inicio=&data_fim= para um período livre.
   ======================================================================= */
app.get('/relatorio', async (req, res) => {
  try {
    const relatorio = await montarRelatorio(req.query);
    res.json(relatorio);
  } catch (err) {
    if (err.status) return res.status(err.status).json({ error: err.message });
    console.error('Erro ao gerar relatório:', err);
    return res.status(500).json({ error: 'Erro interno ao gerar relatório' });
  }
});

/* =======================================================================
   EXPORTAÇÃO EM EXCEL (.xlsx) JÁ FORMATADA
   -----------------------------------------------------------------------
   Planilha pronta para imprimir/assinar: cabeçalho, larguras de coluna,
   bordas, cores por status, durações como hora real (para o Excel somar)
   e a tabela de adicional noturno em aba própria.
   ======================================================================= */

// Cores (ARGB, como o ExcelJS espera)
const XLS = {
    tinta: 'FF1F2933',
    cabecalhoFundo: 'FF243447',
    cabecalhoTinta: 'FFFFFFFF',
    tituloFundo: 'FF0B5FA5',
    zebra: 'FFF7F9FC',
    borda: 'FFB6C2CF',
    folga: 'FFDFF3E1',
    falta: 'FFFCE1E1',
    abonado: 'FFE4ECFB',
    previsto: 'FFF0F0F0',
    alerta: 'FFFFF3CD',
    negativo: 'FFC0281C',
    positivo: 'FF1B6E2F',
    resumoRotulo: 'FFEDF1F6'
};

const BORDA_FINA = {
    top: { style: 'thin', color: { argb: XLS.borda } },
    left: { style: 'thin', color: { argb: XLS.borda } },
    bottom: { style: 'thin', color: { argb: XLS.borda } },
    right: { style: 'thin', color: { argb: XLS.borda } }
};

// Excel guarda hora como fração de dia. Assim a célula é um valor numérico
// de verdade — o usuário pode somar, e ainda aparece como [h]:mm.
function minutosParaFracaoDia(minutos) {
    return (minutos || 0) / 1440;
}

const ROTULO_DIA_CURTO = {
    'DOMINGO': 'Dom', 'SEGUNDA-FEIRA': 'Seg', 'TERCA-FEIRA': 'Ter', 'QUARTA-FEIRA': 'Qua',
    'QUINTA-FEIRA': 'Qui', 'SEXTA-FEIRA': 'Sex', 'SABADO': 'Sáb'
};

const FUNDO_POR_STATUS = {
    'FOLGA': XLS.folga,
    'FALTA': XLS.falta,
    'ATESTADO': XLS.abonado,
    'FERIADO': XLS.abonado,
    'FERIAS': XLS.abonado,
    'PREVISTO': XLS.previsto
};

const NOMES_MES = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
                   'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];

function descreverPeriodo(periodo) {
    if (periodo.mes) {
        const [ano, mes] = periodo.mes.split('-');
        return NOMES_MES[Number(mes) - 1] + ' de ' + ano;
    }
    return formatarDataBR(periodo.data_inicio) + ' a ' + formatarDataBR(periodo.data_fim);
}

function pintar(celula, argb) {
    celula.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } };
}

// Faixa de título mesclada no topo de uma aba
function escreverTitulo(aba, colunas, texto, subtitulo) {
    const linhaTitulo = aba.addRow([texto]);
    aba.mergeCells(linhaTitulo.number, 1, linhaTitulo.number, colunas);
    const celula = linhaTitulo.getCell(1);
    celula.font = { bold: true, size: 14, color: { argb: XLS.cabecalhoTinta } };
    celula.alignment = { horizontal: 'center', vertical: 'middle' };
    pintar(celula, XLS.tituloFundo);
    linhaTitulo.height = 26;

    if (subtitulo) {
        const linhaSub = aba.addRow([subtitulo]);
        aba.mergeCells(linhaSub.number, 1, linhaSub.number, colunas);
        const celulaSub = linhaSub.getCell(1);
        celulaSub.font = { bold: true, size: 11, color: { argb: XLS.tinta } };
        celulaSub.alignment = { horizontal: 'center', vertical: 'middle' };
        pintar(celulaSub, XLS.resumoRotulo);
        linhaSub.height = 20;
    }
    aba.addRow([]);
}

function estilizarCabecalho(linha, colunas) {
    linha.height = 30;
    for (let c = 1; c <= colunas; c++) {
        const celula = linha.getCell(c);
        celula.font = { bold: true, size: 10, color: { argb: XLS.cabecalhoTinta } };
        celula.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        pintar(celula, XLS.cabecalhoFundo);
        celula.border = BORDA_FINA;
    }
}

function montarAbaRelatorio(planilha, relatorio) {
    const aba = planilha.addWorksheet('Relatório', {
        views: [{ state: 'frozen', ySplit: 0 }],
        pageSetup: { paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 } }
    });

    const COLUNAS = [
        { titulo: 'Data', largura: 12 },
        { titulo: 'Dia', largura: 7 },
        { titulo: 'Entrada', largura: 10 },
        { titulo: 'Saída Janta', largura: 11 },
        { titulo: 'Volta Janta', largura: 11 },
        { titulo: 'Saída', largura: 10 },
        { titulo: 'Trabalhado', largura: 12 },
        { titulo: 'Saldo do Dia', largura: 12 },
        { titulo: 'Adic. Noturno', largura: 13 },
        { titulo: 'Status', largura: 13 },
        { titulo: 'Observação', largura: 42 }
    ];
    aba.columns = COLUNAS.map(c => ({ width: c.largura }));

    escreverTitulo(aba, COLUNAS.length,
        'ESPELHO DE PONTO — ' + relatorio.funcionario.nome.toUpperCase(),
        'Período: ' + descreverPeriodo(relatorio.periodo) +
        '   |   Jornada 17:00 às 02:00   |   Janta 30 min   |   Adicional noturno 22:00 às 05:00');

    const linhaCabecalho = aba.addRow(COLUNAS.map(c => c.titulo));
    estilizarCabecalho(linhaCabecalho, COLUNAS.length);
    // Congela tudo acima da primeira linha de dados
    aba.views = [{ state: 'frozen', ySplit: linhaCabecalho.number }];
    aba.autoFilter = {
        from: { row: linhaCabecalho.number, column: 1 },
        to: { row: linhaCabecalho.number, column: COLUNAS.length }
    };

    const primeiraLinhaDados = linhaCabecalho.number + 1;

    relatorio.pontos.forEach((p, indice) => {
        const semRegistro = p.status !== 'REGISTRADO' && !p.minutosTrabalhados;
        const linha = aba.addRow([
            p.dataBR,
            ROTULO_DIA_CURTO[p.diaSemana] || p.diaSemana,
            p.entrada === 'N/A' ? '—' : p.entrada,
            p.saida_almoco === 'N/A' ? '—' : p.saida_almoco,
            p.volta_almoco === 'N/A' ? '—' : p.volta_almoco,
            p.saida === 'N/A' ? '—' : p.saida,
            semRegistro ? null : minutosParaFracaoDia(p.minutosTrabalhados),
            p.status === 'PREVISTO' ? null : minutosParaFracaoDia(p.saldoMinutos),
            p.minutosNoturnos ? minutosParaFracaoDia(p.minutosNoturnos) : null,
            p.status,
            p.observacao || ''
        ]);
        linha.height = 18;

        const fundo = FUNDO_POR_STATUS[p.status] || (p.inconsistente ? XLS.alerta : (indice % 2 ? XLS.zebra : null));

        for (let c = 1; c <= COLUNAS.length; c++) {
            const celula = linha.getCell(c);
            celula.border = BORDA_FINA;
            celula.font = { size: 10, color: { argb: XLS.tinta } };
            celula.alignment = { horizontal: c === 11 ? 'left' : 'center', vertical: 'middle' };
            if (fundo) pintar(celula, fundo);
        }

        // Durações como hora real, somáveis no Excel
        [7, 8, 9].forEach(c => { linha.getCell(c).numFmt = '[h]:mm'; });

        // Saldo negativo em vermelho, positivo em verde
        const celulaSaldo = linha.getCell(8);
        if (typeof celulaSaldo.value === 'number' && celulaSaldo.value !== 0) {
            celulaSaldo.font = {
                size: 10, bold: true,
                color: { argb: celulaSaldo.value < 0 ? XLS.negativo : XLS.positivo }
            };
            celulaSaldo.numFmt = '[h]:mm;-[h]:mm';
        }

        const celulaStatus = linha.getCell(10);
        celulaStatus.font = { size: 10, bold: true, color: { argb: XLS.tinta } };

        linha.getCell(11).font = { size: 9, italic: true, color: { argb: XLS.tinta } };
    });

    const ultimaLinhaDados = aba.lastRow.number;

    // ---- Linha de totais, com SUBTOTAL para respeitar o filtro ----
    const linhaTotal = aba.addRow([
        'TOTAIS', '', '', '', '', '',
        { formula: `SUBTOTAL(109,G${primeiraLinhaDados}:G${ultimaLinhaDados})` },
        { formula: `SUBTOTAL(109,H${primeiraLinhaDados}:H${ultimaLinhaDados})` },
        { formula: `SUBTOTAL(109,I${primeiraLinhaDados}:I${ultimaLinhaDados})` },
        '', ''
    ]);
    aba.mergeCells(linhaTotal.number, 1, linhaTotal.number, 6);
    linhaTotal.height = 24;
    for (let c = 1; c <= COLUNAS.length; c++) {
        const celula = linhaTotal.getCell(c);
        celula.border = BORDA_FINA;
        celula.font = { bold: true, size: 11, color: { argb: XLS.cabecalhoTinta } };
        celula.alignment = { horizontal: c === 1 ? 'right' : 'center', vertical: 'middle' };
        pintar(celula, XLS.cabecalhoFundo);
    }
    [7, 8, 9].forEach(c => { linhaTotal.getCell(c).numFmt = '[h]:mm;-[h]:mm'; });

    return aba;
}

function montarAbaResumo(planilha, relatorio) {
    const aba = planilha.addWorksheet('Resumo', {
        pageSetup: { paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 }
    });
    aba.columns = [{ width: 34 }, { width: 16 }, { width: 52 }];

    escreverTitulo(aba, 3,
        'RESUMO DO PERÍODO',
        relatorio.funcionario.nome + '   |   ' + descreverPeriodo(relatorio.periodo));

    const t = relatorio.totais;
    const HORA = 'hora';
    const NUM = 'numero';

    const secoes = [
        ['HORAS', [
            ['Total trabalhado', t.minutosTrabalhados, HORA, 'Soma das horas líquidas, já sem o intervalo de janta'],
            ['Carga horária do mês', t.metaMinutos, HORA, 'Meta contratual (' + relatorio.regras.metaHoras + ' horas)'],
            ['Saldo do mês', t.saldoMinutos, HORA, 'Total trabalhado menos a carga horária do mês'],
            ['Horas extras', t.horasExtrasMinutos, HORA, 'O que passou da carga horária do mês'],
            ['Horas devidas', t.horasDevidasMinutos, HORA, 'O que faltou para fechar a carga horária'],
            ['Soma dos saldos diários', t.somaSaldosDiarios, HORA, 'Banco de horas dia a dia: cada falta pesa -8h, folga pesa 0']
        ]],
        ['ADICIONAL NOTURNO', [
            ['Adicional noturno (total)', t.minutosNoturnos, HORA, 'Horas na faixa 22:00–05:00, pagas em dobro'],
            ['Trabalhado + adicional noturno', t.minutosComAdicionalNoturno, HORA, 'Total efetivo com a hora noturna dobrada']
        ]],
        ['DIAS', [
            ['Dias trabalhados', t.diasTrabalhados, NUM, ''],
            ['Faltas', t.faltas, NUM, 'Dias sem registro e sem justificativa'],
            ['Folgas', t.folgas, NUM, 'Folga fixa ou marcada à mão — não desconta as 8h'],
            ['Dias abonados', t.abonados, NUM, 'Atestado, feriado ou férias'],
            ['Dias previstos', t.previstos, NUM, 'Dias futuros, ainda sem registro'],
            ['Registros incompletos', t.inconsistencias, NUM, 'Dias com batida faltando — confira antes de fechar']
        ]]
    ];

    for (const [tituloSecao, itens] of secoes) {
        const linhaSecao = aba.addRow([tituloSecao]);
        aba.mergeCells(linhaSecao.number, 1, linhaSecao.number, 3);
        const celulaSecao = linhaSecao.getCell(1);
        celulaSecao.font = { bold: true, size: 10, color: { argb: XLS.cabecalhoTinta } };
        celulaSecao.alignment = { horizontal: 'left', vertical: 'middle' };
        pintar(celulaSecao, XLS.cabecalhoFundo);
        linhaSecao.height = 20;

        for (const [rotulo, valor, tipo, nota] of itens) {
            const linha = aba.addRow([
                rotulo,
                tipo === HORA ? minutosParaFracaoDia(valor) : valor,
                nota
            ]);
            linha.height = 19;

            const celulaRotulo = linha.getCell(1);
            celulaRotulo.font = { bold: true, size: 10, color: { argb: XLS.tinta } };
            celulaRotulo.alignment = { horizontal: 'left', vertical: 'middle' };
            pintar(celulaRotulo, XLS.resumoRotulo);

            const celulaValor = linha.getCell(2);
            celulaValor.alignment = { horizontal: 'center', vertical: 'middle' };
            celulaValor.font = {
                bold: true, size: 11,
                color: { argb: tipo === HORA && valor < 0 ? XLS.negativo : XLS.tinta }
            };
            if (tipo === HORA) celulaValor.numFmt = '[h]:mm;-[h]:mm';

            const celulaNota = linha.getCell(3);
            celulaNota.font = { size: 9, italic: true, color: { argb: XLS.tinta } };
            celulaNota.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };

            for (let c = 1; c <= 3; c++) linha.getCell(c).border = BORDA_FINA;
        }
        aba.addRow([]);
    }

    const rodape = aba.addRow(['Gerado em ' + formatarDataBR(formatarDataISO(new Date())) +
        ' — os valores de hora são numéricos e podem ser somados no Excel.']);
    aba.mergeCells(rodape.number, 1, rodape.number, 3);
    rodape.getCell(1).font = { size: 9, italic: true, color: { argb: XLS.tinta } };

    return aba;
}

function montarAbaNoturno(planilha, relatorio) {
    const aba = planilha.addWorksheet('Adicional Noturno', {
        pageSetup: { paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 }
    });

    const COLUNAS = [
        { titulo: 'Data', largura: 12 },
        { titulo: 'Dia', largura: 8 },
        { titulo: 'Entrada', largura: 11 },
        { titulo: 'Saída', largura: 11 },
        { titulo: 'Adicional Noturno', largura: 18 }
    ];
    aba.columns = COLUNAS.map(c => ({ width: c.largura }));

    escreverTitulo(aba, COLUNAS.length,
        'ADICIONAL NOTURNO (22:00 ÀS 05:00)',
        relatorio.funcionario.nome + '   |   ' + descreverPeriodo(relatorio.periodo) +
        '   |   A contagem normal fica na aba Relatório; aqui vai a hora dobrada');

    const linhaCabecalho = aba.addRow(COLUNAS.map(c => c.titulo));
    estilizarCabecalho(linhaCabecalho, COLUNAS.length);
    aba.views = [{ state: 'frozen', ySplit: linhaCabecalho.number }];

    const primeira = linhaCabecalho.number + 1;

    if (!relatorio.noturnos.length) {
        const vazia = aba.addRow(['Nenhuma hora na faixa noturna neste período.']);
        aba.mergeCells(vazia.number, 1, vazia.number, COLUNAS.length);
        vazia.getCell(1).alignment = { horizontal: 'center' };
        vazia.getCell(1).font = { italic: true, size: 10, color: { argb: XLS.tinta } };
        return aba;
    }

    relatorio.noturnos.forEach((n, indice) => {
        const linha = aba.addRow([
            n.dataBR,
            ROTULO_DIA_CURTO[n.diaSemana] || n.diaSemana,
            n.entrada,
            n.saida,
            minutosParaFracaoDia(n.minutosNoturnos)
        ]);
        linha.height = 18;
        for (let c = 1; c <= COLUNAS.length; c++) {
            const celula = linha.getCell(c);
            celula.border = BORDA_FINA;
            celula.font = { size: 10, color: { argb: XLS.tinta } };
            celula.alignment = { horizontal: 'center', vertical: 'middle' };
            if (indice % 2) pintar(celula, XLS.zebra);
        }
        linha.getCell(5).numFmt = '[h]:mm';
        linha.getCell(5).font = { size: 10, bold: true, color: { argb: XLS.tinta } };
    });

    const ultima = aba.lastRow.number;

    const linhaTotal = aba.addRow(['TOTAL DE ADICIONAL NOTURNO', '', '', '',
        { formula: `SUM(E${primeira}:E${ultima})` }]);
    aba.mergeCells(linhaTotal.number, 1, linhaTotal.number, 4);
    linhaTotal.height = 24;
    for (let c = 1; c <= COLUNAS.length; c++) {
        const celula = linhaTotal.getCell(c);
        celula.border = BORDA_FINA;
        celula.font = { bold: true, size: 11, color: { argb: XLS.cabecalhoTinta } };
        celula.alignment = { horizontal: c === 1 ? 'right' : 'center', vertical: 'middle' };
        pintar(celula, XLS.cabecalhoFundo);
    }
    linhaTotal.getCell(5).numFmt = '[h]:mm';

    const linhaEfetivo = aba.addRow(['TRABALHADO + ADICIONAL NOTURNO', '', '', '',
        minutosParaFracaoDia(relatorio.totais.minutosComAdicionalNoturno)]);
    aba.mergeCells(linhaEfetivo.number, 1, linhaEfetivo.number, 4);
    linhaEfetivo.height = 22;
    for (let c = 1; c <= COLUNAS.length; c++) {
        const celula = linhaEfetivo.getCell(c);
        celula.border = BORDA_FINA;
        celula.font = { bold: true, size: 10, color: { argb: XLS.tinta } };
        celula.alignment = { horizontal: c === 1 ? 'right' : 'center', vertical: 'middle' };
        pintar(celula, XLS.resumoRotulo);
    }
    linhaEfetivo.getCell(5).numFmt = '[h]:mm';

    return aba;
}

app.get('/relatorio/exportar', async (req, res) => {
    try {
        const relatorio = await montarRelatorio(req.query);

        const planilha = new ExcelJS.Workbook();
        planilha.creator = 'Sistema de Ponto';
        planilha.created = new Date();

        montarAbaRelatorio(planilha, relatorio);
        montarAbaNoturno(planilha, relatorio);
        montarAbaResumo(planilha, relatorio);

        const referencia = relatorio.periodo.mes || relatorio.periodo.data_inicio;
        const nomeLimpo = String(relatorio.funcionario.nome)
            .normalize('NFD').replace(/[̀-ͯ]/g, '')
            .replace(/[^A-Za-z0-9]+/g, '-').replace(/^-|-$/g, '').toLowerCase();
        const nomeArquivo = `espelho-ponto-${nomeLimpo || relatorio.funcionario.id}-${referencia}.xlsx`;

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${nomeArquivo}"`);
        await planilha.xlsx.write(res);
        res.end();
    } catch (err) {
        if (err.status) return res.status(err.status).send(err.message);
        console.error('Erro ao exportar relatório:', err);
        return res.status(500).send('Erro ao exportar o relatório');
    }
});


/* Marca/desmarca o tipo de um dia — usado para transformar em FOLGA o dia
   que caiu como FALTA, já que a folga da semana não é fixa. */
app.post('/ponto/tipo-dia', async (req, res) => {
  const { funcionario_id, data, tipo_dia } = req.body;

  if (!funcionario_id || !data) {
    return res.status(400).json({ error: 'Informe o funcionário e a data' });
  }

  const tipo = tipo_dia ? String(tipo_dia).toUpperCase() : null;
  if (tipo && !TIPOS_DIA_VALIDOS.includes(tipo)) {
    return res.status(400).json({ error: 'Tipo inválido. Use: ' + TIPOS_DIA_VALIDOS.join(', ') });
  }

  try {
    const [existente] = await query(
      `SELECT id, entrada, saida_almoco, volta_almoco, saida
         FROM pontos WHERE funcionario_id = ? AND data = ?`,
      [funcionario_id, data]
    );

    if (existente) {
      const semBatidas = !existente.entrada && !existente.saida_almoco &&
                         !existente.volta_almoco && !existente.saida;
      if (!tipo && semBatidas) {
        // A linha só existia para guardar a marcação: sem ela, não sobra nada
        await query('DELETE FROM pontos WHERE id = ?', [existente.id]);
      } else {
        await query('UPDATE pontos SET tipo_dia = ? WHERE id = ?', [tipo, existente.id]);
      }
    } else if (tipo) {
      await query(
        'INSERT INTO pontos (funcionario_id, data, tipo_dia) VALUES (?, ?, ?)',
        [funcionario_id, data, tipo]
      );
    }

    return res.json({
      message: tipo
        ? 'Dia ' + formatarDataBR(data) + ' marcado como ' + tipo + '.'
        : 'Marcação do dia ' + formatarDataBR(data) + ' removida.'
    });
  } catch (err) {
    console.error('Erro ao marcar tipo do dia:', err);
    return res.status(500).json({ error: 'Erro ao marcar o tipo do dia' });
  }
});


app.post('/ponto/editar', async (req, res) => {
    const { funcionario_id, data, entrada, saida_almoco, volta_almoco, saida, tipo_dia } = req.body;

    if (!funcionario_id || !data) {
        return res.status(400).send('Informe o funcionário e a data');
    }

    const tipo = tipo_dia ? String(tipo_dia).toUpperCase() : null;
    if (tipo && !TIPOS_DIA_VALIDOS.includes(tipo)) {
        return res.status(400).send('Tipo de dia inválido');
    }

    // Os horários são gravados exatamente como batidos. A virada de dia
    // (saída às 02:00 de uma entrada às 17:00) é resolvida no cálculo do
    // relatório, e não deslocando o horário na gravação.
    const valor = (h) => (h && String(h).trim() !== '' ? String(h).trim() : null);

    try {
        const [existing] = await query(
            'SELECT id FROM pontos WHERE funcionario_id = ? AND data = ?',
            [funcionario_id, data]
        );

        if (existing) {
            await query(
                `UPDATE pontos
                    SET entrada = ?, saida_almoco = ?, volta_almoco = ?, saida = ?, tipo_dia = ?
                  WHERE id = ?`,
                [valor(entrada), valor(saida_almoco), valor(volta_almoco), valor(saida), tipo, existing.id]
            );
            console.log('✅ Registro atualizado com sucesso:', { funcionario_id, data, tipo });
            return res.send('Registro de ponto atualizado com sucesso');
        }

        await query(
            `INSERT INTO pontos (funcionario_id, data, entrada, saida_almoco, volta_almoco, saida, tipo_dia)
             VALUES (?, ?, ?, ?, ?, ?, ?)`,
            [funcionario_id, data, valor(entrada), valor(saida_almoco), valor(volta_almoco), valor(saida), tipo]
        );
        console.log('🆕 Novo registro inserido:', { funcionario_id, data, tipo });
        return res.send('Novo registro de ponto criado com sucesso');
    } catch (err) {
        console.error('❌ Erro ao editar/inserir ponto:', err);
        res.status(500).send('Erro ao editar ou inserir ponto');
    }
});


app.get('/ponto/buscar', async (req, res) => {
    const { funcionario_id, data } = req.query; // Obtém os parâmetros enviados na URL

    // Consulta SQL para buscar os dados do ponto
    const sql = `
        SELECT entrada, saida_almoco, volta_almoco, saida, tipo_dia
        FROM pontos
        WHERE funcionario_id = ? AND data = ?
    `;

    try {
        const [result] = await query(sql, [funcionario_id, data]);

        if (!result) {
            return res.status(404).send('Registro não encontrado');
        }

        res.json(result); // Retorna os dados encontrados
    } catch (err) {
        console.error('Erro ao buscar registro de ponto:', err);
        res.status(500).send('Erro ao buscar registro de ponto');
    }
});


app.post('/funcionarios/cadastrar', async (req, res) => {
    const { id, nome, email, dia_folga, senha } = req.body;

    try {
        // Gerando o hash da senha
        bcrypt.hash(senha, saltRounds, async (err, hash) => {
            if (err) {
                console.error('Erro ao gerar o hash da senha:', err);
                return res.status(500).send('Erro ao gerar a senha');
            }

            // Inserindo o nome, email e o hash da senha no banco de dados
            const sql = 'INSERT INTO funcionarios (id, nome, email, dia_folga, senha) VALUES (?, ?, ?, ?, ?)';
            const data = await query(sql, [id, nome, email, dia_folga, hash]); // Salvando o hash no banco
            console.log("data", data);
            res.redirect('/');
        });
    } catch (err) {
        if (err.code === 'ER_DUP_ENTRY') {
            return res.status(400).send('Email já cadastrado.');
        }
        console.error('Erro ao cadastrar funcionário:', err);
        return res.status(500).send('Erro ao cadastrar funcionário');
    }
});

app.get('/funcionarios', async (req, res) => {
    const sql = 'SELECT id, nome FROM funcionarios';
    try {
        const result = await query(sql);
        res.json(result);
    } catch (err) {
        console.error('Erro ao buscar funcionários:', err);
        res.status(500).json({ error: 'Erro ao buscar funcionários' });
    }
});

app.post('/deletar-funcionario', (req, res) => {
    const { funcionario_id } = req.body;

    if (!funcionario_id)
        return res.status(400).json({ message: 'ID do funcionário é obrigatório' });

    const verificarFuncionarioSQL = 'SELECT * FROM funcionarios WHERE id = ?';
    const deletePontosSQL = 'DELETE FROM pontos WHERE funcionario_id = ?';
    const deleteFuncionarioSQL = 'DELETE FROM funcionarios WHERE id = ?';

    // Verifica se o funcionário existe
    db.query(verificarFuncionarioSQL, [funcionario_id], (err, results) => {
        if (err)
            return res.status(500).send('Erro ao verificar funcionário');

        if (results.length === 0)
            return res.status(404).json({ message: 'Funcionário não encontrado' });

        // Verifica se o ID é 3107 e impede a exclusão
        if (funcionario_id == 6)
            return res.status(403).json({ message: 'Administrador do sistema não pode ser excluído' });

        // Exclui pontos associados ao funcionário
        db.query(deletePontosSQL, [funcionario_id], (err) => {
            if (err)
                return res.status(500).send('Erro ao excluir pontos');

            // Exclui o funcionário
            db.query(deleteFuncionarioSQL, [funcionario_id], (err, result) => {
                if (err)
                    return res.status(500).send('Erro ao excluir funcionário');

                if (result.affectedRows === 0)
                    return res.status(404).json({ message: 'Funcionário não encontrado para exclusão' });

                res.redirect('/login');
            });
        });
    });
});



app.get('/login', (req, res) => {
    res.render('login');
});

// Rota para autenticação de login
app.post('/login', (req, res) => {
    const { funcionario_id } = req.body;

    if (!funcionario_id) {
        return res.status(400).json({ message: 'Informe o ID do funcionário.' });
    }

    const sql = 'SELECT id FROM funcionarios WHERE id = ?';
    db.query(sql, [funcionario_id], (err, results) => {
        if (err) {
            console.error("Erro na consulta:", err);
            return res.status(500).json({ message: 'Erro no servidor.' });
        }
        if (results.length === 0) {
            return res.status(404).json({ message: 'Funcionário não encontrado.' });
        }

        console.log("Login (apenas ID) bem-sucedido:", funcionario_id);
        req.session.funcionarioId = funcionario_id;
        res.status(200).json({ autenticado: true, message: 'Login realizado com sucesso!' });
    });
});


// Middleware para verificar autenticação do usuário
function verificarAutenticacao(req, res, next) {
    if (req.session && req.session.funcionarioId) return next();
    res.redirect('/login');
}

// Rota protegida para o ponto (index)
app.get('/index/:funcionario_id', verificarAutenticacao, (req, res) => {
    const funcionarioId = req.params.funcionario_id; // ID do funcionário logado
    const sql = 'SELECT * FROM funcionarios';

    db.query(sql, (error, results) => {
        if (error) {
            console.error('Erro ao buscar funcionários:', error);
            return res.status(500).send('Erro ao buscar funcionários');
        }

        // Renderize o EJS com todos os funcionários e o ID do funcionário logado
        res.render('index', { funcionarios: results, funcionarioLogado: funcionarioId });
    });
});


app.get('/buscar', async (req, res) => {
    const sql = 'SELECT id, nome FROM funcionarios';

    try {
        const funcionarios = await query(sql); // Executa a consulta no banco de dados
        console.log("resultado funcionarios", funcionarios);

        res.json(funcionarios); // Retorna os dados como JSON para o frontend
    } catch (error) {
        console.error('Erro ao buscar funcionários:', error);
        res.status(500).send('Erro ao buscar funcionários');
    }
});

// Rota para verificar uma senha específica para autenticação adicional
app.post('/verificar-senha', async (req, res) => {
    const { senha, funcionario_id } = req.body;

    const sql = 'SELECT id, nome FROM funcionarios WHERE id = ?';

    try {
        const [funcionario] = await query(sql, [funcionario_id]);

        if (funcionario.id !== 3302) {
            return res.status(404).json({ message: 'Acesso Negado' });
        }

        const senhaCorreta = '123'; // 👈 Ainda fixa, você pode usar bcrypt depois

        if (senha === senhaCorreta) {
            res.json({ autenticado: true });
        } else {
            res.json({ autenticado: false });
        }
    } catch (err) {
        console.error('Erro ao verificar senha:', err);
        res.status(500).json({ message: 'Erro interno no servidor' });
    }
});
// Logout
app.get('/logout', (req, res) => {
    req.session.destroy(err => {
        if (err) {
            return res.status(500).send('Erro ao fazer logout');
        }
        res.redirect('/login');
    });
});

app.listen(port, () => {
    console.log(`Servidor rodando na porta ${port}`);
});
