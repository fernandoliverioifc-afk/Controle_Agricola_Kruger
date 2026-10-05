const express = require('express');
const session = require('express-session');
const bcrypt = require('bcryptjs');
const ExcelJS = require('exceljs');
const sqlite3 = require('sqlite3').verbose();
const fs = require('fs');
const path = require('path');

const app = express();
const db = new sqlite3.Database('./fazenda.db');
const CAMINHO_PLANILHA = path.join(__dirname, 'controle_financeiro_agricola.xlsx');

app.use(express.json());
app.use(express.static('public'));

app.use(session({
  secret: process.env.SESSION_SECRET || 'chave_secreta_agricola_2026',
  resave: false,
  saveUninitialized: false,
  cookie: { maxAge: 24 * 60 * 60 * 1000 }
}));

// Fila simples para evitar que requisições concorrentes corrompam o arquivo Excel
let filaExcel = Promise.resolve();
function executarNoExcel(tarefaAsync) {
  filaExcel = filaExcel.then(async () => {
    try {
      await tarefaAsync();
    } catch (err) {
      console.error('Erro na operação do Excel:', err);
    }
  });
  return filaExcel;
}

// Função para converter qualquer unidade para a menor fração correspondente
function calcularMenorUnidade(quantidade, unidade) {
  const qtd = parseFloat(quantidade) || 0;
  const unit = (unidade || '').toLowerCase();

  switch (unit) {
    case 'kg':
    case 'l':
      return qtd * 1000;              // 1 Kg = 1000g | 1 L = 1000ml
    case 't':
      return qtd * 1000000;           // 1 Tonelada = 1.000.000g
    case 'm3':
      return qtd * 1000000;           // 1 m³ = 1.000.000ml
    case 'saca_sementes':
      return qtd * 60000;             // 1 Saca = 60.000 sementes
    default:
      return qtd;                     // g, ml, sementes, pacote, unidade, caixa, dose, galão
  }
}

// --- INICIALIZAÇÃO DO BANCO DE DADOS ---
db.serialize(() => {
  db.run(`CREATE TABLE IF NOT EXISTS usuarios (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    usuario TEXT UNIQUE,
    senha TEXT
  )`);

  db.run(`CREATE TABLE IF NOT EXISTS produtos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nome TEXT,
    qtd_comprada REAL,
    unidade TEXT,
    total_ml_g REAL,
    preco_total REAL,
    custo_unitario REAL,
    estoque_restante REAL
  )`);

  db.run(`CREATE TABLE IF NOT EXISTS manejos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    data TEXT,
    talhao TEXT,
    produto_nome TEXT,
    qtd_gasta REAL,
    custo_manejo REAL,
    usuario TEXT
  )`);

  const senhaHash = bcrypt.hashSync('123', 10);
  db.run(`INSERT OR IGNORE INTO usuarios (usuario, senha) VALUES ('agricultor', ?)`, [senhaHash]);
});

async function inicializarPlanilha() {
  if (!fs.existsSync(CAMINHO_PLANILHA)) {
    const workbook = new ExcelJS.Workbook();
    
    const abaCompras = workbook.addWorksheet('Compras e Estoque');
    abaCompras.columns = [
      { header: 'ID Produto', key: 'id', width: 12 },
      { header: 'Nome Produto', key: 'nome', width: 25 },
      { header: 'Qtd Comprada', key: 'qtdComprada', width: 15 },
      { header: 'Unidade', key: 'unidade', width: 15 },
      { header: 'Total Calculado', key: 'totalMlGramas', width: 18 },
      { header: 'Preço Pago (R$)', key: 'precoTotal', width: 15 },
      { header: 'Custo Unitário (R$)', key: 'custoUnitario', width: 20 },
      { header: 'Estoque Restante', key: 'estoqueRestante', width: 22 }
    ];

    const abaManejos = workbook.addWorksheet('Manejos e Aplicacoes');
    abaManejos.columns = [
      { header: 'Data', key: 'data', width: 15 },
      { header: 'Talhão / Área', key: 'talhao', width: 20 },
      { header: 'Produto Usado', key: 'produto', width: 25 },
      { header: 'Qtd Gasta', key: 'qtdGasta', width: 18 },
      { header: 'Custo Proporcional (R$)', key: 'custoManejo', width: 22 },
      { header: 'Lançado por', key: 'usuario', width: 18 }
    ];

    await workbook.xlsx.writeFile(CAMINHO_PLANILHA);
  }
}
inicializarPlanilha();

function autenticado(req, res, next) {
  if (req.session.usuario) return next();
  res.status(401).json({ erro: 'Não autorizado. Faça login.' });
}

// --- ROTAS DA API ---

app.post('/api/login', (req, res) => {
  const { usuario, senha } = req.body;
  db.get(`SELECT * FROM usuarios WHERE usuario = ?`, [usuario], (err, user) => {
    if (err || !user || !bcrypt.compareSync(senha, user.senha)) {
      return res.status(401).json({ erro: 'Usuário ou senha incorretos.' });
    }
    req.session.usuario = user.usuario;
    res.json({ mensagem: 'Login realizado com sucesso!', usuario: user.usuario });
  });
});

app.post('/api/logout', (req, res) => {
  req.session.destroy();
  res.json({ mensagem: 'Sessão encerrada.' });
});

app.get('/api/sessao', (req, res) => {
  res.json({ usuario: req.session.usuario || null });
});

// Alterar Senha do Usuário
app.post('/api/alterar-senha', autenticado, (req, res) => {
  const { senhaAtual, novaSenha } = req.body;
  const usuarioLogado = req.session.usuario;

  if (!novaSenha || novaSenha.length < 3) {
    return res.status(400).json({ erro: 'A nova senha deve ter pelo menos 3 caracteres.' });
  }

  db.get(`SELECT * FROM usuarios WHERE usuario = ?`, [usuarioLogado], (err, user) => {
    if (err || !user || !bcrypt.compareSync(senhaAtual, user.senha)) {
      return res.status(400).json({ erro: 'A senha atual está incorreta.' });
    }

    const novaSenhaHash = bcrypt.hashSync(novaSenha, 10);
    db.run(`UPDATE usuarios SET senha = ? WHERE usuario = ?`, [novaSenhaHash, usuarioLogado], (err) => {
      if (err) return res.status(500).json({ erro: 'Erro ao atualizar a senha.' });
      res.json({ mensagem: 'Senha alterada com sucesso!' });
    });
  });
});

// Baixar Planilha Excel
app.get('/api/download-excel', autenticado, (req, res) => {
  if (fs.existsSync(CAMINHO_PLANILHA)) {
    res.download(CAMINHO_PLANILHA, 'controle_financeiro_agricola.xlsx');
  } else {
    res.status(404).json({ erro: 'Planilha não encontrada.' });
  }
});

// Cadastrar Novo Produto
app.post('/api/compras', autenticado, (req, res) => {
  const { nome, qtdComprada, unidade, preco, tipoPreco } = req.body;

  const totalMlGramas = calcularMenorUnidade(qtdComprada, unidade);
  const qtdNum = parseFloat(qtdComprada) || 1;
  const valorNum = parseFloat(preco) || 0;

  let precoTotal = 0;
  let custoUnitario = 0;

  if (tipoPreco === 'unitario') {
    precoTotal = valorNum * qtdNum;
    custoUnitario = precoTotal / totalMlGramas;
  } else {
    precoTotal = valorNum;
    custoUnitario = precoTotal / totalMlGramas;
  }

  db.run(
    `INSERT INTO produtos (nome, qtd_comprada, unidade, total_ml_g, preco_total, custo_unitario, estoque_restante)
     VALUES (?, ?, ?, ?, ?, ?, ?)`,
    [nome, parseFloat(qtdComprada), unidade, totalMlGramas, precoTotal, custoUnitario, totalMlGramas],
    function (err) {
      if (err) return res.status(500).json({ erro: 'Erro ao salvar no banco.' });

      const novoProdutoId = this.lastID;

      executarNoExcel(async () => {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(CAMINHO_PLANILHA);
        const abaCompras = workbook.getWorksheet('Compras e Estoque');

        abaCompras.addRow({
          id: novoProdutoId,
          nome: nome,
          qtdComprada: parseFloat(qtdComprada),
          unidade: unidade,
          totalMlGramas: totalMlGramas,
          precoTotal: precoTotal,
          custoUnitario: custoUnitario,
          estoqueRestante: totalMlGramas
        });

        await workbook.xlsx.writeFile(CAMINHO_PLANILHA);
      });

      res.json({ mensagem: 'Produto cadastrado!', produtoId: novoProdutoId });
    }
  );
});

// Reabastecer Produto Existente
app.post('/api/reabastecer', autenticado, (req, res) => {
  const { produtoId, qtdAdicionada, unidade, preco, tipoPreco } = req.body;

  const mlGramasNovos = calcularMenorUnidade(qtdAdicionada, unidade);
  const qtdNum = parseFloat(qtdAdicionada) || 1;
  const valorNum = parseFloat(preco) || 0;

  const precoPagoCalculado = (tipoPreco === 'unitario') ? (valorNum * qtdNum) : valorNum;

  db.get(`SELECT * FROM produtos WHERE id = ?`, [produtoId], (err, produto) => {
    if (err || !produto) return res.status(404).json({ erro: 'Produto não encontrado.' });

    const novoEstoqueTotal = produto.estoque_restante + mlGramasNovos;
    const novoPrecoAcumulado = produto.preco_total + precoPagoCalculado;
    const novoTotalMlGComprado = produto.total_ml_g + mlGramasNovos;
    
    const novoCustoUnitario = novoPrecoAcumulado / novoTotalMlGComprado;

    db.run(
      `UPDATE produtos SET 
        estoque_restante = ?, 
        preco_total = ?, 
        total_ml_g = ?, 
        custo_unitario = ? 
       WHERE id = ?`,
      [novoEstoqueTotal, novoPrecoAcumulado, novoTotalMlGComprado, novoCustoUnitario, produtoId],
      (err) => {
        if (err) return res.status(500).json({ erro: 'Erro ao reabastecer estoque.' });

        executarNoExcel(async () => {
          const workbook = new ExcelJS.Workbook();
          await workbook.xlsx.readFile(CAMINHO_PLANILHA);
          const abaCompras = workbook.getWorksheet('Compras e Estoque');

          abaCompras.eachRow((row, rowNumber) => {
            if (rowNumber > 1 && String(row.getCell(1).value) === String(produtoId)) {
              row.getCell(5).value = novoTotalMlGComprado;
              row.getCell(6).value = novoPrecoAcumulado;
              row.getCell(7).value = novoCustoUnitario;
              row.getCell(8).value = novoEstoqueTotal;
            }
          });

          await workbook.xlsx.writeFile(CAMINHO_PLANILHA);
        });

        res.json({ mensagem: 'Estoque reabastecido com sucesso!', produtoId: produtoId });
      }
    );
  });
});

// Listar Produtos
app.get('/api/produtos', autenticado, (req, res) => {
  db.all(`SELECT id, nome, unidade, estoque_restante AS estoque, custo_unitario AS custoUnitario FROM produtos`, [], (err, rows) => {
    if (err) return res.status(500).json({ erro: 'Erro ao carregar produtos.' });
    res.json(rows);
  });
});

// Excluir Produto do Estoque
app.delete('/api/produtos/:id', autenticado, (req, res) => {
  const produtoId = req.params.id;

  db.run(`DELETE FROM produtos WHERE id = ?`, [produtoId], function(err) {
    if (err) return res.status(500).json({ erro: 'Erro ao excluir produto.' });

    executarNoExcel(async () => {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(CAMINHO_PLANILHA);
      const abaCompras = workbook.getWorksheet('Compras e Estoque');

      if (abaCompras) {
        let linhaParaRemover = null;
        abaCompras.eachRow((row, rowNumber) => {
          if (rowNumber > 1 && String(row.getCell(1).value) === String(produtoId)) {
            linhaParaRemover = rowNumber;
          }
        });

        if (linhaParaRemover) {
          abaCompras.spliceRows(linhaParaRemover, 1);
          await workbook.xlsx.writeFile(CAMINHO_PLANILHA);
        }
      }
    });

    res.json({ mensagem: 'Produto excluído com sucesso!' });
  });
});

// Registrar Manejo
app.post('/api/manejos', autenticado, (req, res) => {
  const { produtoId, talhao, qtdGastaMlG, data } = req.body;
  const qtdManejo = parseFloat(qtdGastaMlG);
  const dataFormatada = data || new Date().toISOString().split('T')[0];

  db.get(`SELECT * FROM produtos WHERE id = ?`, [produtoId], (err, produto) => {
    if (err || !produto) return res.status(404).json({ erro: 'Produto não encontrado.' });

    if (produto.estoque_restante < qtdManejo) {
      return res.status(400).json({ erro: 'Estoque insuficiente.' });
    }

    const custoManejo = qtdManejo * produto.custo_unitario;
    const novoEstoque = produto.estoque_restante - qtdManejo;

    db.run(`UPDATE produtos SET estoque_restante = ? WHERE id = ?`, [novoEstoque, produtoId], (err) => {
      if (err) return res.status(500).json({ erro: 'Erro ao atualizar estoque.' });

      db.run(
        `INSERT INTO manejos (data, talhao, produto_nome, qtd_gasta, custo_manejo, usuario) VALUES (?, ?, ?, ?, ?, ?)`,
        [dataFormatada, talhao, produto.nome, qtdManejo, custoManejo, req.session.usuario],
        function (err) {
          if (err) return res.status(500).json({ erro: 'Erro ao registrar manejo.' });

          executarNoExcel(async () => {
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.readFile(CAMINHO_PLANILHA);
            const abaCompras = workbook.getWorksheet('Compras e Estoque');
            const abaManejos = workbook.getWorksheet('Manejos e Aplicacoes');

            if (abaCompras) {
              abaCompras.eachRow((row, rowNumber) => {
                if (rowNumber > 1 && String(row.getCell(1).value) === String(produtoId)) {
                  row.getCell(8).value = novoEstoque;
                }
              });
            }

            if (abaManejos) {
              abaManejos.addRow({
                data: dataFormatada,
                talhao: talhao,
                produto: produto.nome,
                qtdGasta: qtdManejo,
                custoManejo: custoManejo,
                usuario: req.session.usuario
              });
            }

            await workbook.xlsx.writeFile(CAMINHO_PLANILHA);
          });

          res.json({ mensagem: 'Manejo registrado!', custoManejo: custoManejo });
        }
      );
    });
  });
});

// Excluir Lançamento de Manejo (Devolve a quantidade gasta ao Estoque)
app.delete('/api/manejos/:id', autenticado, (req, res) => {
  const manejoId = req.params.id;

  db.get(`SELECT * FROM manejos WHERE id = ?`, [manejoId], (err, manejo) => {
    if (err || !manejo) return res.status(404).json({ erro: 'Lançamento de manejo não encontrado.' });

    db.run(
      `UPDATE produtos SET estoque_restante = estoque_restante + ? WHERE nome = ?`,
      [manejo.qtd_gasta, manejo.produto_nome],
      (err) => {
        if (err) console.error('Erro ao devolver estoque:', err);

        db.run(`DELETE FROM manejos WHERE id = ?`, [manejoId], function(err) {
          if (err) return res.status(500).json({ erro: 'Erro ao excluir manejo.' });

          executarNoExcel(async () => {
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.readFile(CAMINHO_PLANILHA);
            const abaManejos = workbook.getWorksheet('Manejos e Aplicacoes');

            if (abaManejos) {
              let linhaParaRemover = null;
              abaManejos.eachRow((row, rowNumber) => {
                if (
                  rowNumber > 1 &&
                  String(row.getCell(1).value) === String(manejo.data) &&
                  String(row.getCell(3).value) === String(manejo.produto_nome) &&
                  parseFloat(row.getCell(4).value) === parseFloat(manejo.qtd_gasta)
                ) {
                  linhaParaRemover = rowNumber;
                }
              });

              if (linhaParaRemover) {
                abaManejos.spliceRows(linhaParaRemover, 1);
                await workbook.xlsx.writeFile(CAMINHO_PLANILHA);
              }
            }
          });

          res.json({ mensagem: 'Lançamento excluído e quantidade devolvida ao estoque!' });
        });
      }
    );
  });
});

// Listar Talhões Cadastrados (Únicos)
app.get('/api/talhoes', autenticado, (req, res) => {
  db.all(`SELECT DISTINCT talhao FROM manejos ORDER BY talhao ASC`, [], (err, rows) => {
    if (err) return res.status(500).json({ erro: 'Erro ao buscar talhões.' });
    res.json(rows.map(r => r.talhao));
  });
});

// Relatório do Mês (Com Suporte a Filtro por Talhão)
app.get('/api/relatorio/mensal', autenticado, (req, res) => {
  const { anoMes, talhao } = req.query;

  let sql = `SELECT * FROM manejos WHERE data LIKE ?`;
  let params = [`${anoMes}%`];

  if (talhao) {
    sql += ` AND talhao = ?`;
    params.push(talhao);
  }

  sql += ` ORDER BY data DESC`;

  db.all(sql, params, (err, rows) => {
    if (err) return res.status(500).json({ erro: 'Erro ao buscar relatório.' });

    let totalGastoMes = 0;
    const historico = rows.map(r => {
      totalGastoMes += parseFloat(r.custo_manejo) || 0;
      return {
        id: r.id,
        data: r.data,
        talhao: r.talhao,
        produto: r.produto_nome,
        qtdGasta: r.qtd_gasta,
        custo: r.custo_manejo,
        usuario: r.usuario
      };
    });

    res.json({ totalGastoMes, historico });
  });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Servidor rodando em http://localhost:${PORT}`);
});