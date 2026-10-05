document.addEventListener('DOMContentLoaded', () => {
  const hoje = new Date();
  document.getElementById('m-data').value = hoje.toISOString().split('T')[0];
  document.getElementById('filtro-mes').value = `${hoje.getFullYear()}-${String(hoje.getMonth() + 1).padStart(2, '0')}`;
  
  document.getElementById('filtro-mes').addEventListener('change', buscarRelatorio);
  document.getElementById('filtro-talhao').addEventListener('change', buscarRelatorio);

  checarSessao();
});

function alternarAbaCompra(aba) {
  const btnNovo = document.getElementById('btn-tab-novo');
  const btnReabastecer = document.getElementById('btn-tab-reabastecer');
  const formNovo = document.getElementById('form-compra-nova');
  const formReabastecer = document.getElementById('form-reabastecer');

  if (aba === 'novo') {
    btnNovo.classList.add('active');
    btnReabastecer.classList.remove('active');
    formNovo.classList.remove('hidden');
    formReabastecer.classList.add('hidden');
  } else {
    btnReabastecer.classList.add('active');
    btnNovo.classList.remove('active');
    formReabastecer.classList.remove('hidden');
    formNovo.classList.add('hidden');
  }
}

async function checarSessao() {
  const res = await fetch('/api/sessao');
  const data = await res.json();

  if (data.usuario) {
    document.getElementById('tela-login').classList.add('hidden');
    document.getElementById('tela-app').classList.remove('hidden');
    document.getElementById('status-usuario').innerHTML = `
      Usuário: <b>${data.usuario}</b> | 
      <button onclick="logout()" style="width:auto; padding:5px 10px; background:#c62828;">Sair</button>
    `;
    carregarProdutos();
    carregarTalhoes();
    buscarRelatorio();
  } else {
    document.getElementById('tela-login').classList.remove('hidden');
    document.getElementById('tela-app').classList.add('hidden');
    document.getElementById('status-usuario').innerHTML = '';
  }
}

// Login
document.getElementById('form-login').addEventListener('submit', async (e) => {
  e.preventDefault();
  const res = await fetch('/api/login', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      usuario: document.getElementById('login-user').value,
      senha: document.getElementById('login-pass').value
    })
  });
  const data = await res.json();
  if (res.ok) checarSessao();
  else alert(data.erro);
});

async function logout() {
  await fetch('/api/logout', { method: 'POST' });
  checarSessao();
}

// Alterar Senha
document.getElementById('form-alterar-senha').addEventListener('submit', async (e) => {
  e.preventDefault();
  const res = await fetch('/api/alterar-senha', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      senhaAtual: document.getElementById('senha-atual').value,
      novaSenha: document.getElementById('nova-senha').value
    })
  });
  const data = await res.json();
  if (res.ok) {
    alert(data.mensagem);
    e.target.reset();
  } else {
    alert(data.erro || 'Erro ao alterar a senha.');
  }
});

// Baixar Planilha Excel
function baixarExcel() {
  window.location.href = '/api/download-excel';
}

// Cadastrar Novo Produto
document.getElementById('form-compra-nova').addEventListener('submit', async (e) => {
  e.preventDefault();
  const res = await fetch('/api/compras', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      nome: document.getElementById('c-nome').value,
      qtdComprada: document.getElementById('c-qtd').value,
      unidade: document.getElementById('c-unidade').value,
      tipoPreco: document.getElementById('c-tipo-preco').value,
      preco: document.getElementById('c-preco').value
    })
  });
  
  const data = await res.json();

  if (res.ok) {
    alert('Novo produto cadastrado no estoque!');
    e.target.reset();
    await carregarProdutos(data.produtoId);

    const select = document.getElementById('m-produto');
    select.scrollIntoView({ behavior: 'smooth' });
    select.focus();
  } else alert(data.erro || 'Erro ao cadastrar.');
});

// Reabastecer Produto Existente
document.getElementById('form-reabastecer').addEventListener('submit', async (e) => {
  e.preventDefault();
  const res = await fetch('/api/reabastecer', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      produtoId: document.getElementById('r-produto').value,
      qtdAdicionada: document.getElementById('r-qtd').value,
      unidade: document.getElementById('r-unidade').value,
      tipoPreco: document.getElementById('r-tipo-preco').value,
      preco: document.getElementById('r-preco').value
    })
  });

  const data = await res.json();

  if (res.ok) {
    alert('Estoque reabastecido com sucesso!');
    e.target.reset();
    await carregarProdutos(data.produtoId);

    const select = document.getElementById('m-produto');
    select.scrollIntoView({ behavior: 'smooth' });
    select.focus();
  } else alert(data.erro || 'Erro ao reabastecer.');
});

// Traduz a unidade para exibição amigável ao usuário
function obterRotuloUnidade(unidadeOriginal) {
  const u = (unidadeOriginal || '').toLowerCase();
  
  if (u === 'kg' || u === 't') return 'g';
  if (u === 'l' || u === 'm3') return 'ml';
  if (u === 'saca_sementes' || u === 'sementes') return 'sementes';
  if (u === 'pacote') return 'pct';
  if (u === 'caixa') return 'cx';
  if (u === 'unidade') return 'un';
  
  return unidadeOriginal;
}

// Carregar Produtos nos Selects e Tabela de Gerenciamento
async function carregarProdutos(produtoIdParaSelecionar = null) {
  try {
    const res = await fetch('/api/produtos');
    const produtos = await res.json();
    
    const selectManejo = document.getElementById('m-produto');
    const selectReabastecer = document.getElementById('r-produto');
    const tbodyGerenciador = document.getElementById('tabela-produtos-gerenciador');
    
    let optionsManejo = '<option value="">Selecione o Produto...</option>';
    let optionsReabastecer = '<option value="">Selecione o Produto para Reabastecer...</option>';
    let linhasTabela = '';

    if (Array.isArray(produtos) && produtos.length > 0) {
      produtos.forEach(p => {
        const rotuloUnidade = obterRotuloUnidade(p.unidade);

        const option = `<option value="${p.id}">${p.nome} (Estoque: ${p.estoque} ${rotuloUnidade})</option>`;
        optionsManejo += option;
        optionsReabastecer += option;

        linhasTabela += `
          <tr>
            <td>${p.id}</td>
            <td><b>${p.nome}</b></td>
            <td>${p.estoque} ${rotuloUnidade}</td>
            <td>
              <button class="btn-danger" onclick="excluirProduto(${p.id}, '${p.nome}')">Excluir</button>
            </td>
          </tr>
        `;
      });

      selectManejo.innerHTML = optionsManejo;
      selectReabastecer.innerHTML = optionsReabastecer;
      tbodyGerenciador.innerHTML = linhasTabela;

      if (produtoIdParaSelecionar) {
        selectManejo.value = String(produtoIdParaSelecionar);
      }
    } else {
      selectManejo.innerHTML = optionsManejo;
      selectReabastecer.innerHTML = optionsReabastecer;
      tbodyGerenciador.innerHTML = '<tr><td colspan="4" style="text-align:center;">Nenhum produto em estoque.</td></tr>';
    }
  } catch (err) {
    console.error('Erro ao buscar produtos:', err);
  }
}

// Carregar Talhões para o filtro de relatório
async function carregarTalhoes() {
  try {
    const res = await fetch('/api/talhoes');
    const talhoes = await res.json();
    const selectFiltro = document.getElementById('filtro-talhao');
    const valorAtual = selectFiltro.value;

    let options = '<option value="">Todos os Talhões</option>';
    if (Array.isArray(talhoes)) {
      talhoes.forEach(t => {
        options += `<option value="${t}">${t}</option>`;
      });
    }
    selectFiltro.innerHTML = options;
    selectFiltro.value = valorAtual;
  } catch (err) {
    console.error('Erro ao buscar talhões:', err);
  }
}

// Excluir Produto
async function excluirProduto(id, nome) {
  if (confirm(`Tem certeza que deseja excluir o produto "${nome}" do estoque?`)) {
    const res = await fetch(`/api/produtos/${id}`, { method: 'DELETE' });
    const data = await res.json();
    if (res.ok) {
      alert(data.mensagem);
      carregarProdutos();
    } else alert(data.erro || 'Erro ao excluir produto.');
  }
}

// Excluir Manejo (Devolve o produto ao Estoque)
async function excluirManejo(id) {
  if (confirm('Tem certeza que deseja apagar este lançamento? A quantidade gasta será devolvida ao estoque.')) {
    const res = await fetch(`/api/manejos/${id}`, { method: 'DELETE' });
    const data = await res.json();
    if (res.ok) {
      alert(data.mensagem);
      carregarProdutos();
      carregarTalhoes();
      buscarRelatorio();
    } else alert(data.erro || 'Erro ao excluir manejo.');
  }
}

// Lançar Manejo
document.getElementById('form-manejo').addEventListener('submit', async (e) => {
  e.preventDefault();
  
  const dataPreenchida = document.getElementById('m-data').value;

  const res = await fetch('/api/manejos', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      produtoId: document.getElementById('m-produto').value,
      talhao: document.getElementById('m-talhao').value,
      qtdGastaMlG: document.getElementById('m-qtd').value,
      data: dataPreenchida
    })
  });
  
  const data = await res.json();
  if (res.ok) {
    alert(`Manejo Lançado! Custo: R$ ${data.custoManejo.toFixed(2)}`);
    e.target.reset();
    
    document.getElementById('m-data').value = dataPreenchida;
    
    carregarProdutos();
    carregarTalhoes();
    buscarRelatorio();
  } else alert(data.erro);
});

// Buscar Relatório do Mês e/ou Por Talhão
async function buscarRelatorio() {
  const mes = document.getElementById('filtro-mes').value;
  const talhao = document.getElementById('filtro-talhao').value;
  if (!mes) return;

  let url = `/api/relatorio/mensal?anoMes=${mes}`;
  if (talhao) {
    url += `&talhao=${encodeURIComponent(talhao)}`;
  }

  const res = await fetch(url);
  const data = await res.json();

  const totalGasto = parseFloat(data.totalGastoMes) || 0;
  const textoTalhao = talhao ? ` no ${talhao}` : '';
  
  document.getElementById('total-mes').innerText = `Total Gasto${textoTalhao}: R$ ${totalGasto.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  
  const tbody = document.getElementById('tabela-relatorio');

  if (data.historico && data.historico.length > 0) {
    let linhasHtml = '';
    data.historico.forEach(h => {
      const custoFormatado = (parseFloat(h.custo) || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      linhasHtml += `
        <tr>
          <td>${h.data}</td>
          <td><b>${h.talhao}</b></td>
          <td>${h.produto}</td>
          <td>${h.qtdGasta}</td>
          <td>R$ ${custoFormatado}</td>
          <td>
            <button class="btn-danger" onclick="excluirManejo(${h.id})">Excluir</button>
          </td>
        </tr>
      `;
    });
    tbody.innerHTML = linhasHtml;
  } else {
    tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;">Nenhum manejo registrado para estes filtros.</td></tr>';
  }
}