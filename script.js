const botao = document.querySelector('.botao-carrinho');
botao.addEventListener('click', () => {
    window.scrollTo({
    top: document.body.scrollHeight,
    behavior: 'smooth'
  });
});
const btnFinalizar = document.getElementById('btnFinalizar');
const modal = document.getElementById('modalCheckout');
const btnFechar = document.getElementById('btnFechar');
const formCheckout = document.getElementById('formCheckout');

const livros = [
     {
        id: 1,
        autor: 'Tolkien',
        titulo: 'O Hobbit',
        preco: 45.50,
        genero: 'fantasia',
        capa: "https://m.media-amazon.com/images/I/511+-lOOtsL._SY445_SX342_ML2_.jpg",
        sinopse: '"O Hobbit" acompanha Bilbo Bolseiro, um pacato hobbit recrutado pelo mago Gandalf e por treze anões para uma perigosa missão de resgate de um tesouro ancestral. Em uma jornada ele tera muitos desafios.'
 },
 {
        id: 2,
        autor: 'Marco Túlio',
        titulo: 'Authenticgames: A batalha da Torre',
        preco: 45.50,
        genero: 'aventura',
        capa: "https://m.media-amazon.com/images/I/81D5LriGuWL._SL1500_.jpg",
        sinopse: '"Um sequestro misterioso coloca o nosso herói AUTHENTICGAMES em uma grande enrascada, e ele vai precisar de ajuda para sair dessa.'
 },
  {
        id: 10,
        titulo: "Authenticgames: A batalha contra Herobrine",
        autor: "Marco Túlio",
        genero: "ficção de ação e aventura",
        preco: 49.90,
        capa: "https://m.media-amazon.com/images/I/91y8WC8ACTL._SL1500_.jpg",
        sinopse: "Depois de ser resgatado de um terrível sequestro, AUTHENTICGAMES está de volta para proteger a Vila Farmer. Mas antes ele precisará encarar uma nova aventura para recuperar sua espada de diamante.",
        ano: 2016,
        paginas: 112,
        estrelas: 5
 },
  {
        id: 11,
        titulo: "Authenticgames: A batalha Contra o Ender Dragon",
        autor: "Marco Túlio",
        genero: "ficção de ação e aventura",
        preco: 49.90,
        capa: "https://m.media-amazon.com/images/I/81SEuYohRXL._SL1500_.jpg",
        sinopse: "No último livro da trilogia AuthenticGames, Builder, Nina e Authentic vão até a sombria dimensão do The End para enfrentar novos desafios e salvar o mundo da superfície das terríveis ameaças do Ender Dragon.",
        ano: 2017,
        paginas: 112,
        estrelas: 5
  },
  {
        id: 12,
        titulo: "Viagem à Alma Tricolor em 7 Epopeias",
        autor: "Léo Gerchmann",
        genero: "Crônica Esportiva",
        preco: 35.00,
        capa: "https://m.media-amazon.com/images/I/81qapBfiIwL._SL1500_.jpg",
        sinopse: "O livro que você tem em mãos é uma assumida ousadia. Pretende desvendar os escaninhos de uma paixão.",
        ano: 2016,
        paginas: 125,
        estrelas: 5
  },
  {
        id: 15,
        titulo: "I Have No Mouth and I Must Scream: Stories",
        autor: "Harlan Ellison",
        genero: "Ficção Cientifica e Terror Psicológico",
        preco: 2064.32,
        capa: "https://m.media-amazon.com/images/I/91ew3HlZmXL._SL1500_.jpg",
        sinopse: "A história-título pós-apocalíptica, vencedora do prêmio Hugo, desta coletânea de sete contos e a história que encerra o volume, 'Pretty Maggie Moneyeyes.",
        ano: 2014,
        paginas: 162,
        estrelas: 4
  },
  {
        id: 17,
        titulo: "Authenticgames: vivendo uma vida autêntica",
        autor: "Marco Túlio",
        genero: "Autobiografia",
        preco: 49.90,
        capa: "https://m.media-amazon.com/images/I/81qSgdApcTL._SL1500_.jpg",
        sinopse: "O mineiro Marco Túlio sempre foi apaixonado por games. Tão apaixonado que decidiu enfrentar a timidez e criar um canal no YouTube para falar dos jogos de que gostava. Com seu jeito simples e engraçado, Marco Túlio transformou o AuthenticGames em ponto de encontro para quase 4 milhões de crianças e adolescentes. É lá que eles trocam ideias, aprendem estratégias secretas sobre Minecraft e acompanham as séries exclusivas. Neste livro, os fãs vão saber como surgiu o projeto do canal, quem são os amigos da internet que o Authentic levou para a vida real e muito mais! Um dos youtubers mais amados do Brasil conta todos os seus segredos. Mais de 1 bilhão de visualizações!",
        ano: 2016,
        paginas: 160,
        estrelas: 5
  },
  {
        id: 19,
        titulo: "Como Conversar com Seu Cachorro Sobre Homossexualidade e Comunismo",
        autor: "Alexander Blythe",
        genero: "Ficção Humorística e Sátira",
        preco: 5.99,
        capa: "https://m.media-amazon.com/images/I/71hTp3F2VQL._SL1500_.jpg",
        sinopse: "Você será guiado por um percurso singular, onde os diálogos habituais entre tutor e pet são transformados em conversas significativas sobre temas contemporâneos e sociais.",
        ano: 2024,
        paginas: 38,
        estrelas: 4
  },
  {
        id: 22,
        titulo: "Fundamentos do xadrez",
        autor: "José Raul Capablanca",
        genero: "Esporte e Estratégia",
        preco: 59.90,
        capa: "https://m.media-amazon.com/images/I/71eTThl+2oL._SY466_.jpg",
        sinopse: "Um guia completo para aprender os fundamentos do xadrez.",
        ano: 1921,
        paginas: 224,
        estrelas: 4
  },
  {
        id: 38,
        titulo: "A fúria dos mobs: 1",
        autor: "Spok",
        genero: "ficção de ação e aventura",
        preco: 25.90,
        capa: "https://m.media-amazon.com/images/I/81OSY3bKtvL._SY466_.jpg",
        sinopse: "Spok acordou no mundo fantástico de Minecraft. O susto de estar numa dimensão toda quadrada e esquisita não é tudo… Logo Spok irá encontrar um bando de monstros terríveis que querem acabar com ele. Mas ele não está só. Neste surpreendente “A fúria dos Mobs” ele conta com seus amigos Pac, Mike, Authentic, Toddynho, Moonkase, Cauê, Likea, Nofaxu, Malena, Jabuti e Jazz. Todos são personagens e todos correm dos inimigos Pedro e Maya, a aranha, entre outros seres aterrorizantes. E tem uma coisa ainda pior! Algo muito estranho está acontecendo com esses monstros, que sempre foram meio abobalhados… Eles não querer acabar só com Spok e seus amigos, e com uns poucos aldeões indefesos. Eles querem aprisionar todos os seus migos numa fortaleza sombria do Nether. Para sempre! Eles querem o mundo de Minecraft só para eles. Mas os Mobs não vão conseguir isso tão fácil. Numa pequena aldeia, onde todos os amigos de Spok moram, eles vão resistir e lutar contra esse terrível exército. É tudo ou nada! Embarque nessa aventura eletrizante",
        ano: 2016,
        paginas: 148,
        estrelas: 5
  },
  {
        id: 41,
        titulo: "De volta ao jogo: Uma Aventura Não Oficial De Minecraft",
        autor: "RezendeEvil",
        genero: "ficção de ação e aventura",
        preco: 25.90,
        capa: "https://m.media-amazon.com/images/I/81d8WWA0PfS._SL1500_.jpg",
        sinopse: "De volta ao jogo é uma aventura não oficial de Minecraft, escrita por RezendeEvil. A história segue o personagem principal, que é transportado para o mundo de Minecraft e deve enfrentar desafios e inimigos para sobreviver e encontrar uma maneira de voltar para casa.",
        ano: 2016,
        paginas: 128,
        estrelas: 5
  },
];

const lista = document.getElementById('lista-livros');

for (const livro of livros) {
  lista.innerHTML += `
    <article>
                  <img src="${livro.capa}" alt="imagem livro">
            <div>
                <h1>${livro.titulo}</h1>
                <p class="artigo-autor">Genero: ${livro.genero}</p>
                <p class="artigo-autor">Autor: ${livro.autor}</p>
                <p>${livro.sinopse}</p>
                <button data-id="${livro.id}">Comprar</button>
                <button data-detalhes="${livro.id}">detalhes</button>
                <h1>R$${livro.preco}</h1>
            </div>
    </article>`;
}

const carrinho = [];

const area =  document.getElementById('lista-livros');


function adicionarAoCarrinho(id) {
  const existente = carrinho.find(item => item.id === id);
  if (existente) {
    existente.qtd++;
  } else {
    carrinho.push({id: id, qtd: 1});
  }
  renderCarrinho();
}

lista.addEventListener('click', (e) => {
  // botão "detalhes": abre a janela com a descrição do livro
  const idDetalhes = Number(e.target.dataset.detalhes);
  if (idDetalhes) {
    abrirDetalhes(idDetalhes);
    return;
  }

  // botão "Comprar": ignora cliques que não vieram de um botão com data-id
  const id = Number(e.target.dataset.id);
  if (!id) return;
  adicionarAoCarrinho(id);
});

// ---------- Detalhes do livro (modal) ----------
const modalDetalhes = document.getElementById('modalDetalhes');
const conteudoDetalhes = document.getElementById('conteudoDetalhes');

function abrirDetalhes(id) {
  const livro = livros.find(l => l.id === id);
  if (!livro) return;

  // ano, páginas e estrelas só existem em alguns livros
  const ano = livro.ano ? `<p><strong>Ano:</strong> ${livro.ano}</p>` : '';
  const paginas = livro.paginas ? `<p><strong>Páginas:</strong> ${livro.paginas}</p>` : '';
  const estrelas = livro.estrelas ? `<p><strong>Avaliação:</strong> ${'★'.repeat(livro.estrelas)}</p>` : '';

  conteudoDetalhes.innerHTML = `
    <img src="${livro.capa}" alt="Capa de ${livro.titulo}" class="detalhes-capa">
    <h2>${livro.titulo}</h2>
    <p><strong>Autor:</strong> ${livro.autor}</p>
    <p><strong>Gênero:</strong> ${livro.genero}</p>
    ${ano}${paginas}${estrelas}
    <p>${livro.sinopse}</p>
    <p class="detalhes-preco">R$ ${livro.preco}</p>
    <button type="button" class="btn-confirmar" data-comprar-detalhes="${livro.id}">Comprar</button>`;

  modalDetalhes.hidden = false;
}

modalDetalhes.addEventListener('click', (e) => {
  // fundo escuro ou botão ✕ fecham a janela
  if (e.target === modalDetalhes || e.target.id === 'btnFecharDetalhes') {
    modalDetalhes.hidden = true;
    return;
  }
  // botão Comprar dentro da janela
  const idCompra = Number(e.target.dataset.comprarDetalhes);
  if (idCompra) {
    adicionarAoCarrinho(idCompra);
    modalDetalhes.hidden = true;
  }
});

// Botões do carrinho: +, - e Remover (um listener só, no local-carrinho)
document.getElementById('local-carrinho').addEventListener('click', (e) => {
  const id = Number(e.target.dataset.id);
  const acao = e.target.dataset.acao;
  const item = carrinho.find(i => i.id === id);
  if (!item) return; // clicou fora de um botão

  if (acao === 'mais') {
    item.qtd++;
  } else if (acao === 'menos') {
    item.qtd--;
  } else if (acao === 'remover') {
    item.qtd = 0;
  }

  // quantidade zerada: tira o item do carrinho
  if (item.qtd <= 0) {
    carrinho.splice(carrinho.indexOf(item), 1);
  }

  renderCarrinho();
});

function renderCarrinho() {
  const localdocarrinho = document.getElementById('local-carrinho');
  localdocarrinho.innerHTML = '';

  let subtotal = 0;

  for (const item of carrinho) {
    const livro = livros.find(f => f.id === item.id);
    subtotal += livro.preco * item.qtd;

    localdocarrinho.innerHTML += `
      <div class="item-carrinho">
        <img src="${livro.capa}" alt="Capa do livro" class="img-thumb">
        <div class="detalhes-item">
          <h3>${livro.titulo}</h3>
          <p class="preco-unitario">R$ ${livro.preco}</p>
          <div class="quantidade">
            <button data-id="${item.id}" data-acao="menos">-</button>
            <span>${item.qtd}</span>
            <button data-id="${item.id}" data-acao="mais">+</button>
          </div>
        </div>
        <button class="btn-remover" data-id="${item.id}" data-acao="remover">Remover</button>
      </div>`;
  }

  // frete fixo de R$ 10, só se houver algo no carrinho
  const frete = carrinho.length > 0 ? 10 : 0;
  const total = subtotal + frete;

  document.getElementById('subtotal').textContent = subtotal;
  document.getElementById('frete').textContent = frete;
  document.getElementById('total').textContent = total;
}

renderCarrinho();

// ---------- Finalizar compra (modal) ----------
const resumoCompra = document.getElementById('resumoCompra');
const mensagemSucesso = document.getElementById('mensagemSucesso');

btnFinalizar.addEventListener('click', () => {
  if (carrinho.length === 0) {
    alert('Seu carrinho está vazio. Adicione um livro primeiro!');
    return;
  }

  // resumo do pedido: uma linha por livro + total
  let linhas = '';
  for (const item of carrinho) {
    const livro = livros.find(l => l.id === item.id);
    linhas += `<li>${item.qtd}x ${livro.titulo} — R$ ${livro.preco * item.qtd}</li>`;
  }
  resumoCompra.innerHTML = `
    <ul>${linhas}</ul>
    <p><strong>Total com frete: R$ ${document.getElementById('total').textContent}</strong></p>`;

  formCheckout.hidden = false;
  mensagemSucesso.hidden = true;
  modal.hidden = false;
});

btnFechar.addEventListener('click', () => {
  modal.hidden = true;
});

// clicar no fundo escuro também fecha
modal.addEventListener('click', (e) => {
  if (e.target === modal) modal.hidden = true;
});

// URL do Google Apps Script (termina em /exec). Veja apps-script.gs.
// Vazia = o pedido só é simulado na tela, nada é gravado.
const URL_PLANILHA = 'https://script.google.com/macros/s/AKfycby-sX5lrW3o-4ote9E_oOF_49eG8VzTABPEd8gBZU533uTk2Ez7BkUu8I52QrKsyS52/exec';

formCheckout.addEventListener('submit', async (e) => {
  e.preventDefault(); // não recarrega a página
  const nome = formCheckout.elements.nome.value;

  if (URL_PLANILHA) {
    const botaoConfirmar = formCheckout.querySelector('button[type="submit"]');
    botaoConfirmar.disabled = true;
    botaoConfirmar.textContent = 'Enviando...';

    // texto dos livros ("2x O Hobbit; 1x ...") e valor total com frete
    const livrosTexto = carrinho.map(item => {
      const livro = livros.find(l => l.id === item.id);
      return `${item.qtd}x ${livro.titulo}`;
    }).join('; ');
    const valor = Math.round(Number(document.getElementById('total').textContent) * 100) / 100;

    try {
      // corpo em texto simples: evita o "preflight" do navegador (CORS)
      const resposta = await fetch(URL_PLANILHA, {
        method: 'POST',
        body: JSON.stringify({ nome: nome, livro: livrosTexto, valor: valor })
      });
      const resultado = await resposta.json();
      if (!resultado.ok) throw new Error(resultado.erro);
    } catch (erro) {
      mensagemSucesso.textContent = 'Não foi possível registrar o pedido. Tente novamente.';
      mensagemSucesso.hidden = false;
      return;
    } finally {
      botaoConfirmar.disabled = false;
      botaoConfirmar.textContent = 'Confirmar pedido';
    }
  }

  mensagemSucesso.textContent = `Obrigado, ${nome}! Seu pedido foi confirmado.`;
  mensagemSucesso.hidden = false;
  formCheckout.hidden = true;
  resumoCompra.innerHTML = '';

  carrinho.length = 0; // esvazia o carrinho
  renderCarrinho();
  formCheckout.reset();
});
