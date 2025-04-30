let dadosOriginais = [];
      
fetch('emails.xlsx')
    .then(response => response.arrayBuffer())
    .then(data => {
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        dadosOriginais = jsonData;
        renderizarTabela(jsonData);
  })
  .catch(error => {
        console.error('Erro ao carregar o arquivo Excel:', error);
  });

function renderizarTabela(dados) {
    const tbody = document.querySelector('#tabela-dados tbody');
    tbody.innerHTML = '';
    dados.forEach(row => {
        const tr = document.createElement('tr');
        const tdNome = document.createElement('td');
        tdNome.textContent = row.NOME || '';
        const tdEmail = document.createElement('td');
        tdEmail.textContent = row.EMAIL || '';
        tr.appendChild(tdNome);
        tr.appendChild(tdEmail);
        tbody.appendChild(tr);
    });
}

document.getElementById('busca-nome').addEventListener('input', function () {
    const termo = this.value.toLowerCase();
    const filtrado = dadosOriginais.filter(item =>
        item.NOME && item.NOME.toLowerCase().includes(termo)
    );
    renderizarTabela(filtrado);
});