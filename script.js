function processarExcel() {
    const fileInput = document.getElementById("excelFile");
    
    if (!fileInput.files || fileInput.files.length === 0) {
        alert("Por favor, selecione o arquivo do relatório primeiro.");
        return;
    }
    
    // CASO CORRIGIDO: Agora pegamos o arquivo correto usando [0]
    const arquivo = fileInput.files[0]; 
    const extensao = arquivo.name.split('.').pop().toLowerCase();
    const reader = new FileReader();
    const expressaoRegex = /\b(DB\d{8}|DC\d{8}|F\d{8}|\d{8})\b/g;

    if (extensao === 'xlsx' || extensao === 'xls') {
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                let pedidosExcel = [];

                workbook.SheetNames.forEach(sheetName => {
                    const worksheet = workbook.Sheets[sheetName];
                    const csvTexto = XLSX.utils.sheet_to_csv(worksheet);
                    
                    let correspondencia;
                    while ((correspondencia = expressaoRegex.exec(csvTexto)) !== null) {
                        const codigoIdentificado = correspondencia[0]; // Captura o texto exato
                        if (!pedidosExcel.includes(codigoIdentificado) && codigoIdentificado !== '24004716') {
                            pedidosExcel.push(codigoIdentificado);
                        }
                    }
                });
                gerarTabelaComparacao(pedidosExcel);
            } catch (erro) {
                console.error(erro);
                alert("Erro ao decodificar a planilha Excel: " + erro.message);
            }
        };
        reader.readAsArrayBuffer(arquivo);
    } else {
        reader.onload = function(e) {
            try {
                const textoPlano = e.target.result;
                let pedidosTexto = [];
                let correspondencia;
                while ((correspondencia = expressaoRegex.exec(textoPlano)) !== null) {
                    const codigoIdentificado = correspondencia[0]; // Captura o texto exato
                    if (!pedidosTexto.includes(codigoIdentificado) && codigoIdentificado !== '24004716') {
                        pedidosTexto.push(codigoIdentificado);
                    }
                }
                if (pedidosTexto.length === 0) {
                    const linhas = textoPlano.split(/\r?\n/);
                    for (let i = 1; i < linhas.length; i++) {
                        const colunas = linhas[i].split(/[,;]/);
                        if (colunas && colunas.length > 0) {
                            const ped = limparValor(colunas[0]);
                            if (ped && ped.length >= 4 && !ped.startsWith('Ped Cli')) {
                                if (!pedidosTexto.includes(ped)) pedidosTexto.push(ped);
                            }
                        }
                    }
                }
                gerarTabelaComparacao(pedidosTexto);
            } catch (erro) {
                console.error(erro);
                alert("Erro ao ler o arquivo de texto: " + erro.message);
            }
        };
        reader.readAsText(arquivo, 'UTF-8');
    }
}


function exportarParaBlocoNotas() {
    // Seleciona todas as linhas da tabela de resultados
    const linhas = document.querySelectorAll("#tabelaResultado tr");
    
    if (linhas.length === 0 || (linhas.length === 1 && linhas[0].innerText.includes("Nenhum registro"))) {
        alert("Não há dados na tabela para exportar.");
        return;
    }

    let conteudoTexto = "CONCILIAÇÃO DE PEDIDOS ORACLE - DIEBOLD NIXDORF\r\n";
    conteudoTexto += `Data/Hora da Exportação: ${new Date().toLocaleString('pt-BR')}\r\n`;
    conteudoTexto += "--------------------------------------------------\r\n";
    conteudoTexto += "PEDIDO\t\t|\tSTATUS\r\n";
    conteudoTexto += "--------------------------------------------------\r\n";

    // Percorre as linhas preenchendo o conteúdo do arquivo
    linhas.forEach(linha => {
        const colunas = linha.querySelectorAll("td");
        if (colunas.length >= 2) {
            const pedido = colunas[0].innerText.trim();
            const status = colunas[1].innerText.trim();
            conteudoTexto += `${pedido}\t\t|\t${status}\r\n`;
        }
    });

    // Cria o arquivo Blob em formato de texto plano
    const blob = new Blob([conteudoTexto], { type: "text/plain;charset=utf-8" });
    
    // Cria um link temporário para forçar o download
    const linkTemporario = document.createElement("a");
    linkTemporario.href = URL.createObjectURL(blob);
    linkTemporario.download = `Conciliacao_Pedidos_${new Date().toISOString().slice(0,10)}.txt`;
    
    // Simula o clique e remove o link do navegador
    document.body.appendChild(linkTemporario);
    linkTemporario.click();
    document.body.removeChild(linkTemporario);
}

function filtrarResultados() {
    console.log("Filtro acionado");

    const filtro = document.getElementById("filtroStatus").value;
    console.log("Valor:", filtro);

    const linhas = document.querySelectorAll("#tabelaResultado tr");
    console.log("Linhas encontradas:", linhas.length);

    linhas.forEach(linha => {

        const status = linha.cells[1].textContent;

        console.log(status);

        if (filtro === "todos") {
            linha.style.display = "";
        }
        else if (filtro === "expedicao") {
            linha.style.display =
                status.includes("Expedição")
                    ? ""
                    : "none";
        }
        else if (filtro === "naoseparado") {
            linha.style.display =
                status.includes("Não Separado")
                    ? ""
                    : "none";
        }
    });
}
function gerarTabelaComparacao(pedidosDoRelatorio) {
    const tabelaResultado = document.getElementById("tabelaResultado");
    tabelaResultado.innerHTML = "";

    const arraySimplesManuais = pedidosManuais.map(item => item.texto);

    pedidosDoRelatorio.forEach(pedido => {
        const encontrado = arraySimplesManuais.includes(pedido);
        const tr = document.createElement("tr");

        tr.innerHTML = `
            <td><strong>${pedido}</strong></td>
            <td>
                <span style="font-weight:bold;color:${encontrado ? '#22bb33' : '#d9534f'}">
                    ${encontrado ? '✅ Expedição' : '⚠️ Não Separado'}
                </span>
            </td>
        `;

        tabelaResultado.appendChild(tr);
    });

    filtrarResultados();
}
function gerarTabelaComparacao(pedidosDoRelatorio) {

    const tabelaResultado = document.getElementById("tabelaResultado");
    tabelaResultado.innerHTML = "";

    const arraySimplesManuais = pedidosManuais.map(item => item.texto);

    pedidosDoRelatorio.forEach(pedido => {

        const encontrado = arraySimplesManuais.includes(pedido);

        const tr = document.createElement("tr");

        tr.innerHTML = `
            <td><strong>${pedido}</strong></td>
            <td>
                <span style="
                    font-weight:bold;
                    color:${encontrado ? '#22bb33' : '#d9534f'};
                ">
                    ${encontrado ? '✅ Expedição' : '⚠️ Não Separado'}
                </span>
            </td>
        `;

        tabelaResultado.appendChild(tr);
    });

    filtrarResultados();
}

function filtrarResultados() {
    const filtro = document.getElementById("filtroStatus").value;
    const linhas = document.querySelectorAll("#tabelaResultado tr");

    linhas.forEach(linha => {

        const status = linha.cells[1].textContent.trim();

        if (filtro === "todos") {
            linha.style.display = "";
        }
        else if (filtro === "expedicao") {
            linha.style.display =
                status.includes("Expedição")
                    ? ""
                    : "none";
        }
        else if (filtro === "naoseparado") {
            linha.style.display =
                status.includes("Não Separado")
                    ? ""
                    : "none";
        }
    });
}
function imprimirResultados() {

    const tabela = document.getElementById("tabelaResultado");

    if (!tabela || tabela.rows.length === 0) {
        alert("Não existem resultados para imprimir.");
        return;
    }

    const janela = window.open("", "_blank");

    janela.document.write(`
        <!DOCTYPE html>
        <html>
        <head>
            <title>Conciliação de Pedidos Oracle</title>

            <style>
                body{
                    font-family: Arial, sans-serif;
                    margin:20px;
                }

                h1{
                    text-align:center;
                    margin-bottom:20px;
                }

                .data{
                    margin-bottom:20px;
                    font-size:14px;
                }

                table{
                    width:100%;
                    border-collapse:collapse;
                }

                th, td{
                    border:1px solid #000;
                    padding:8px;
                    text-align:left;
                }

                th{
                    background-color:#f0f0f0;
                }
            </style>

        </head>
        <body>

            <h1>CONCILIAÇÃO DE PEDIDOS ORACLE</h1>

            <div class="data">
                Impressão gerada em:
                ${new Date().toLocaleString('pt-BR')}
            </div>

            <table>
                <thead>
                    <tr>
                        <th>Pedido Relatório</th>
                        <th>Status</th>
                    </tr>
                </thead>
                <tbody>
                    ${tabela.innerHTML}
                </tbody>
            </table>

        </body>
        </html>
    `);

    janela.document.close();
    janela.focus();

    setTimeout(() => {
        janela.print();
    }, 500);
}
``
