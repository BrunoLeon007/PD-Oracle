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
/** @type {import("./xlsx")} */
const XLSX = globalThis.XLSX;
