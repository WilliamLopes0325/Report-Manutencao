console.log("SCRIPT CARREGADO");

let dadosExcel = [];

window.onload = function() {

    console.log("HTML CARREGADO");

    document
      .getElementById("excel")
      .addEventListener("change", lerExcel);

   const hoje = new Date();

const ano = hoje.getFullYear();

const mes =
    String(hoje.getMonth() + 1)
        .padStart(2, "0");

const dia =
    String(hoje.getDate())
        .padStart(2, "0");

document.getElementById("data").value =
    `${ano}-${mes}-${dia}`;

};
function gerarPreview() {

    console.log("BOTÃO CLICADO");
    console.table(dadosExcel);

    gerarBlocosPorLinha();

    const data =
        document.getElementById("data").value;

    const turno =
        document.getElementById("turno").value;

    const observacoes =
        document.getElementById("observacoes").value;

    const pendencias =
        document.getElementById("pendencias").value;

    let dataFormatada = "";

    if (data) {
        dataFormatada =
            new Date(data)
            .toLocaleDateString("pt-BR");
    }

    document.getElementById("tituloEmail")
        .innerText =
        `REPORT MANUTENÇÃO — ${turno}`;

    document.getElementById("dataPreview")
        .innerText =
        `Data: ${dataFormatada}`;

    document.getElementById("turnoPreview")
        .innerText =
        `Turno: ${turno}`;

    document.getElementById("previewObservacoes")
        .innerText =
        observacoes;

    document.getElementById("previewPendencias")
        .innerText =
        pendencias;
}
function lerExcel(event) {

  console.log("ARQUIVO SELECIONADO");

  const arquivo = event.target.files[0];

  if (!arquivo) return;

  const reader = new FileReader();

  reader.onload = function(e) {

    const data = new Uint8Array(e.target.result);

    const workbook = XLSX.read(data, {
      type: "array"
    });

    const primeiraPlanilha =
      workbook.SheetNames[0];

    const worksheet =
      workbook.Sheets[primeiraPlanilha];

    dadosExcel =
      XLSX.utils.sheet_to_json(worksheet);

      console.log("Excel LIDO");
      console.table(dadosExcel);

    alert(
      `${dadosExcel.length} registros carregados`
    );
  };

  reader.readAsArrayBuffer(arquivo);
}

function atualizarIndicadores() {

  const total =
    dadosExcel.length;

  document.getElementById("total")
    .innerText =
    total;
}
function gerarBlocosPorLinha() {

    const container =
        document.getElementById("areas");

    if (!container) return;

    container.innerHTML = "";

    const grupos = {};

    dadosExcel.forEach(item => {

        const linha =
            item["Linha"] || "SEM LINHA";

        if (!grupos[linha]) {
            grupos[linha] = [];
        }

        grupos[linha].push(item);
    });

    for (const linha in grupos) {

        let html = `
            <div class="bloco">
                <div class="bloco-titulo">
                    ${linha}
                </div>
        `;

        grupos[linha].forEach(item => {

            html += `
                <p><strong>Processo:</strong></p>
                <ul>
                    <li>${item["Processo"] || ""}</li>
                </ul>

                <p><strong>Sintoma:</strong></p>
                <ul>
                    <li>${item["Sintoma"] || ""}</li>
                </ul>

                <p><strong>Comentário Técnico:</strong></p>
                <ul>
                    <li>${item["Comentario Tecnico"] || ""}</li>
                </ul>

                <hr>
            `;
        });

        html += `</div>`;

        container.innerHTML += html;
    }
}
function limparCampos() {

    // Formulário
    const hoje = new Date();

    document.getElementById("data").value =
    hoje.toISOString().split("T")[0];
    document.getElementById("turno").selectedIndex = 0;
    document.getElementById("assunto").value = "Report Manutenção";

    document.getElementById("observacoes").value = "";
    document.getElementById("pendencias").value = "";

    document.getElementById("excel").value = "";

    // Prévia
    document.getElementById("tituloEmail").innerText =
        "REPORT MANUTENÇÃO — T1";

    document.getElementById("dataPreview").innerText =
        "Data:";

    document.getElementById("turnoPreview").innerText =
        "Turno:";

    document.getElementById("previewObservacoes").innerText = "";
    document.getElementById("previewPendencias").innerText = "";

    // Tabela
    document.getElementById("areas").innerHTML = "";

    // Limpa memória
    dadosExcel = [];

    console.log("Report limpo");
}
