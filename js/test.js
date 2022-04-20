(async function () {
  // lidar com o JSON da planilha

  function handleJSON(arr) {
    let output = [];

    arr.forEach((elem, index) => {
      if (
        typeof output[output.length - 1] == "undefined" ||
        !Object.values(output[output.length - 1]).includes(elem.Email)
      ) {
        let duplicatas = arr
          .filter(({ Email }) => Email == elem.Email)
          .map((elem) => elem.Vale);

        let valesObj = {};

        for (let i = 0; i < duplicatas.length; i++) {
          valesObj[`Vale${i + 1}`] = duplicatas[i];
        }

        output.push({
          Colaborador: elem.Colaborador,
          Email: elem.Email,
          ...valesObj,
        });
      }
    });

    return output;
  }

  function renderStatus(mensagem) {
    return (document.querySelector(".status").innerText = mensagem);
  }

  // lidar com o upload da planilha

  const arquivoInput = document.getElementById("arquivo-xlsx");

  arquivoInput.onchange = () => {
    const arquivoSelecionado = arquivoInput.files[0];
    const leitor = new FileReader();

    leitor.onload = function (e) {
      let data = new Uint8Array(e.target.result);
      let workbook = XLSX.read(data, { type: "array" });
      let planilha = workbook.Sheets[workbook.SheetNames[0]];

      renderStatus(".status", "Processando planilha...");

      json = XLSX.utils.sheet_to_json(planilha);

      renderStatus("Gerando planilha formatada...");

      gerarPlanilha(json, "ValesColaboradores.xlsx");
    };

    leitor.readAsArrayBuffer(arquivoSelecionado);
  };

  // gerar download da planilha com as colunas de vales

  function gerarPlanilha(json, nome) {
    let ws = XLSX.utils.json_to_sheet(handleJSON(json));
    let wb = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
    XLSX.writeFile(wb, nome);
  }
})();
