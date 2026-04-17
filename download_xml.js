(function () {
  'use strict';

  function getCookie(nome) {
    let cookies = document.cookie.split(';');

    for (let c of cookies) {
      let [key, value] = c.trim().split('=');

      if (key === nome) {
        return decodeURIComponent(value);
      }
    }

    return null;
  }

  function formatarDataBR() {
    try {
      return new Date().toLocaleString("pt-BR", {
        timeZone: "America/Sao_Paulo",
        year: "numeric",
        month: "2-digit",
        day: "2-digit",
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit"
      });
    } catch {
      return "sem_data";
    }
  }

  function interceptar() {
    if (window.MDFe && window.MDFe.exec) {

      const originalExec = window.MDFe.exec;

      window.MDFe.exec = function (btn, id, tipo) {

        let selecionado = document.querySelector('input[name="operacao"]:checked');

        if (selecionado && selecionado.value === "G") {

          let linhaCorreta = document.querySelector('tr.swrg-list-sel');

          let numeroCte = linhaCorreta
            ?.querySelector('td[swni="no"]')
            ?.innerText
            ?.trim();

          if (numeroCte) {

            let usuarioSistema = getCookie("swu") || "desconhecido";
            let dataFormatada = formatarDataBR();

            setTimeout(() => {

              let conteudo = JSON.stringify({
                tipo: "MDFe",
                id: id,
                numeroCte: numeroCte,
                usuarioSistema: usuarioSistema,
                data: dataFormatada
              });

              let blob = new Blob([conteudo], { type: "application/json" });
              let a = document.createElement("a");

              a.href = URL.createObjectURL(blob);
              a.download = "controle_mdfe.json";
              a.click();

            }, 100);

          } else {
            console.warn("Não encontrou número do MDF-e");
          }
        }

        return originalExec.apply(this, arguments);
      };

      console.log("[MDFe] Hook aplicado com sucesso");

    } else {
      setTimeout(interceptar, 1000);
    }
  }

  interceptar();
})();