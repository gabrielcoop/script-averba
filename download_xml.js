(function () {
  'use strict';

  function getCookie(nome) {
    for (let c of document.cookie.split(';')) {
      let [key, value] = c.trim().split('=');
      if (key === nome) return decodeURIComponent(value);
    }
    return null;
  }

  function formatarDataBR() {
    try {
      const agora = new Date();

      return agora.toLocaleString('pt-BR', {
        timeZone: 'America/Sao_Paulo',
        hour12: false
      });
    } catch {
      return new Date().toLocaleString('pt-BR');
    }
  }

  function gerarControle(id, numero, tipo) {
    const usuarioSistema = getCookie("swu") || "desconhecido";
    const dataBR = formatarDataBR();

    const conteudo = JSON.stringify({
      tipo: tipo,
      id: id,
      numero: numero,
      usuarioSistema: usuarioSistema,
      data: dataBR
    }, null, 2);

    const blob = new Blob([conteudo], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "controle_mdfe.json";

    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }

  function interceptar() {
    if (window.MDFe && window.MDFe.exec) {

      const originalExec = window.MDFe.exec;

      window.MDFe.exec = function (btn, id, tipo) {

        const selecionado = document.querySelector('input[name="operacao"]:checked');

        if (selecionado && selecionado.value === "G") {

          const linhaCorreta = document.querySelector('tr.swrg-list-sel');
          if (linhaCorreta?.querySelector('td[swni="no"]')) {
            const numero = linhaCorreta
              ?.querySelector('td[swni="no"]')
              ?.innerText
              ?.trim();
            if (numero) {
              gerarControle(id, numero, "CT-e");
            }
          }

          else if (linhaCorreta?.querySelector('td[swni="numero"]')) {
            const numero = linhaCorreta
              ?.querySelector('td[swni="numero"]')
              ?.innerText
              ?.trim();
            if (numero) {
              gerarControle(id, numero, "MDF-e");
            }
          }
        }

        return originalExec.apply(this, arguments);
      };


    } else {
      setTimeout(interceptar, 1000);
    }
  }

  interceptar();
})();