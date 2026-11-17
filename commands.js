const msalConfig = {
  auth: {
    clientId: "31f4d827-b982-4c84-97f6-16bb08d2c1c1",
    authority: "https://login.microsoftonline.com/engenhariaep.com.br",
    redirectUri: "https://luanmoutinho.github.io/assinatura-outlook/commands.html"
  },
  cache: {
    cacheLocation: "localStorage"
  }
};

let msalInstance = null;
let cacheMemoria = null;

function obterMsalInstance() {
  if (msalInstance) return msalInstance;
  try {
    if (typeof msal !== "undefined" && typeof window !== "undefined") {
      msalInstance = new msal.PublicClientApplication(msalConfig);
    }
  } catch (e) {
    msalInstance = null;
  }
  return msalInstance;
}

function aguardar(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function inserirAssinatura(event, isAutomatico = false) {
  try {
    const usuario = await obterDadosUsuario(isAutomatico);

    const nome = (usuario.displayName || "").toLocaleUpperCase("pt-BR");
    const cargo = usuario.jobTitle || "";
    const telefone = (usuario.businessPhones && usuario.businessPhones[0]) || usuario.mobilePhone || "";
    const email = usuario.mail || usuario.userPrincipalName || "";
    const endereco = usuario.streetAddress || "";

    const linhaCargo = cargo
      ? `<tr><td colspan="2" style="padding-top:4px;font-family:Arial,sans-serif;font-size:14px;font-weight:900;color:#706e6e;">${cargo}</td></tr>`
      : "";

    const linhaTelefone = telefone
      ? `<tr>
          <td width="16" valign="middle" style="padding-top:5px;width:16px;text-align:center;">
            <img src="https://raw.githubusercontent.com/luanmoutinho/assinaturas-fotos/refs/heads/main/icone-telefone.png" width="14" style="display:block;border:0;margin:0 auto;" alt="">
          </td>
          <td valign="middle" style="padding-top:5px;padding-left:6px;font-family:Arial,sans-serif;font-size:12px;color:#706e6e;line-height:1.2;">
            ${telefone}
          </td>
        </tr>`
      : "";

    const linhaEmail = email
      ? `<tr>
          <td width="16" valign="middle" style="padding-top:5px;width:16px;text-align:center;">
            <img src="https://raw.githubusercontent.com/luanmoutinho/assinaturas-fotos/refs/heads/main/icone-email.png" width="14" style="display:block;border:0;margin:0 auto;" alt="">
          </td>
          <td valign="middle" style="padding-top:5px;padding-left:6px;font-family:Arial,sans-serif;font-size:12px;color:#706e6e;line-height:1.2;">
            <a href="mailto:${email}" style="color:#706e6e;text-decoration:none;">${email}</a>
          </td>
        </tr>`
      : "";

    const linhaEndereco = endereco
      ? `<tr>
          <td width="16" valign="middle" style="padding-top:5px;width:16px;text-align:center;">
            <img src="https://raw.githubusercontent.com/luanmoutinho/assinaturas-fotos/refs/heads/main/icone-local.png" width="14" style="display:block;border:0;margin:0 auto;" alt="">
          </td>
          <td valign="middle" style="padding-top:5px;padding-left:6px;font-family:Arial,sans-serif;font-size:12px;color:#706e6e;line-height:1.2;">
            ${endereco}
          </td>
        </tr>`
      : "";

    const assinaturaHTML = `
<table cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;table-layout:fixed;background-color:transparent;">
  <tr>
    <td valign="top" width="540" style="background-color:transparent;padding:0;">
      <table cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;table-layout:fixed;background-color:transparent;">
        <tr>
          <td valign="top" style="padding:15px 0 10px 0;">
            <table cellpadding="0" cellspacing="0" border="0" width="160">
              <tr>
                <td colspan="3" style="padding:0 0 5px 0;">
                  <img src="https://raw.githubusercontent.com/luanmoutinho/assinaturas-fotos/refs/heads/main/_LogoGEPv15.png" width="160" style="display:block;border:0;" alt="Logo">
                </td>
              </tr>
              <tr>
                <td style="padding-top:2px;">
                  <img src="https://raw.githubusercontent.com/luanmoutinho/assinaturas-fotos/refs/heads/main/_ISO9001new.png" width="50" style="display:block;border:0;" alt="ISO9001">
                </td>
                <td style="padding-top:2px;padding-left:5px;">
                  <img src="https://raw.githubusercontent.com/luanmoutinho/assinaturas-fotos/refs/heads/main/_ISO14001new.png" width="50" style="display:block;border:0;" alt="ISO14001">
                </td>
                <td style="padding-top:2px;padding-left:5px;">
                  <img src="https://raw.githubusercontent.com/luanmoutinho/assinaturas-fotos/refs/heads/main/_ISO45001new.png" width="50" style="display:block;border:0;" alt="ISO45001">
                </td>
              </tr>
            </table>
          </td>
          <td valign="top" style="padding:15px 0 0 15px;">
            <table cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">
              <tr>
                <td colspan="2" style="font-family:Arial,sans-serif;font-size:18px;font-weight:900;color:#706e6e;text-transform:uppercase;line-height:1.2;">
                  ${nome}
                </td>
              </tr>
              ${linhaCargo}
              ${linhaTelefone}
              ${linhaEmail}
              ${linhaEndereco}
              <tr>
                <td width="16" valign="middle" style="padding-top:5px;width:16px;text-align:center;">
                  <img src="https://raw.githubusercontent.com/luanmoutinho/assinaturas-fotos/refs/heads/main/icone-link.png" width="14" style="display:block;border:0;margin:0 auto;" alt="">
                </td>
                <td valign="middle" style="padding-top:5px;padding-left:6px;font-family:Arial,sans-serif;font-size:12px;color:#706e6e;line-height:1.2;">
                  <a href="https://grupoengenhariaep.com.br/" style="color:#706e6e;text-decoration:none;">grupoengenhariaep.com.br</a>
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
    <td valign="top" width="60" style="width:60px;background-color:transparent;">
      <img src="https://raw.githubusercontent.com/luanmoutinho/assinaturas-fotos/refs/heads/main/newDetail.png" width="60" style="display:block;border:0;width:60px;height:auto;" alt="">
    </td>
  </tr>
  <tr>
    <td colspan="2" style="padding:10px 0 0 0;">
      <p style="font-family:Arial,sans-serif;font-size:8px;color:#999999;text-align:justify;line-height:1.2;margin:0;">
        Este e-mail e quaisquer arquivos transmitidos com ele são confidenciais e destinados exclusivamente ao uso do indivíduo ou entidade a quem são endereçados.
      </p>
    </td>
  </tr>
</table>
`;

    if (isAutomatico) {
      await aguardar(400);
    }

    const corpo = Office.context.mailbox.item.body;

    if (typeof corpo.setSignatureAsync === "function") {
      corpo.setSignatureAsync(
        assinaturaHTML,
        { coercionType: Office.CoercionType.Html },
        function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            corpo.setAsync(
              `<br><br>${assinaturaHTML}`,
              { coercionType: Office.CoercionType.Html },
              function () {
                if (event && typeof event.completed === "function") event.completed();
              }
            );
          } else {
            if (event && typeof event.completed === "function") event.completed();
          }
        }
      );
    } else {
      corpo.setAsync(
        `<br><br>${assinaturaHTML}`,
        { coercionType: Office.CoercionType.Html },
        function () {
          if (event && typeof event.completed === "function") event.completed();
        }
      );
    }
  } catch (erro) {
    if (event && typeof event.completed === "function") {
      event.completed();
    }
  }
}

async function obterDadosUsuario(isAutomatico = false) {
  const CHAVE_CACHE = "assinatura_gep_perfil_v8";
  const CHAVE_BACKUP = "assinatura_gep_backup_v8";
  const TEMPO_EXPIRACAO_MS = 4 * 60 * 60 * 1000;

  try {
    if (cacheMemoria && Date.now() - cacheMemoria.timestamp < TEMPO_EXPIRACAO_MS) {
      return cacheMemoria.perfil;
    }

    if (typeof localStorage !== "undefined") {
      const cacheSalvo = localStorage.getItem(CHAVE_CACHE);
      if (cacheSalvo) {
        const dadosCache = JSON.parse(cacheSalvo);
        if (Date.now() - dadosCache.timestamp < TEMPO_EXPIRACAO_MS && dadosCache.perfil && dadosCache.perfil.jobTitle) {
          cacheMemoria = dadosCache;
          return dadosCache.perfil;
        }
      }
    }

    let accessToken = null;
    const emailLogado = Office.context.mailbox.userProfile.emailAddress;
    const instanciaMsal = obterMsalInstance();

    if (instanciaMsal) {
      await instanciaMsal.initialize();

      const tokenRequest = {
        scopes: ["User.Read"],
        loginHint: emailLogado
      };

      let authResult = null;
      const contas = instanciaMsal.getAllAccounts();
      const contaAlvo = contas.find(c => c.username && c.username.toLowerCase() === emailLogado.toLowerCase()) || contas[0];

      if (contaAlvo) {
        try {
          authResult = await instanciaMsal.acquireTokenSilent({
            ...tokenRequest,
            account: contaAlvo
          });
        } catch (e) {
          authResult = null;
        }
      }

      if (!authResult) {
        try {
          authResult = await instanciaMsal.ssoSilent(tokenRequest);
        } catch (erroSilencioso) {
          if (!isAutomatico) {
            authResult = await instanciaMsal.acquireTokenPopup(tokenRequest);
          }
        }
      }

      if (authResult && authResult.accessToken) {
        accessToken = authResult.accessToken;
      }
    }

    if (!accessToken && typeof Office !== "undefined" && Office.auth && typeof Office.auth.getAccessToken === "function") {
      try {
        accessToken = await Office.auth.getAccessToken({
          allowSignInPrompt: !isAutomatico,
          allowConsentPrompt: !isAutomatico,
          forMSGraphAccess: true
        });
      } catch (e) {
        accessToken = null;
      }
    }

    if (!accessToken) {
      throw new Error("Token indisponível");
    }

    const campos = "displayName,jobTitle,businessPhones,mobilePhone,mail,userPrincipalName,streetAddress";
    const resposta = await fetch(`https://graph.microsoft.com/v1.0/me?$select=${campos}`, {
      headers: {
        Authorization: `Bearer ${accessToken}`
      }
    });

    if (!resposta.ok) {
      throw new Error(`Graph HTTP ${resposta.status}`);
    }

    const perfil = await resposta.json();
    const objetoCache = {
      timestamp: Date.now(),
      perfil: perfil
    };

    cacheMemoria = objetoCache;
    if (typeof localStorage !== "undefined") {
      localStorage.setItem(CHAVE_CACHE, JSON.stringify(objetoCache));
      localStorage.setItem(CHAVE_BACKUP, JSON.stringify(perfil));
    }

    return perfil;
  } catch (erro) {
    if (typeof localStorage !== "undefined") {
      const backupSalvo = localStorage.getItem(CHAVE_BACKUP);
      if (backupSalvo) {
        return JSON.parse(backupSalvo);
      }
    }

    const perfilOutlook = Office.context.mailbox.userProfile;
    return {
      displayName: perfilOutlook.displayName || "",
      jobTitle: "",
      businessPhones: [],
      mobilePhone: "",
      mail: perfilOutlook.emailAddress || "",
      streetAddress: ""
    };
  }
}

async function onNewMessageComposeHandler(event) {
  await inserirAssinatura(event, true);
}

if (typeof globalThis !== "undefined") {
  globalThis.inserirAssinatura = inserirAssinatura;
  globalThis.onNewMessageComposeHandler = onNewMessageComposeHandler;
}

if (typeof Office !== "undefined" && Office.actions && typeof Office.actions.associate === "function") {
  Office.actions.associate("inserirAssinatura", inserirAssinatura);
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
}

if (typeof Office !== "undefined") {
  Office.onReady(function () {
    if (Office.actions && typeof Office.actions.associate === "function") {
      Office.actions.associate("inserirAssinatura", inserirAssinatura);
      Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
    }
  });
}
