Office.onReady(() => {});

const msalConfig = {
  auth: {
    clientId: "c4232071-8cb2-4038-a318-8b82a1118dd5",
    authority: "https://login.microsoftonline.com/engenhariaep.com.br",
    redirectUri: "https://luanmoutinho.github.io/assinatura-outlook/commands.html"
  },
  cache: {
    cacheLocation: "localStorage"
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

async function inserirAssinatura(event) {
  try {
    const usuario = await obterDadosUsuario();

    const nome = usuario.displayName || "";
    const cargo = usuario.jobTitle || "";
    const telefone = (usuario.businessPhones && usuario.businessPhones[0]) || usuario.mobilePhone || "";
    const email = usuario.mail || usuario.userPrincipalName || "";
    const endereco = usuario.streetAddress || "";

    const linhaCargo = cargo
      ? `<tr><td style="padding-top:4px;font-family:Arial,sans-serif;font-size:14px;font-weight:900;color:#706e6e;">${cargo}</td></tr>`
      : "";

    const linhaTelefone = telefone
      ? `<tr><td style="padding-top:4px;font-family:Arial,sans-serif;font-size:12px;color:#706e6e;"><img src="https://raw.githubusercontent.com/luanmoutinho/assinaturas-fotos/refs/heads/main/icone-telefone.png" width="14" style="vertical-align:middle;" alt=""><span style="color:#706e6e;margin-left:5px;">${telefone}</span></td></tr>`
      : "";

    const linhaEmail = email
      ? `<tr><td style="padding-top:4px;font-family:Arial,sans-serif;font-size:12px;color:#706e6e;"><img src="https://raw.githubusercontent.com/luanmoutinho/assinaturas-fotos/refs/heads/main/icone-email.png" width="14" style="vertical-align:middle;" alt=""><a href="mailto:${email}" style="color:#706e6e;text-decoration:none;margin-left:5px;">${email}</a></td></tr>`
      : "";

    const linhaEndereco = endereco
      ? `<tr><td style="padding-top:4px;font-family:Arial,sans-serif;font-size:12px;color:#706e6e;"><img src="https://raw.githubusercontent.com/luanmoutinho/assinaturas-fotos/refs/heads/main/icone-local.png" width="14" style="vertical-align:middle;" alt=""><span style="color:#706e6e;margin-left:5px;">${endereco}</span></td></tr>`
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
            <table>
              <tr>
                <td style="font-family:Arial,sans-serif;font-size:18px;font-weight:900;color:#706e6e;text-transform:uppercase;line-height:1.2;">
                  ${nome}
                </td>
              </tr>
              ${linhaCargo}
              ${linhaTelefone}
              ${linhaEmail}
              ${linhaEndereco}
              <tr>
                <td style="padding-top:4px;font-family:Arial,sans-serif;font-size:12px;color:#706e6e;">
                  <img src="https://raw.githubusercontent.com/luanmoutinho/assinaturas-fotos/refs/heads/main/icone-link.png" width="14" style="vertical-align:middle;" alt="">
                  <a href="https://grupoengenhariaep.com.br/" style="color:#706e6e;text-decoration:none;margin-left:5px;">grupoengenhariaep.com.br</a>
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

    Office.context.mailbox.item.body.setSignatureAsync(
      assinaturaHTML,
      { coercionType: Office.CoercionType.Html },
      function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Erro ao inserir assinatura: " + asyncResult.error.message);
        }
        event.completed();
      }
    );
  } catch (erro) {
    console.error("Falha ao processar assinatura:", erro);
    event.completed();
  }
}

async function obterDadosUsuario() {
  const CHAVE_CACHE = "assinatura_gep_perfil_v3";
  const TEMPO_EXPIRACAO_MS = 4 * 60 * 60 * 1000;

  try {
    const cacheSalvo = localStorage.getItem(CHAVE_CACHE);
    if (cacheSalvo) {
      const dadosCache = JSON.parse(cacheSalvo);
      if (Date.now() - dadosCache.timestamp < TEMPO_EXPIRACAO_MS) {
        return dadosCache.perfil;
      }
    }

    await msalInstance.initialize();

    const emailLogado = Office.context.mailbox.userProfile.emailAddress;
    const tokenRequest = {
      scopes: ["User.Read"],
      loginHint: emailLogado
    };

    let authResult;
    const contas = msalInstance.getAllAccounts();
    if (contas.length > 0) {
      authResult = await msalInstance.acquireTokenSilent({
        ...tokenRequest,
        account: contas[0]
      });
    } else {
      try {
        authResult = await msalInstance.ssoSilent(tokenRequest);
      } catch (erroSilencioso) {
        authResult = await msalInstance.acquireTokenPopup(tokenRequest);
      }
    }

    const campos = "displayName,jobTitle,businessPhones,mobilePhone,mail,userPrincipalName,streetAddress";
    const resposta = await fetch(`https://graph.microsoft.com/v1.0/me?$select=${campos}`, {
      headers: {
        Authorization: `Bearer ${authResult.accessToken}`
      }
    });

    if (!resposta.ok) {
      throw new Error(`Erro HTTP Graph: ${resposta.status}`);
    }

    const perfil = await resposta.json();

    localStorage.setItem(
      CHAVE_CACHE,
      JSON.stringify({
        timestamp: Date.now(),
        perfil: perfil
      })
    );

    return perfil;
  } catch (erro) {
    console.warn("Fallback local ativado:", erro);
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

Office.actions.associate("inserirAssinatura", inserirAssinatura);
