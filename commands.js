Office.onReady(() => {});

async function inserirAssinatura(event) {
  try {
    const usuario = await obterDadosUsuario();

    let assinaturaHTML = `
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
                  %%DisplayName%%
                </td>
              </tr>
              <tr>
                <td style="padding-top:4px;font-family:Arial,sans-serif;font-size:14px;font-weight:900;color:#706e6e;">
                  %%Title%%
                </td>
              </tr>
              <tr>
                <td style="padding-top:4px;font-family:Arial,sans-serif;font-size:12px;color:#706e6e;">
                  <img src="https://raw.githubusercontent.com/luanmoutinho/assinaturas-fotos/refs/heads/main/icone-telefone.png" width="14" style="vertical-align:middle;" alt="">
                  <span style="color:#706e6e;margin-left:5px;">%%Phone%%</span>
                </td>
              </tr>
              <tr>
                <td style="padding-top:4px;font-family:Arial,sans-serif;font-size:12px;color:#706e6e;">
                  <img src="https://raw.githubusercontent.com/luanmoutinho/assinaturas-fotos/refs/heads/main/icone-email.png" width="14" style="vertical-align:middle;" alt="">
                  <a href="mailto:%%Email%%" style="color:#706e6e;text-decoration:none;margin-left:5px;">%%Email%%</a>
                </td>
              </tr>
              <tr>
                <td style="padding-top:4px;font-family:Arial,sans-serif;font-size:12px;color:#706e6e;">
                  <img src="https://raw.githubusercontent.com/luanmoutinho/assinaturas-fotos/refs/heads/main/icone-local.png" width="14" style="vertical-align:middle;" alt="">
                  <span style="color:#706e6e;margin-left:5px;">%%Street%%</span>
                </td>
              </tr>
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

    // Mapeamento ajustado exatamente para os placeholders do seu HTML
    const mapaPlaceholders = {
      "%%DisplayName%%": usuario.displayName || "",
      "%%Title%%": usuario.jobTitle || "",
      "%%Phone%%": (usuario.businessPhones && usuario.businessPhones[0]) || usuario.mobilePhone || "",
      "%%Email%%": usuario.mail || usuario.userPrincipalName || "",
      "%%Street%%": usuario.streetAddress || ""
    };

    for (const [chave, valor] of Object.entries(mapaPlaceholders)) {
      const regex = new RegExp(chave, "gi");
      assinaturaHTML = assinaturaHTML.replace(regex, valor);
    }

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

/**
 * Busca os dados do usuário no Microsoft Graph usando SSO e salva em cache por 4 horas
 */
async function obterDadosUsuario() {
  const CHAVE_CACHE = "assinatura_gep_perfil_cache";
  const TEMPO_EXPIRACAO_MS = 4 * 60 * 60 * 1000; // 4 horas

  try {
    const cacheSalvo = localStorage.getItem(CHAVE_CACHE);
    if (cacheSalvo) {
      const dadosCache = JSON.parse(cacheSalvo);
      if (Date.now() - dadosCache.timestamp < TEMPO_EXPIRACAO_MS) {
        return dadosCache.perfil;
      }
    }

    // Solicita token silencioso do Office SSO
    const token = await Office.auth.getAccessToken({
      allowSignInPrompt: false,
      allowConsentPrompt: false,
      forMSGraphAccess: true
    });

    // Consulta o Microsoft Graph pedindo explicitamente o streetAddress
    const campos = "displayName,jobTitle,businessPhones,mobilePhone,mail,userPrincipalName,streetAddress";
    const resposta = await fetch(`https://graph.microsoft.com/v1.0/me?$select=${campos}`, {
      headers: {
        Authorization: `Bearer ${token}`
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
    console.warn("Não foi possível obter dados do Graph, usando fallback local:", erro);
    // Fallback com dados básicos do perfil logado no Outlook caso o Graph falhe
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

// Registra a função para o evento OnNewMessageCompose do manifest.xml
Office.actions.associate("inserirAssinatura", inserirAssinatura);