using System;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Web.WebView2.Core;
using System.IO;

namespace Pre_Launch_Tool
{
    public partial class M365ChatWindow : Window
    {
        private const int AgentWarmupDelayMs = 4000;
        private const int PostPasteDelayMs = 2500;
        private readonly string _targetUrl;
        private readonly string _tsvContent;
        private bool _messageSent = false;
      private bool _didHeaderRecovery = false;

        public M365ChatWindow(string targetUrl, string tsvContent)
        {
            InitializeComponent();
            _targetUrl = targetUrl;
            _tsvContent = tsvContent;

            Loaded += M365ChatWindow_Loaded;
        }

        private async void M365ChatWindow_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
            SetLoadingState(true);
            // Use per-user cache to avoid sharing cookies across Windows accounts.
            string userDataFolder = Path.Combine(
              Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
              "PreLaunchTool",
              "WebView2Cache");

            Directory.CreateDirectory(userDataFolder);

                var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder, null);
                await WebView.EnsureCoreWebView2Async(env);

                WebView.CoreWebView2.NavigationCompleted += CoreWebView2_NavigationCompleted;
                WebView.CoreWebView2.DOMContentLoaded += CoreWebView2_DOMContentLoaded;

                TbStatus.Text = "Carregando p�gina do agente...";
                WebView.CoreWebView2.Navigate(_targetUrl);
            }
            catch (Exception ex)
            {
                TbStatus.Text = $"Erro ao inicializar WebView2: {ex.Message}";
                SetLoadingState(false);
            }
        }

        private async void CoreWebView2_DOMContentLoaded(object? sender, CoreWebView2DOMContentLoadedEventArgs e)
        {
            // Best-effort: try to pre-fill the corporate email on Microsoft login pages.
            // IMPORTANT: In many corporate environments, full SSO login cannot be automated (MFA/policies).
            // This just saves a click/typing when the email field exists.
            try
            {
                string user = Environment.UserName;
                string email = user.Contains('@') ? user : user + "@bosch.com";
                string escapedEmail = EscapeForJavaScript(email);

                string js = $$"""
(() => {
  const loginInput = document.querySelector('input[name="loginfmt"], input[type="email"], input[autocomplete="username"]');
  if (!loginInput) return 'NO_LOGIN_INPUT';
  if (loginInput.value && loginInput.value.length > 0) return 'ALREADY_FILLED';
  loginInput.focus();
  loginInput.value = {{escapedEmail}};
  loginInput.dispatchEvent(new Event('input', { bubbles: true }));
  loginInput.dispatchEvent(new Event('change', { bubbles: true }));

  const nextBtn = document.querySelector('input[type="submit"], button[type="submit"], #idSIButton9');
  if (nextBtn) {
    // do not auto-click; let user confirm (some pages require interaction)
    return 'FILLED_EMAIL';
  }
  return 'FILLED_EMAIL_NO_BUTTON';
})();
""";

                await WebView.CoreWebView2.ExecuteScriptAsync(js);
            }
            catch
            {
                // ignore
            }
        }

        private async void CoreWebView2_NavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            if (!e.IsSuccess)
            {
                TbStatus.Text = "Erro ao carregar a página. Verifique sua conexão.";
              SetLoadingState(false);
                return;
            }

          if (!_didHeaderRecovery)
          {
            try
            {
              string isHeaderTooLong = await WebView.CoreWebView2.ExecuteScriptAsync(
                "(() => { const t = (document.title || '').toLowerCase(); const b = (document.body && document.body.innerText ? document.body.innerText : '').toLowerCase(); return t.includes('header field too long') || b.includes('request header field is too long'); })();");

              if (isHeaderTooLong == "true")
              {
                _didHeaderRecovery = true;
                TbStatus.Text = "Sessão inválida detectada. Limpando cache/cookies e recarregando...";

                await WebView.CoreWebView2.CallDevToolsProtocolMethodAsync("Network.clearBrowserCookies", "{}");
                await WebView.CoreWebView2.CallDevToolsProtocolMethodAsync("Network.clearBrowserCache", "{}");

                WebView.CoreWebView2.Navigate(_targetUrl);
                return;
              }
            }
            catch
            {
              // If detection fails, continue normal flow.
            }
          }

            // Only attempt to paste and send if we're on the target agent URL
            // This prevents sending messages during login redirects
            string currentUrl = WebView.CoreWebView2.Source;
            if (!currentUrl.Contains("m365.cloud.microsoft") || !currentUrl.Contains("/chat"))
            {
                TbStatus.Text = "Página carregada. Se necessário, realize o login SSO...";
                return;
            }

            TbStatus.Text = "Página carregada. Aguardando estabilização do agente...";
            await Task.Delay(AgentWarmupDelayMs);
            TbStatus.Text = "Agente pronto. Iniciando envio automático...";

            await TryPasteAndSendAsync();
        }

        private async Task TryPasteAndSendAsync()
        {
            if (_messageSent) return;

            try { Clipboard.SetText(_tsvContent); } catch { }

            string escapedContent = EscapeForJavaScript(_tsvContent);

            int maxAttempts = 300; // up to ~10 minutes
            for (int attempt = 0; attempt < maxAttempts; attempt++)
            {
                if (_messageSent) return;

                try
                {
                    // 1) Wait until a prompt editor is available.
                    // The page UI changes frequently, so we accept multiple known composer shapes.
                    string waitAgentJs = """
(() => {
  const editor = document.querySelector('#m365-chat-editor-target-element')
    || document.querySelector('span[contenteditable="true"][data-lexical-editor="true"]')
    || document.querySelector('[contenteditable="true"][role="textbox"]')
    || document.querySelector('textarea')
    || document.querySelector('[contenteditable="true"]');
  return editor ? 'AGENT_READY' : 'AGENT_NOT_READY';
})();
""";

                    string agentReady = await WebView.CoreWebView2.ExecuteScriptAsync(waitAgentJs);
                    agentReady = agentReady.Trim('"');
                    if (agentReady != "AGENT_READY")
                    {
                        TbStatus.Text = $"Aguardando agente 'Pr�-Launch tool'... ({attempt + 1}/{maxAttempts})";
                        await Task.Delay(2000);
                        continue;
                    }

                    // 2) Paste into editor
                    string pasteJs = $$"""
(() => {
  const editor = document.querySelector('#m365-chat-editor-target-element')
    || document.querySelector('span[contenteditable="true"][data-lexical-editor="true"]')
    || document.querySelector('[contenteditable="true"][role="textbox"]')
    || document.querySelector('textarea')
    || document.querySelector('[contenteditable="true"]');
  if (!editor) return 'NOT_READY_EDITOR';

  editor.focus();

  try {
    const sel = window.getSelection();
    const range = document.createRange();
    range.selectNodeContents(editor);
    sel.removeAllRanges();
    sel.addRange(range);
    document.execCommand('delete', false);
  } catch (e) {
  }

  let ok = false;
  try {
    ok = document.execCommand('insertText', false, {{escapedContent}});
  } catch (e) {
  }

  if (!ok && 'value' in editor) {
    try {
      editor.value = {{escapedContent}};
      editor.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: {{escapedContent}} }));
      ok = true;
    } catch (e) {
    }
  }

  try {
    editor.dispatchEvent(new Event('input', { bubbles: true }));
  } catch (e) {
  }

  return ok ? 'TEXT_SET' : 'TEXT_SET_FALLBACK';
})();
""";

                    string pasteResult = await WebView.CoreWebView2.ExecuteScriptAsync(pasteJs);
                    pasteResult = pasteResult.Trim('"');

                    if (pasteResult != "TEXT_SET" && pasteResult != "TEXT_SET_FALLBACK")
                    {
                        TbStatus.Text = $"Aguardando campo de entrada... ({attempt + 1}/{maxAttempts})";
                        await Task.Delay(2000);
                        continue;
                    }

                    TbStatus.Text = "Conteudo inserido. Aguardando antes de enviar...";
                    await Task.Delay(PostPasteDelayMs);
                    TbStatus.Text = "Conteudo inserido. Aguardando botao de envio...";

                    // 3) Wait for the send arrow SVG to exist, then click it.
                    // Use robust selector: fluent icon svg with viewBox 0 0 24 24 and that specific path.
                    string clickSendJs = """
(() => {
  const svgs = Array.from(document.querySelectorAll('svg'));
  const sendSvg = svgs.find(svg => {
    const vb = (svg.getAttribute('viewBox') || '').trim();
    if (vb !== '0 0 24 24') return false;
    const path = svg.querySelector('path');
    if (!path) return false;
    const d = (path.getAttribute('d') || '').trim();
    return d === 'M13.27 4.2a.75.75 0 0 0-1.04 1.1l6.25 5.95H3.75a.75.75 0 0 0 0 1.5h14.73l-6.25 5.95a.75.75 0 0 0 1.04 1.1l7.42-7.08a1 1 0 0 0 0-1.44L13.27 4.2Z';
  });

  if (!sendSvg) return 'NO_SEND_SVG';

  // Click the closest button if possible; otherwise click the svg itself.
  const btn = sendSvg.closest('button');
  if (btn) {
    try {
      btn.disabled = false;
      btn.removeAttribute('disabled');
    } catch (e) {}
    btn.click();
    return 'SENT_CLICK_BUTTON';
  }

  sendSvg.click();
  return 'SENT_CLICK_SVG';
})();
""";

                    // Retry a few seconds for send icon to appear
                    bool sent = false;
                    for (int s = 0; s < 30; s++)
                    {
                        string sendResult = await WebView.CoreWebView2.ExecuteScriptAsync(clickSendJs);
                        sendResult = sendResult.Trim('"');
                        if (sendResult == "SENT_CLICK_BUTTON" || sendResult == "SENT_CLICK_SVG")
                        {
                            _messageSent = true;
                            TbStatus.Text = $"Mensagem enviada ({sendResult}).";
                          SetLoadingState(false);
                            sent = true;
                            break;
                        }

                        if (sendResult == "NO_SEND_SVG")
                        {
                            string shortcutJs = """
(() => {
  const editor = document.querySelector('#m365-chat-editor-target-element')
    || document.querySelector('span[contenteditable="true"][data-lexical-editor="true"]')
    || document.querySelector('[contenteditable="true"][role="textbox"]')
    || document.querySelector('textarea')
    || document.querySelector('[contenteditable="true"]');
  if (!editor) return 'NO_EDITOR';
  editor.focus();
  const eventInit = { bubbles: true, cancelable: true, key: 'Enter', code: 'Enter', which: 13, keyCode: 13, ctrlKey: true };
  editor.dispatchEvent(new KeyboardEvent('keydown', eventInit));
  editor.dispatchEvent(new KeyboardEvent('keyup', eventInit));
  return 'SENT_CTRL_ENTER';
})();
""";
                            string shortcutResult = await WebView.CoreWebView2.ExecuteScriptAsync(shortcutJs);
                            shortcutResult = shortcutResult.Trim('"');
                            if (shortcutResult == "SENT_CTRL_ENTER")
                            {
                                _messageSent = true;
                                TbStatus.Text = "Mensagem enviada (SENT_CTRL_ENTER).";
                              SetLoadingState(false);
                                sent = true;
                                break;
                            }
                        }

                        await Task.Delay(1000);
                    }

                    if (sent) return;

                    TbStatus.Text = "Bot�o de envio n�o apareceu. Use Ctrl+Enter/colar manualmente.";
                    SetLoadingState(false);
                    return;
                }
                catch
                {
                    // ignore and retry
                }

                TbStatus.Text = $"Aguardando chat ficar dispon�vel... ({attempt + 1}/{maxAttempts})";
                await Task.Delay(2000);
            }

            TbStatus.Text = "Tempo esgotado. Use Ctrl+V para colar manualmente.";
            try { Clipboard.SetText(_tsvContent); } catch { }
            SetLoadingState(false);
        }

          private void SetLoadingState(bool isLoading)
          {
            try
            {
              if (PbLoading != null)
                PbLoading.Visibility = isLoading ? Visibility.Visible : Visibility.Collapsed;
            }
            catch
            {
              // ignore visual update failures
            }
          }

        private static string EscapeForJavaScript(string text)
        {
            if (string.IsNullOrEmpty(text)) return "''";

            text = text
                .Replace("\\", "\\\\")
                .Replace("'", "\\'")
                .Replace("\"", "\\\"")
                .Replace("\r\n", "\\n")
                .Replace("\r", "\\n")
                .Replace("\n", "\\n")
                .Replace("\t", "\\t");

            return "\"" + text + "\"";
        }
    }
}
