import customtkinter as ctk
import threading
import time
import os
import datetime
import re
import urllib.request
import unicodedata
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# ==============================
# CONFIGURAÇÕES INICIAIS
# ==============================

RESPONSAVEIS = {
    "Carla Zambi Meirelles": "020.212.397-99",
    "Gabriela Lopes Salgado Novaes": "107.853.177-32",
    "Larissa Janiques Pinto": "138.905.957-07"
}

LOG_FILE = "logs_ueci.txt"

# Detecta o usuário do computador
USUARIO_PC = os.getenv("USERNAME", "Usuário")

# Ambiente SISPREV: por padrão PRODUÇÃO. Pode ser alterado via variável de ambiente SISPREV_BASE_URL
# Ex.: SISPREV_BASE_URL=https://hom.previdencia.es.gov.br/sisprevweb
BASE_URL = os.getenv("SISPREV_BASE_URL", "https://previdencia.es.gov.br/sisprevweb")

ASSINANTES_POR_USUARIO = {
    "albert.junior": "ALBERT IGLÉSIA CORREA DOS SANTOS JUNIOR",
    "larissa.janiques": "LARISSA JANIQUES PINTO",
    "carla.meirelles": "CARLA ZAMBI MEIRELLES",
    "gabriela.novaes": "GABRIELA LOPES SALGADO NOVAES",
}

def obter_assinante_nome():
    try:
        return ASSINANTES_POR_USUARIO.get((USUARIO_PC or "").lower())
    except Exception:
        return None

def _normalize_text(s: str) -> str:
    """Normaliza texto: remove acentos, sobe para maiúsculas e elimina pontuação/espaços.
    Útil para detectar variações como 'C.P.A.D', 'C P A D', etc.
    """
    try:
        s = s or ""
        s = s.strip()
        s = unicodedata.normalize('NFD', s)
        s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')  # remove diacríticos
        s = s.upper()
        s = re.sub(r"[^A-Z0-9]+", "", s)
        return s
    except Exception:
        return (s or "").upper()

# Timings (ajustáveis) para a etapa de tramitação
MODAL_OPEN_DELAY = 0.6            # antes: 1.2
AFTER_SELECT_DELAY = 0.3          # antes: 0.6
SYNC_TIMEOUT_PRIMARY = 1.2        # antes: 8.0 (espera curta após colar)
SYNC_TIMEOUT_PRECLICK = 0.5       # antes: 2.0 (espera curtíssima antes do clique)

# ==============================
# FUNÇÕES PRINCIPAIS
# ==============================

def registrar_log(mensagem):
    """Registra uma mensagem no arquivo de log com data e hora."""
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] {mensagem}\n")

def atualizar_status(msg):
    """Atualiza o texto do status dinamicamente."""
    status_label.configure(text=msg)
    root.update_idletasks()

def mostrar_aviso_e_encerrar(msg: str, segundos: int = 5):
    """Exibe um aviso em uma janelinha por N segundos e encerra a aplicação."""
    def _show():
        try:
            top = ctk.CTkToplevel(root)
            top.title("Aviso")
            top.attributes("-topmost", True)
            # Centraliza janela
            try:
                root.update_idletasks()
                rw = 420; rh = 140
                rx = root.winfo_x() + (root.winfo_width() - rw)//2
                ry = root.winfo_y() + (root.winfo_height() - rh)//2
                top.geometry(f"{rw}x{rh}+{max(rx,0)}+{max(ry,0)}")
            except Exception:
                top.geometry("420x140")

            frame = ctk.CTkFrame(top, corner_radius=12)
            frame.pack(expand=True, fill="both", padx=12, pady=12)

            label = ctk.CTkLabel(
                frame,
                text=msg,
                font=ctk.CTkFont(size=14, weight="bold"),
                justify="center",
                wraplength=380,
            )
            label.pack(expand=True, fill="both", padx=12, pady=(18, 6))

            sub = ctk.CTkLabel(
                frame,
                text=f"Fechando em {segundos} segundo(s)…",
                font=ctk.CTkFont(size=12)
            )
            sub.pack(pady=(0, 12))

            # Agenda encerramento
            root.after(max(1000, segundos*1000), root.destroy)
        except Exception:
            # Fallback: encerra sem UI se algo der errado
            root.after(0, root.destroy)

    # Garante execução no loop principal do Tk
    root.after(0, _show)

def porta_debug_aberta():
    """Verifica rapidamente se o Chrome está disponível em localhost:9222."""
    try:
        with urllib.request.urlopen("http://localhost:9222/json/version", timeout=0.8) as resp:
            return resp.status == 200
    except Exception:
        return False

def abrir_concessao(driver, wait):
    """Abre a tela Benefício > Concessão preferencialmente por URL direta, com fallback no menu.
    Retorna True em caso de sucesso, False caso contrário.
    """
    url_concessao = f"{BASE_URL}/ProcessoBeneficio/ConProcessoBeneficio.aspx"
    # Tenta por URL direta (mais rápido e estável) com pequenas tentativas
    ultimo_erro = None
    for tentativa in range(2):
        try:
            driver.get(url_concessao)
            WebDriverWait(driver, 10, poll_frequency=0.3).until(
                EC.presence_of_element_located((By.ID, "ctl00_ContentCampos_ddlSetor"))
            )
            registrar_log("Concessão aberta via URL direta")
            return True
        except Exception as e:
            ultimo_erro = e
            registrar_log(f"[Aviso] Falha ao abrir via URL direta (tentativa {tentativa+1}/2): {e}")
            try:
                time.sleep(1 + tentativa)
            except Exception:
                pass

    # Falhou a abertura direta; tenta estabelecer contexto partindo da home/base
    try:
        driver.get(BASE_URL)
        time.sleep(0.6)
    except Exception as e:
        registrar_log(f"[Aviso] Falha ao carregar BASE_URL: {e}")
    # Verifica se caiu em tela de login/aviso e conduz o fluxo
    try:
        atual = driver.current_url
    except Exception:
        atual = ""

    if ("/Login/" in atual) or ("AvisoLogin" in atual):
        try:
            # Clica em "Clique aqui para logar novamente." se existir
            link = driver.find_elements(By.XPATH, "//a[contains(.,'Clique aqui') and contains(.,'logar')]")
            if link:
                driver.execute_script("arguments[0].click();", link[0])
        except Exception:
            pass

        atualizar_status("🔐 Sessão expirada — faça login. Aguardando até 60s…")
        try:
            WebDriverWait(driver, 60, poll_frequency=0.5).until(
                lambda d: ("/Login/" not in d.current_url) and ("AvisoLogin" not in d.current_url)
            )
        except TimeoutException:
            registrar_log("[Erro] Login não detectado no tempo limite (60s).")
            return False

    # Após base/login, tenta novamente via URL direta somente mais uma vez
    try:
        driver.get(url_concessao)
        WebDriverWait(driver, 10, poll_frequency=0.3).until(
            EC.presence_of_element_located((By.ID, "ctl00_ContentCampos_ddlSetor"))
        )
        registrar_log("Concessão aberta via URL direta após base/login")
        return True
    except Exception as e2:
        registrar_log(f"[Aviso] Ainda não abriu via URL direta após base/login: {e2}")

    # Fallback: tenta via menu
    try:
        # Alguns layouts exigem normalizar espaços no texto
        try:
            beneficio_menu = WebDriverWait(driver, 5, poll_frequency=0.2).until(
                EC.element_to_be_clickable((By.XPATH, "//a[normalize-space(.)='Benefício']"))
            )
        except Exception:
            beneficio_menu = driver.find_element(By.XPATH, "//a[contains(.,'Benef')]" )
        beneficio_menu.click()

        concessao_link = WebDriverWait(driver, 5, poll_frequency=0.2).until(
            EC.element_to_be_clickable((By.XPATH, "//a[normalize-space(.)='Concessão']"))
        )
        concessao_link.click()

        WebDriverWait(driver, 10, poll_frequency=0.3).until(
            EC.presence_of_element_located((By.ID, "ctl00_ContentCampos_ddlSetor"))
        )
        registrar_log("Concessão aberta via menu")
        return True
    except Exception as e:
        registrar_log(f"[Erro] Não foi possível abrir Concessão por URL nem por menu: {e}")
        if ultimo_erro:
            registrar_log(f"[Det] Último erro de navegação direta: {ultimo_erro}")
        return False

def preencher_informacoes_controle_interno(driver, wait, nome_responsavel, cpf_responsavel):
    """Preenche o parecer e os dados do responsável do Controle Interno dentro do processo."""
    try:
        # Abre aba "Mais Informações do Processo" - tenta pelo ID primeiro (mais rápido)
        try:
            aba_info = wait.until(EC.element_to_be_clickable(
                (By.ID, "__tab_ctl00_ContentCampos_TabContainer1_tabTCE")
            ))
            driver.execute_script("arguments[0].click();", aba_info)
            print("[OK] Aba clicada via ID")
        except:
            aba_info = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//span[contains(text(),'Mais Informações do Processo')]")
            ))
            driver.execute_script("arguments[0].click();", aba_info)
            print("[OK] Aba clicada via XPATH")
        time.sleep(1)

        # Seleciona “Não foi objeto do exame”
        select_parecer = Select(wait.until(
            EC.element_to_be_clickable((By.ID, "ctl00_ContentCampos_TabContainer1_tabTCE_parecerControleInternoTCE"))
        ))
        select_parecer.select_by_visible_text("Não foi objeto do exame")

        # Ativa e preenche campos de CPF e Nome
        driver.execute_script("document.getElementById('ctl00_ContentCampos_TabContainer1_tabTCE_txtCPFRespControleInternoTCE').removeAttribute('disabled');")
        driver.execute_script("document.getElementById('ctl00_ContentCampos_TabContainer1_tabTCE_txtNomeRespControleInternoTCE').removeAttribute('disabled');")
        time.sleep(0.3)

        campo_cpf = driver.find_element(By.ID, "ctl00_ContentCampos_TabContainer1_tabTCE_txtCPFRespControleInternoTCE")
        campo_nome = driver.find_element(By.ID, "ctl00_ContentCampos_TabContainer1_tabTCE_txtNomeRespControleInternoTCE")

        campo_cpf.clear()
        campo_cpf.send_keys(cpf_responsavel)
        campo_nome.clear()
        campo_nome.send_keys(nome_responsavel)

        # Salvar
        btn_salvar = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentToolBar_btnSalvar")))
        driver.execute_script("arguments[0].click();", btn_salvar)
        time.sleep(1)

        print(f"[OK] Informações preenchidas para {nome_responsavel}")
        registrar_log(f"[OK] Informações preenchidas para {nome_responsavel}")

    except Exception as e:
        print(f"[Erro] Falha ao preencher informações de controle interno: {e}")
        registrar_log(f"[Erro] Falha ao preencher informações de controle interno: {e}")
        raise


def preencher_editor_observacao(driver, wait, texto_html: str):
    """Preenche o campo de observação do painel de tramitação.
    Ordem de tentativa: contenteditable -> iframe -> textarea/hidden.
    Sempre tenta sincronizar o hidden (se existir) após preencher o editor visual.
    Retorna True se conseguiu preencher, senão False.
    """
    html = texto_html.replace("\n", "<br>")

    # 1) Contenteditable direto (mais comum)
    try:
        sucesso = driver.execute_script(
            """
            var el = document.querySelector('body[contenteditable="true"], [contenteditable="true"]');
            if(!el) return false;
            el.scrollIntoView({block:'center'});
            el.focus();
            el.innerHTML = arguments[0];
            ['input','keyup','change','blur'].forEach(function(evt){
                var e = new Event(evt, {bubbles:true}); el.dispatchEvent(e);
            });
            var ta = document.getElementById('ctl00_ContentToolBar_txtObservacao');
            if(ta){
                ta.value = el.innerText;
                ['input','keyup','change','blur'].forEach(function(evt){
                    var e = new Event(evt, {bubbles:true}); ta.dispatchEvent(e);
                });
            }
            return !!(el.innerText && el.innerText.length);
            """,
            html
        )
        if sucesso:
            return True
    except Exception:
        pass

    # 2) Editor em iframe
    try:
        candidatos = driver.find_elements(By.XPATH, "//iframe[contains(@id,'txtObservacao') or contains(@name,'txtObservacao') or contains(@id,'ContentToolBar') or contains(@id,'Editor')]")
        for frame in candidatos:
            try:
                driver.switch_to.frame(frame)
                body = WebDriverWait(driver, 5, poll_frequency=0.2).until(
                    EC.presence_of_element_located((By.TAG_NAME, "body"))
                )
                body.click(); time.sleep(0.2)
                driver.execute_script("arguments[0].innerHTML = arguments[1];", body, html)
                ok = driver.execute_script("return arguments[0].innerText && arguments[0].innerText.length;", body)
                driver.switch_to.default_content()
                if ok and int(ok) > 0:
                    # sincroniza hidden
                    try:
                        driver.execute_script(
                            """
                            var ta=document.getElementById('ctl00_ContentToolBar_txtObservacao');
                            if(ta){
                                ta.value=arguments[0];
                                ['input','keyup','change','blur'].forEach(function(evt){
                                    var e = new Event(evt, {bubbles:true}); ta.dispatchEvent(e);
                                });
                            }
                            """,
                            texto_html
                        )
                    except Exception:
                        pass
                    return True
            except Exception:
                driver.switch_to.default_content()
                continue
    except Exception:
        pass

    # 3) Tentativa direta no elemento conhecido (textarea/hidden)
    try:
        elem = WebDriverWait(driver, 5, poll_frequency=0.3).until(
            EC.presence_of_element_located((By.ID, "ctl00_ContentToolBar_txtObservacao"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", elem)
        time.sleep(0.2)
        try:
            elem.clear()
        except Exception:
            pass
        try:
            elem.send_keys(texto_html)
        except Exception:
            driver.execute_script("arguments[0].value = arguments[1];", elem, texto_html)

        preenchido = driver.execute_script("return arguments[0].value || arguments[0].textContent || '';", elem)
        if preenchido and len(preenchido.strip()) >= min(20, len(texto_html)//2):
            return True
    except Exception:
        pass

    return False


def aguardar_sincronizacao_observacao(driver, texto_html: str, timeout: int = 8) -> bool:
    """Aguarda o editor/textarea refletir o texto inserido antes de clicar em Tramitar.
    Considera sincronizado quando o comprimento do conteúdo visível/oculto atinge
    pelo menos 60% do texto esperado (com mínimo de 30 caracteres).
    """
    min_len = max(15, int(len(texto_html) * 0.3))
    try:
        WebDriverWait(driver, timeout, poll_frequency=0.3).until(
            lambda d: d.execute_script(
                """
                var el = document.querySelector('body[contenteditable="true"], [contenteditable="true"]');
                var len1 = el ? (el.innerText || '').trim().length : 0;
                var ta = document.getElementById('ctl00_ContentToolBar_txtObservacao');
                var len2 = ta ? ((ta.value || ta.textContent || '').trim().length) : 0;
                return Math.max(len1, len2) >= arguments[0];
                """,
                min_len
            )
        )
        # pequenas pausas adicionais para estabilidade do DOM
        time.sleep(0.3)
        return True
    except Exception:
        return False


def forcar_sincronizacao_observacao(driver, texto_html: str) -> int:
    """Força a sincronização do texto do despacho nos campos submetidos pelo formulário.
    - Garante o HTML no editor contenteditable (se existir)
    - Preenche todos os inputs/textarea com id ou name contendo 'Observacao'/'txtObservacao'
    - Remove atributo 'disabled' para garantir que o valor seja submetido no postback
    Retorna a quantidade de elementos atualizados.
    """
    try:
        atualizados = driver.execute_script(
            r"""
            (function(html){
                function syncIn(doc){
                    var updated = 0;
                    try{
                        // Converte HTML para texto simples
                        var tmp = doc.createElement('div');
                        tmp.innerHTML = html;
                        var plain = (tmp.innerText || tmp.textContent || '').trim();

                        // 1) Todos os contenteditable
                        var eds = Array.from(doc.querySelectorAll('body[contenteditable="true"], [contenteditable="true"]'));
                        eds.forEach(function(ed){
                            try{
                                ed.innerHTML = html;
                                ['input','keyup','change','blur'].forEach(function(evt){ ed.dispatchEvent(new Event(evt,{bubbles:true})); });
                                updated++;
                            }catch(e){}
                        });

                        // 2) Campos com id/name contendo 'observa' (textarea, hidden, text)
                        var nodes = Array.from(doc.querySelectorAll('textarea, input[type="hidden"], input[type="text"]'))
                            .filter(function(n){
                                var id = (n.id||'').toLowerCase();
                                var nm = (n.name||'').toLowerCase();
                                return id.includes('observa') || nm.includes('observa');
                            });
                        nodes.forEach(function(n){
                            try{ n.removeAttribute('disabled'); }catch(e){}
                            try{ n.disabled = false; }catch(e){}
                            try{ n.value = plain; }catch(e){}
                            try{ if(n.tagName && n.tagName.toLowerCase()==='textarea'){ n.textContent = plain; } }catch(e){}
                            try{ n.setAttribute('value', plain); }catch(e){}
                            ['input','keyup','change','blur'].forEach(function(evt){ n.dispatchEvent(new Event(evt,{bubbles:true})); });
                            updated++;
                        });

                        // 3) Campo específico por ID
                        var ta = doc.getElementById('ctl00_ContentToolBar_txtObservacao');
                        if(ta){
                            try{ ta.removeAttribute('disabled'); }catch(e){}
                            try{ ta.disabled = false; }catch(e){}
                            try{ ta.value = plain; }catch(e){}
                            try{ ta.textContent = plain; }catch(e){}
                            try{ ta.setAttribute('value', plain); }catch(e){}
                            ['input','keyup','change','blur'].forEach(function(evt){ ta.dispatchEvent(new Event(evt,{bubbles:true})); });
                            updated++;
                        }

                        // 3.1) Campos por name com caminho qualificado
                        var byName = [];
                        try{ byName = Array.from(doc.getElementsByName('ctl00$ContentToolBar$txtObservacao')); }catch(e){}
                        if(byName && byName.length){
                            byName.forEach(function(n){
                                try{ n.removeAttribute('disabled'); }catch(e){}
                                try{ n.disabled = false; }catch(e){}
                                try{ n.value = plain; }catch(e){}
                                try{ n.textContent = plain; }catch(e){}
                                try{ n.setAttribute('value', plain); }catch(e){}
                                ['input','keyup','change','blur'].forEach(function(evt){ n.dispatchEvent(new Event(evt,{bubbles:true})); });
                                updated++;
                            });
                        }

                        return updated;
                    }catch(e){ return updated||0; }
                }

                var total = 0;
                // Documento principal
                total += syncIn(document);
                // Iframes mesma origem
                try{
                    for(var i=0;i<window.frames.length;i++){
                        try{
                            var fdoc = window.frames[i].document; // mesma origem
                            total += syncIn(fdoc);
                        }catch(e){}
                    }
                }catch(e){}

                // 4) Validadores WebForms (grupo vgTramitar)
                try{ if(window.Page_ClientValidate) Page_ClientValidate('vgTramitar'); }catch(e){}
                return total;
            })(arguments[0]);
            """,
            texto_html
        )
        return int(atualizados or 0)
    except Exception:
        return 0


def diagnosticar_observacao_campos(driver) -> str:
    """Coleta diagnóstico dos campos relacionados à observação para log.
    Retorna uma string resumida por campo: id|name|type|disabled|len(value).
    """
    try:
        info = driver.execute_script(
            r"""
            (function(){
                function len(v){ return (v||'').toString().trim().length; }
                function collect(doc, label){
                    var out=[];
                    var sel = 'textarea, input[type="hidden"], input[type="text"]';
                    var all = Array.from(doc.querySelectorAll(sel)).filter(function(n){
                        var id=(n.id||'').toLowerCase(); var nm=(n.name||'').toLowerCase();
                        return id.includes('observa') || nm.includes('observa');
                    });
                    all.forEach(function(n){
                        out.push([
                            label,
                            n.id||'', n.name||'', n.type||n.tagName, !!n.disabled,
                            len(n.value||n.textContent||'')
                        ].join('|'));
                    });
                    return out;
                }
                var res = collect(document, 'root');
                try{
                    for(var i=0;i<window.frames.length;i++){
                        try{ res = res.concat(collect(window.frames[i].document, 'frame'+i)); }catch(e){}
                    }
                }catch(e){}
                return res.join(';');
            })();
            """
        )
        return info or ''
    except Exception:
        return ''


def _e_pagina_resultado(driver) -> bool:
    """Heurística para detectar a página/aba de resultado (visualização de relatório/PDF)."""
    try:
        url = (driver.current_url or "")
        url_l = url.lower()
        if "/relatorios/visualizarelatorio.aspx" in url_l:
            return True
        if url_l.endswith(".pdf"):
            return True
        # Botão específico do relatório
        try:
            if driver.find_elements(By.ID, "btnFechar"):
                return True
        except Exception:
            pass
        # Link "Fechar" comum na barra do visualizador
        try:
            links = driver.find_elements(By.XPATH, "//a[normalize-space(.)='Fechar' or contains(.,'Fechar')]")
            if links:
                return True
        except Exception:
            pass
    except Exception:
        pass
    return False


def fechar_pagina_resultado(driver, wait, handles_antes: set | None = None, delay_seconds: float = 0.0, wait_new_tab_seconds: float = 8.0):
    """Fecha a aba/janela de resultado (se aberta) e retorna ao contexto original.
    - Se abriu nova aba/janela após a tramitação, fecha-a (aguardando opcionalmente alguns segundos antes).
    - Se navegou na mesma janela para o relatório, tenta clicar em 'Fechar' ou voltar (também com espera opcional).
    """
    try:
        try:
            main_handle = driver.current_window_handle
        except Exception:
            main_handle = None

        # 1) Espera por nova aba por até wait_new_tab_seconds
        try:
            base_set = set(handles_antes) if handles_antes else set()
        except Exception:
            base_set = set()
        if not base_set and main_handle:
            base_set = {main_handle}
        novas = []
        t0 = time.time()
        while time.time() - t0 < (wait_new_tab_seconds or 0):
            try:
                atuais = set(driver.window_handles)
            except Exception:
                atuais = set()
            novas = [h for h in atuais if h not in base_set]
            if novas:
                break
            time.sleep(0.25)

        # 2) Fecha diretamente as novas abas detectadas
        for h in novas:
            try:
                driver.switch_to.window(h)
                if delay_seconds and delay_seconds > 0:
                    time.sleep(delay_seconds)
                try:
                    url_now = (driver.current_url or '')
                except Exception:
                    url_now = ''
                try:
                    registrar_log(f"[Close] Fechando nova aba detectada (url='{url_now[:120]}') via driver.close()")
                except Exception:
                    pass
                try:
                    driver.close()
                except Exception:
                    pass
            finally:
                if main_handle:
                    try:
                        driver.switch_to.window(main_handle)
                    except Exception:
                        pass

        # 3) Varrida de segurança: fecha qualquer aba de relatório/PDF residual
        try:
            atuais_all = list(set(driver.window_handles))
        except Exception:
            atuais_all = []
        for h in atuais_all:
            if main_handle and h == main_handle:
                continue
            try:
                driver.switch_to.window(h)
                url_l = (driver.current_url or '').lower()
                if ('visualizarelatorio.aspx' in url_l) or url_l.endswith('.pdf') or ('/relatorios/' in url_l):
                    try:
                        registrar_log(f"[Close] Fechando aba residual de relatório (url='{url_l[:120]}')")
                    except Exception:
                        pass
                    try:
                        driver.close()
                    except Exception:
                        pass
            except Exception:
                pass
        try:
            if main_handle:
                driver.switch_to.window(main_handle)
        except Exception:
            pass

        # 4) Caso a navegação para o relatório tenha sido na mesma guia, tenta fechar/voltar
        try:
            if _e_pagina_resultado(driver):
                # Fechar por link/botão
                try:
                    btn = WebDriverWait(driver, 2, poll_frequency=0.2).until(
                        EC.element_to_be_clickable((By.XPATH, "//a[@id='btnFechar' or normalize-space(.)='Fechar' or contains(.,'Fechar')]"))
                    )
                    driver.execute_script("arguments[0].click();", btn)
                    time.sleep(0.2)
                except Exception:
                    try:
                        driver.execute_script("try{ if(window.__doPostBack){__doPostBack('btnFechar','');} }catch(e){}")
                        time.sleep(0.2)
                    except Exception:
                        pass
                # Últimos recursos
                try:
                    driver.back()
                except Exception:
                    try:
                        driver.get(f"{BASE_URL}/ProcessoBeneficio/ConProcessoBeneficio.aspx")
                    except Exception:
                        pass
        except Exception:
            pass
    except Exception as e:
        registrar_log(f"[Aviso] Falha ao fechar página de resultado: {e}")


def ja_dentro_do_setor(driver) -> bool:
    """Retorna True se a tela atual já é a lista com 'Processos a Receber'/'Dentro do Setor'."""
    try:
        if driver.find_elements(By.ID, "ctl00_ContentCampos_AccordionPane1_header_lblProcessoReceber"):
            return True
        if driver.find_elements(By.ID, "ctl00_ContentCampos_AccordionPane2_header_lblProcessoSetor"):
            return True
    except Exception:
        pass
    return False


def obter_estado_concessao(driver) -> str:
    """Retorna o estado atual da tela de Concessão.
    Valores possíveis:
      - 'selecionar_setor': dropdown de setor visível (precisa selecionar e clicar OK)
      - 'dentro_setor'    : listas 'Processos a Receber'/'Dentro do Setor' presentes
      - 'desconhecido'    : não conseguiu identificar
    """
    try:
        # Seletor de setor visível indica etapa de seleção pendente
        seletores = driver.find_elements(By.ID, "ctl00_ContentCampos_ddlSetor")
        for s in seletores:
            try:
                if s.is_displayed():
                    return "selecionar_setor"
            except Exception:
                continue
    except Exception:
        pass

    # Caso contrário, verifica se as caixas de processos existem (estado 'dentro do setor')
    try:
        if driver.find_elements(By.ID, "ctl00_ContentCampos_AccordionPane1_header_lblProcessoReceber"):
            return "dentro_setor"
        if driver.find_elements(By.ID, "ctl00_ContentCampos_AccordionPane2_header_lblProcessoSetor"):
            return "dentro_setor"
    except Exception:
        pass

    return "desconhecido"


def tramitar_para_presidente(driver, wait, nome_responsavel):
    """Tramita o processo para o Gabinete do Presidente."""
    try:
        # Clicar no botão Tramitar
        btn_tramitar = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentToolBar_btnTramitar")))
        driver.execute_script("arguments[0].scrollIntoView(true);", btn_tramitar)
        time.sleep(0.3)
        driver.execute_script("arguments[0].click();", btn_tramitar)
        # dá tempo do painel/modal de tramitação abrir por completo
        time.sleep(MODAL_OPEN_DELAY)

        # Selecionar Despacho
        select_despacho = Select(WebDriverWait(driver, 10, poll_frequency=0.3).until(
            EC.element_to_be_clickable((By.ID, "ctl00_ContentToolBar_ddlDespacho"))
        ))
        select_despacho.select_by_value("4")
        time.sleep(AFTER_SELECT_DELAY)

        # Selecionar Setor (Gabinete do Presidente)
        select_setor = Select(WebDriverWait(driver, 10, poll_frequency=0.3).until(
            EC.element_to_be_clickable((By.ID, "ctl00_ContentToolBar_ddlSetor"))
        ))
        select_setor.select_by_value("15")
        time.sleep(AFTER_SELECT_DELAY)

        # Corpo da tramitação (com assinatura do usuário logado, centralizada e em negrito)
        assinante = obter_assinante_nome() or nome_responsavel.upper()
        texto_tramitacao = (
            "<p>Ao Gabinete do Presidente Executivo,</p>"
            "<p>Encaminha-se, para assinatura, o ato constante da minuta anexa ao processo.</p>"
            "<p>Registre-se que a análise desta Unidade de Controle Interno quanto às concessões de aposentadoria, reserva remunerada, reforma e pensão, nos termos do Anexo VII da Instrução Normativa TCE nº 68, de 8 de dezembro de 2020, ainda depende de regulamentação específica, razão pela qual não houve emissão de parecer técnico sobre o presente ato.</p>"
            "<p>Respeitosamente,</p>"
            f"<div style='text-align:center; margin-top:12px;'><b>{assinante}</b></div>"
        )

        # Preenche o corpo (textarea/iframe/editor)
        if not preencher_editor_observacao(driver, wait, texto_tramitacao):
            registrar_log("[Aviso] Não foi possível preencher o corpo via editor; tentando fallback direto no campo por ID…")
            try:
                corpo = WebDriverWait(driver, 10, poll_frequency=0.3).until(
                    EC.element_to_be_clickable((By.ID, "ctl00_ContentToolBar_txtObservacao"))
                )
                driver.execute_script("arguments[0].scrollIntoView(true);", corpo)
                time.sleep(0.2)
                corpo.clear(); time.sleep(0.2)
                corpo.send_keys(texto_tramitacao)
            except Exception as e:
                registrar_log(f"[Erro] Falha ao preencher corpo da tramitação: {e}")
                raise

        # Aguarda sincronização do conteúdo com o editor/campo oculto antes de prosseguir
        if not aguardar_sincronizacao_observacao(driver, texto_tramitacao, timeout=SYNC_TIMEOUT_PRIMARY):
            registrar_log("[Aviso] Conteúdo do despacho pode não ter sincronizado totalmente; prosseguindo mesmo assim.")

        # Confirmar tramitação (somente se o texto estiver realmente presente)
        # Aguarda um curto período para sincronização e clica imediatamente em "Tramitar"
        len_ok = aguardar_sincronizacao_observacao(driver, texto_tramitacao, timeout=SYNC_TIMEOUT_PRECLICK)
        if not len_ok:
            registrar_log("[Aviso] Texto do despacho possivelmente ainda sincronizando; prosseguindo com clique em 'Tramitar'.")

        # Diagnóstico antes da sincronização forçada
        try:
            diag_antes = diagnosticar_observacao_campos(driver)
            if diag_antes:
                registrar_log(f"[Diag] Antes sync: {diag_antes}")
        except Exception:
            pass

        # Força sincronização final dos campos que serão enviados no postback
        try:
            qtd = forcar_sincronizacao_observacao(driver, texto_tramitacao)
            registrar_log(f"[Info] Campos de observação sincronizados (qtd={qtd}).")
            # Pequena pausa para JS de página reagir
            time.sleep(0.5)
        except Exception as e:
            registrar_log(f"[Aviso] Falha ao forçar sincronização final: {e}")

        # Diagnóstico depois da sincronização forçada
        try:
            diag_depois = diagnosticar_observacao_campos(driver)
            if diag_depois:
                registrar_log(f"[Diag] Depois sync: {diag_depois}")
        except Exception:
            pass

        # Verificação final: garante que o texto do despacho está presente antes do clique
        try:
            need_len = max(15, int(len(texto_tramitacao) * 0.3))
            js_len_script = """
                var ta=document.getElementById('ctl00_ContentToolBar_txtObservacao');
                var v=ta?(ta.value||ta.textContent||''):'';
                var ed=document.querySelector('body[contenteditable="true"], [contenteditable="true"]');
                var e=(ed?(ed.innerText||''):'');
                return Math.max(v.trim().length, e.trim().length);
            """
            plain_len = driver.execute_script(js_len_script)
            if (plain_len or 0) < need_len:
                registrar_log(f"[Aviso] Texto do despacho ainda curto (len={plain_len}); forçando sincronização extra...")
                try:
                    qtd2 = forcar_sincronizacao_observacao(driver, texto_tramitacao)
                    registrar_log(f"[Info] Sincronização extra aplicou em {qtd2} elemento(s).")
                    time.sleep(0.3)
                except Exception as e:
                    registrar_log(f"[Aviso] Falha ao aplicar sincronização extra: {e}")

                plain_len2 = driver.execute_script(js_len_script)
                if (plain_len2 or 0) < need_len:
                    # Fallback definitivo: cria/preenche um campo hidden com o nome esperado no formulário principal
                    try:
                        plain_text = re.sub(r'<[^>]+>', '', texto_tramitacao)
                        created = driver.execute_script(
                            r"""
                            (function(plain){
                                function ensureIn(doc){
                                    try{
                                        var name='ctl00$ContentToolBar$txtObservacao';
                                        var id='ctl00_ContentToolBar_txtObservacao';
                                        var ta = doc.getElementById(id);
                                        if(!ta){
                                            var byName = [];
                                            try{ byName = doc.getElementsByName(name); }catch(e){}
                                            if(byName && byName.length){ ta = byName[0]; }
                                        }
                                        if(!ta){
                                            var form = doc.querySelector('form');
                                            if(form){
                                                ta = doc.createElement('input');
                                                ta.type = 'hidden';
                                                ta.name = name;
                                                ta.id = id;
                                                form.appendChild(ta);
                                            }
                                        }
                                        if(ta){
                                            try{ ta.removeAttribute('disabled'); }catch(e){}
                                            try{ ta.disabled = false; }catch(e){}
                                            try{ ta.value = plain; }catch(e){}
                                            try{ ta.textContent = plain; }catch(e){}
                                            try{ ta.setAttribute('value', plain); }catch(e){}
                                            try{ ['input','keyup','change','blur'].forEach(function(evt){ ta.dispatchEvent(new Event(evt,{bubbles:true})); }); }catch(e){}
                                            return true;
                                        }
                                    }catch(e){}
                                    return false;
                                }
                                var ok = ensureIn(document);
                                try{
                                    for(var i=0;i<window.frames.length && !ok;i++){
                                        try{ ok = ensureIn(window.frames[i].document); }catch(e){}
                                    }
                                }catch(e){}
                                return !!ok;
                            })(plain);
                            """,
                            plain_text
                        )
                        if created:
                            registrar_log("[Info] Campo 'txtObservacao' criado/preenchido como hidden (fallback).")
                        else:
                            registrar_log("[Aviso] Não foi possível localizar/criar campo 'txtObservacao' nem em iframes.")
                        time.sleep(0.2)
                    except Exception as e:
                        registrar_log(f"[Aviso] Fallback (hidden) no 'txtObservacao' falhou: {e}")
        except Exception:
            pass

        # Encontra o botão por ID ou alternativas e clica
        try:
            btn_tramitar_final = WebDriverWait(driver, 6, poll_frequency=0.2).until(
                EC.element_to_be_clickable((By.ID, "ctl00_ContentToolBar_Button1"))
            )
        except Exception:
            # Fallbacks por texto/valor
            candidatos = driver.find_elements(By.XPATH,
                "//input[@type='submit' and (translate(@value,'TRAMITAR','tramitan')='tramitan' or contains(@value,'Tramitar'))]"
            ) or driver.find_elements(By.XPATH, "//button[normalize-space(.)='Tramitar']|//input[@value='Tramitar']")
            if not candidatos:
                raise RuntimeError("Botão 'Tramitar' não encontrado.")
            btn_tramitar_final = candidatos[0]

        # Guarda as janelas/abas atuais para detectar novas após a tramitação
        try:
            handles_antes = set(driver.window_handles)
        except Exception:
            handles_antes = None

        driver.execute_script("arguments[0].scrollIntoView(true);", btn_tramitar_final)
        time.sleep(0.3)
        driver.execute_script("arguments[0].click();", btn_tramitar_final)
        time.sleep(0.8)

        # Se o alerta não aparecer rapidamente, força o postback da página ASP.NET
        try:
            WebDriverWait(driver, 3).until(EC.alert_is_present())
        except Exception:
            registrar_log("[Aviso] Alerta não apareceu após o clique; tentando __doPostBack...")
            driver.execute_script(
                "if(window.WebForm_DoPostBackWithOptions){WebForm_DoPostBackWithOptions(new WebForm_PostBackOptions('ctl00$ContentToolBar$Button1','',true,'vgTramitar','',false,false));}"
            )
            try:
                WebDriverWait(driver, 4).until(EC.alert_is_present())
            except Exception:
                registrar_log("[Aviso] __doPostBack não abriu alerta; tentando __doPostBack simples...")
                driver.execute_script("if(window.__doPostBack){__doPostBack('ctl00$ContentToolBar$Button1','');}")
                time.sleep(0.8)
        
        # Trata o alerta após tramitar
        try:
            WebDriverWait(driver, 10).until(EC.alert_is_present())
            alerta = driver.switch_to.alert
            msg_alerta = alerta.text
            registrar_log(f"[Alerta] {msg_alerta}")
            alerta.accept()
            time.sleep(0.8)
        except Exception:
            pass

        # Fecha a página/aba de resultado (se aberta) e retorna para prosseguir
        try:
            # Fecha o relatório imediatamente (sem aguardar), usando o XPath fornecido para ser mais rápido
            fechar_pagina_resultado(driver, wait, handles_antes, delay_seconds=0)
        except Exception as e:
            registrar_log(f"[Aviso] Não foi possível fechar a página de resultado automaticamente: {e}")

        registrar_log(f"[OK] Processo tramitado com sucesso")
        print(f"[OK] Processo tramitado com sucesso")

    except Exception as e:
        registrar_log(f"[Erro] Falha ao tramitar processo: {e}")
        print(f"[Erro] Falha ao tramitar processo: {e}")
        raise


def automatizar(responsavel, cpf):
    driver = None
    try:
        atualizar_status("🚀 Iniciando automação...")
        progress.set(0.05)
        registrar_log(f"Iniciado por {USUARIO_PC}")
        registrar_log(f"Responsável Controle Interno: {responsavel} - {cpf}")

        atualizar_status("🔗 Conectando ao Chrome...")
        progress.set(0.1)
        
        import logging
        logging.getLogger('selenium').setLevel(logging.WARNING)
        
        chrome_options = Options()
        chrome_options.debugger_address = "localhost:9222"

        # Garante que o Chrome está rodando com depuração em 9222
        if not porta_debug_aberta():
            msg = (
                "Chrome não encontrado na porta 9222. Abra o Chrome com depuração ativa:\n"
                    r'"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe" --remote-debugging-port=9222 --user-data-dir=\\ChromeDevSession'
            )
            registrar_log(f"[Erro] {msg}")
            raise RuntimeError(msg)
            
        registrar_log("Conectando ao Chrome via Selenium Manager...")
        atualizar_status("⏳ Conectando via Selenium Manager…")
        t0 = time.time()
        driver = webdriver.Chrome(options=chrome_options)
        registrar_log(f"Chrome conectado em {time.time()-t0:.2f}s")
        

        driver.implicitly_wait(2)
        wait = WebDriverWait(driver, 5, poll_frequency=0.3)
        
        registrar_log("Chrome conectado com sucesso")
        atualizar_status("✅ Chrome conectado!")
        progress.set(0.15)

        # ========== 1️⃣ Abrir Benefício > Concessão ==========
        atualizar_status("📂 Acessando módulo Benefício → Concessão...")
        progress.set(0.2)
        registrar_log("Tentando abrir Concessão por URL direta (com fallback no menu)...")

        if not abrir_concessao(driver, wait):
            raise RuntimeError("Não foi possível abrir a tela de Concessão.")
        
        # ========== 2️⃣ Selecionar setor (ou pular se já estiver dentro) ==========
        estado = obter_estado_concessao(driver)
        if estado == "dentro_setor":
            registrar_log("Detectado que já estamos dentro do setor; pulando seleção de setor.")
            atualizar_status("🏢 Setor já selecionado. Continuando…")
        elif estado == "selecionar_setor":
            atualizar_status("🏢 Selecionando setor UECI…")
            progress.set(0.3)
            registrar_log("Aguardando seletor de setor carregar…")

            try:
                seletor_setor = WebDriverWait(driver, 8, poll_frequency=0.2).until(
                    EC.presence_of_element_located((By.ID, "ctl00_ContentCampos_ddlSetor"))
                )
            except Exception:
                # Se o seletor sumiu, reavalia o estado; talvez já entrou no setor
                if obter_estado_concessao(driver) == "dentro_setor":
                    registrar_log("Seletor não está mais presente; já entramos no setor. Prosseguindo…")
                else:
                    raise

            if obter_estado_concessao(driver) == "selecionar_setor":
                try:
                    sel = Select(seletor_setor)
                    try:
                        registrar_log("Seletor encontrado, selecionando UECI (valor 59)…")
                        sel.select_by_value("59")
                    except NoSuchElementException:
                        opts = [o for o in sel.options if 'UECI' in (o.text or '').upper()]
                        if opts:
                            registrar_log("Opção 59 não encontrada; selecionando opção que contém 'UECI'.")
                            sel.select_by_visible_text(opts[0].text)
                        else:
                            registrar_log("[Aviso] Opção de setor UECI não encontrada no seletor.")

                    # Clica OK para confirmar a entrada no setor
                    try:
                        registrar_log("Procurando botão OK…")
                        btn_ok = driver.find_element(By.ID, "ctl00_ContentCampos_Button1")
                        registrar_log("Botão OK encontrado, clicando…")
                        driver.execute_script("arguments[0].click();", btn_ok)
                        # Aguarda as caixas de processos aparecerem
                        WebDriverWait(driver, 10, poll_frequency=0.3).until(
                            lambda d: d.find_elements(By.ID, "ctl00_ContentCampos_AccordionPane2_header_lblProcessoSetor")
                                     or d.find_elements(By.ID, "ctl00_ContentCampos_AccordionPane1_header_lblProcessoReceber")
                        )
                        registrar_log("Setor UECI selecionado com sucesso!")
                    except Exception as e:
                        # Reavalia o estado: se já estiver dentro, segue; caso contrário, propaga o erro
                        if obter_estado_concessao(driver) == "dentro_setor":
                            registrar_log(f"[Aviso] Não foi possível acionar OK, porém já estamos dentro do setor: {e}")
                        else:
                            raise
                except Exception as e:
                    # Se falhou selecionar, mas já estiver dentro, segue
                    if obter_estado_concessao(driver) == "dentro_setor":
                        registrar_log(f"[Aviso] Falha ao selecionar setor, mas já estamos dentro: {e}")
                    else:
                        raise
        else:
            registrar_log("[Aviso] Estado da tela de Concessão não identificado claramente. Prosseguindo com melhor esforço…")

        # ========== 3️⃣ Processos a Receber ==========
        atualizar_status("📦 Verificando processos a receber...")
        progress.set(0.4)
        try:
            # Expande a seção "30 Últimos Processos a Receber"
            btn_receber_box = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentCampos_AccordionPane1_header_lblProcessoReceber")))
            driver.execute_script("arguments[0].click();", btn_receber_box)
            time.sleep(0.8)

            # Busca novamente os checkboxes após abrir
            caixas = driver.find_elements(By.XPATH, "//input[contains(@id,'chk_receber')]")

            if caixas:
                # Descobre o índice da coluna "Setor Enviou" pelo cabeçalho da tabela (quando possível)
                col_setor = -1
                try:
                    tabela = driver.find_element(By.XPATH, "//table[.//input[contains(@id,'chk_receber')]]")
                    cabec = tabela.find_elements(By.XPATH, ".//tr[1]/*[self::th or self::td]")
                    for idx, th in enumerate(cabec):
                        txt = (th.text or "").strip().lower()
                        if "setor" in txt and "enviou" in txt:
                            col_setor = idx
                            break
                except Exception:
                    col_setor = -1

                atualizar_status(f"Encontrados {len(caixas)} processo(s) a receber. Marcando conforme regra...")
                for i in range(len(caixas)):
                    try:
                        # Rebusca o elemento a cada iteração para evitar stale reference
                        caixa = driver.find_elements(By.XPATH, "//input[contains(@id,'chk_receber')]")[i]

                        # Obtém o texto da coluna "Setor Enviou" correspondente à linha deste checkbox
                        setor_txt = ""
                        try:
                            linha = caixa.find_element(By.XPATH, "./ancestor::tr[1]")
                            if col_setor >= 0:
                                tds = linha.find_elements(By.XPATH, "./td")
                                if col_setor < len(tds):
                                    setor_txt = (tds[col_setor].text or "").strip()
                            if not setor_txt:
                                setor_txt = (linha.text or "").strip()
                        except Exception as e:
                            registrar_log(f"[Aviso] Falha ao obter 'Setor Enviou' da linha: {e}")

                        # Normaliza e aplica regra de bloqueio (considera variações como C.P.A.D, C P A D, etc.)
                        s = (setor_txt or "").upper()
                        s_norm = _normalize_text(setor_txt)

                        bloquear_cpad = ("CPAD" in s) or ("CPAD" in s_norm)
                        bloquear_coord = (
                            ("COORDENA" in s and "PROTOCOLO" in s and "ARQUIVO" in s and "DOCUMENTAL" in s)
                            or ("COORDENACAO" in s_norm and "PROTOCOLO" in s_norm and "ARQUIVO" in s_norm and "DOCUMENTAL" in s_norm)
                        )
                        bloquear = bloquear_cpad or bloquear_coord

                        if bloquear:
                            registrar_log(f"[Skip] Processo NÃO recebido (Setor Enviou='{setor_txt}')")
                            continue

                        # Marca a caixa para receber
                        driver.execute_script("arguments[0].scrollIntoView(true);", caixa)
                        driver.execute_script("arguments[0].click();", caixa)
                        time.sleep(0.1)
                    except Exception as e:
                        registrar_log(f"[Aviso] Erro ao analisar/marcar caixa {i+1}: {e}")

                # Clicar no botão "Receber Processos Selecionados"
                try:
                    btn_receber_lote = wait.until(
                        EC.element_to_be_clickable((By.ID, "ctl00_ContentCampos_AccordionPane1_content_imgBtnRecebeLote"))
                    )
                    driver.execute_script("arguments[0].click();", btn_receber_lote)
                    atualizar_status("Aguardando confirmação do recebimento...")
                    
                    # Espera e trata o alerta "Processo recebido com sucesso!"
                    try:
                        WebDriverWait(driver, 5).until(EC.alert_is_present())
                        alerta = driver.switch_to.alert
                        msg = alerta.text
                        registrar_log(f"[Alerta] {msg}")
                        alerta.accept()
                        atualizar_status("Processos recebidos com sucesso.")
                        time.sleep(0.5)
                    except Exception:
                        atualizar_status("Nenhum alerta exibido após o recebimento.")
                        time.sleep(0.3)
                except Exception as e:
                    atualizar_status(f"Falha ao clicar em 'Receber Processos Selecionados': {e}")

            else:
                atualizar_status("Nenhum processo encontrado para receber.")

        except Exception as e:
            atualizar_status(f"Erro ao acessar a caixa de 'Processos a Receber': {e}")

        # -------------- ETAPA 4: PROCESSOS DENTRO DO SETOR ----------------
        atualizar_status("⚙️ Processando processos dentro do setor...")
        print("Verificando processos dentro do setor...")
        try:
            btn_dentro_setor = wait.until(
                EC.element_to_be_clickable((By.ID, "ctl00_ContentCampos_AccordionPane2_header_lblProcessoSetor"))
            )
            driver.execute_script("arguments[0].click();", btn_dentro_setor)
            time.sleep(0.8)

            # Captura todos os botões "Editar" e "Abrir"
            botoes = driver.find_elements(By.XPATH,
                "//input[contains(@id,'AccordionPane2_content_grdProcessoSetor') and (contains(@id,'imgbtnEdit') or contains(@id,'imgbtnAbrir'))]"
            )

            if botoes:
                print(f"Encontrados {len(botoes)} processo(s) dentro do setor. Iniciando tramitação...")

                for index, botao in enumerate(botoes, start=1):
                    try:
                        try:
                            driver.execute_script("arguments[0].scrollIntoView(true);", botao)
                            driver.execute_script("arguments[0].click();", botao)
                        except:
                            # Recarrega e tenta clicar de novo caso o elemento tenha sumido
                            botoes = driver.find_elements(By.XPATH,
                                "//input[contains(@id,'AccordionPane2_content_grdProcessoSetor') and (contains(@id,'imgbtnEdit') or contains(@id,'imgbtnAbrir'))]"
                            )
                            if len(botoes) >= index:
                                botao = botoes[index - 1]
                                driver.execute_script("arguments[0].scrollIntoView(true);", botao)
                                driver.execute_script("arguments[0].click();", botao)

                        time.sleep(1)

                        # Preenche campos e tramita
                        preencher_informacoes_controle_interno(driver, wait, responsavel, cpf)
                        tramitar_para_presidente(driver, wait, responsavel)

                        # Aguarda retorno para a lista principal
                        driver.back()
                        time.sleep(1)

                        # Recarrega lista para evitar stale elements
                        btn_dentro_setor = wait.until(
                            EC.element_to_be_clickable((By.ID, "ctl00_ContentCampos_AccordionPane2_header_lblProcessoSetor"))
                        )
                        driver.execute_script("arguments[0].click();", btn_dentro_setor)
                        time.sleep(0.8)

                        botoes = driver.find_elements(By.XPATH,
                            "//input[contains(@id,'AccordionPane2_content_grdProcessoSetor') and (contains(@id,'imgbtnEdit') or contains(@id,'imgbtnAbrir'))]"
                        )

                    except Exception as e:
                        print(f"[Erro] Falha ao tramitar processo {index}: {e}")
                        driver.get(f"{BASE_URL}/ProcessoBeneficio/ConProcessoBeneficio.aspx")
                        time.sleep(3)
                        continue

                print("Todos os processos foram tramitados com sucesso!")

            else:
                # Nenhum processo na caixa do setor → avisa e encerra após 5s
                registrar_log("Nenhum processo encontrado na caixa 'Dentro do Setor'. Encerrando em 5s...")
                atualizar_status("Nenhum processo na caixa do setor.")
                mostrar_aviso_e_encerrar("não há mais processos para tramitação na caixa do setor", segundos=5)
                return

        except Exception as e:
            print(f"Erro ao acessar a caixa de 'Processos dentro do Setor': {e}")

            # ========== 5️⃣ Mais informações ==========
            aba_info = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='__tab_ctl00_ContentCampos_TabContainer1_tabTCE']")))
            aba_info.click()
            time.sleep(6)

            # Parecer Controle Interno
            select_parecer = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentCampos_TabContainer1_tabTCE_parecerControleInternoTCE")))
            Select(select_parecer).select_by_visible_text("Não foi objeto do exame")

            # CPF e Nome
            driver.execute_script("document.getElementById('ctl00_ContentCampos_TabContainer1_tabTCE_txtCPFRespControleInternoTCE').removeAttribute('disabled');")
            driver.execute_script("document.getElementById('ctl00_ContentCampos_TabContainer1_tabTCE_txtNomeRespControleInternoTCE').removeAttribute('disabled');")
            campo_cpf = driver.find_element(By.ID, "ctl00_ContentCampos_TabContainer1_tabTCE_txtCPFRespControleInternoTCE")
            campo_nome = driver.find_element(By.ID, "ctl00_ContentCampos_TabContainer1_tabTCE_txtNomeRespControleInternoTCE")
            campo_cpf.clear()
            campo_cpf.send_keys(cpf)
            campo_nome.clear()
            campo_nome.send_keys(responsavel)

            btn_salvar = driver.find_element(By.ID, "ctl00_ContentToolBar_btnSalvar")
            btn_salvar.click()
            time.sleep(2)

            # ========== 6️⃣ Tramitar ==========
            btn_tramitar = driver.find_element(By.ID, "ctl00_ContentToolBar_btnTramitar")
            btn_tramitar.click()
            time.sleep(3)

            # Despacho e Setor
            select_despacho = Select(wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentToolBar_ddlDespacho"))))
            select_despacho.select_by_value("4")

            select_setor = Select(wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentToolBar_ddlSetor"))))
            select_setor.select_by_value("15")

            # Corpo da tramitação
            texto_tramitacao = f"""Ao Gabinete do Presidente Executivo,

Encaminha-se, para assinatura, o ato constante da minuta anexa ao processo.

Registre-se que a análise desta Unidade de Controle Interno quanto às concessões de aposentadoria, reserva remunerada, reforma e pensão, nos termos do Anexo VII da Instrução Normativa TCE nº 68, de 8 de dezembro de 2020, ainda depende de regulamentação específica, razão pela qual não houve emissão de parecer técnico sobre o presente ato.

Respeitosamente,

    {responsavel.upper()}"""

            corpo = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentToolBar_txtObservacao")))
            corpo.clear()
            corpo.send_keys(texto_tramitacao)

            btn_tramitar_final = driver.find_element(By.ID, "ctl00_ContentToolBar_Button1")
            btn_tramitar_final.click()
            time.sleep(2)
            
            # Trata o alerta após tramitar
            try:
                WebDriverWait(driver, 5).until(EC.alert_is_present())
                alerta = driver.switch_to.alert
                print(f"[Alerta] {alerta.text}")
                alerta.accept()
                time.sleep(1)
            except:
                pass

            registrar_log(f"Processo {i} tramitado com sucesso.")
            driver.back()
            time.sleep(3)

        progress.set(1.0)
        atualizar_status("✅ Todos os processos foram tramitados com sucesso!")
        registrar_log("Todos os processos concluídos com sucesso.")

    except Exception as e:
        atualizar_status(f"❌ Erro: {str(e)}")
        registrar_log(f"Erro: {str(e)}")
        progress.set(0)
    finally:
        # Garante encerramento do ChromeDriver para evitar arquivos em uso no _MEI* (PyInstaller)
        try:
            if driver:
                driver.quit()
        except Exception:
            pass

# ==============================
# INTERFACE GRÁFICA MODERNA
# ==============================

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

root = ctk.CTk()
root.title("UECI Automação - SISPREV Inteligente")
root.geometry("900x700")
root.resizable(False, False)

# Configurar cor de fundo com gradiente simulado
root.configure(fg_color=("#E8EAF6", "#1A1A2E"))

# ========== CABEÇALHO COM ESTILO ==========
header_frame = ctk.CTkFrame(root, fg_color=("#C5CAE9", "#2D2D44"), corner_radius=20, height=130)
header_frame.pack(fill="x", padx=20, pady=(20, 10))
header_frame.pack_propagate(False)

# Ícone decorativo
icon_label = ctk.CTkLabel(
    header_frame, 
    text="✨", 
    font=ctk.CTkFont(size=45),
    text_color=("#7E57C2", "#BB86FC")
)
icon_label.pack(pady=(10, 0))

title_label = ctk.CTkLabel(
    header_frame, 
    text=f"Olá, {USUARIO_PC}! 💜", 
    font=ctk.CTkFont(size=24, weight="bold"),
    text_color=("#5E35B1", "#E1BEE7")
)
title_label.pack(pady=(5, 0))

subtitle_label = ctk.CTkLabel(
    header_frame,
    text="Sistema Inteligente de Automação UECI",
    font=ctk.CTkFont(size=13),
    text_color=("#9575CD", "#CE93D8")
)
subtitle_label.pack(pady=(2, 0))

# ========== CARD PRINCIPAL ==========
main_card = ctk.CTkFrame(root, fg_color=("#F3E5F5", "#2D2D44"), corner_radius=20)
main_card.pack(fill="both", expand=True, padx=20, pady=10)

# Container interno com padding
inner_container = ctk.CTkFrame(main_card, fg_color="transparent")
inner_container.pack(fill="both", expand=True, padx=30, pady=20)

# ========== SEÇÃO DE SELEÇÃO ==========
selection_frame = ctk.CTkFrame(inner_container, fg_color=("#EDE7F6", "#383854"), corner_radius=15)
selection_frame.pack(fill="x", pady=(0, 15))

selection_label = ctk.CTkLabel(
    selection_frame,
    text="👤  Responsável pelo Controle Interno no SISPREV",
    font=ctk.CTkFont(size=15, weight="bold"),
    text_color=("#6A1B9A", "#CE93D8")
)
selection_label.pack(pady=(15, 8))

responsavel_var = ctk.StringVar(value="Carla Zambi Meirelles")
responsavel_menu = ctk.CTkOptionMenu(
    selection_frame,
    values=list(RESPONSAVEIS.keys()),
    variable=responsavel_var,
    width=350,
    height=40,
    corner_radius=10,
    font=ctk.CTkFont(size=14),
    fg_color=("#9C27B0", "#7B1FA2"),
    button_color=("#7B1FA2", "#9C27B0"),
    button_hover_color=("#6A1B9A", "#AB47BC"),
    dropdown_fg_color=("#E1BEE7", "#4A4A6A"),
    dropdown_hover_color=("#CE93D8", "#5E5E7E"),
    dropdown_text_color=("#4A148C", "#E1BEE7")
)
responsavel_menu.pack(pady=(0, 15))

# ========== BARRA DE PROGRESSO MODERNA ==========
progress_frame = ctk.CTkFrame(inner_container, fg_color="transparent")
progress_frame.pack(fill="x", pady=(0, 15))

progress_label = ctk.CTkLabel(
    progress_frame,
    text="Progresso da Automação",
    font=ctk.CTkFont(size=13, weight="bold"),
    text_color=("#7E57C2", "#BB86FC")
)
progress_label.pack(anchor="w", pady=(0, 8))

progress = ctk.CTkProgressBar(
    progress_frame,
    width=400,
    height=20,
    corner_radius=10,
    fg_color=("#D1C4E9", "#3D3D5C"),
    progress_color=("#9C27B0", "#BB86FC")
)
progress.pack(fill="x")
progress.set(0)

# ========== STATUS COM DESIGN ELEGANTE ==========
status_frame = ctk.CTkFrame(
    inner_container,
    fg_color=("#EDE7F6", "#383854"),
    corner_radius=15,
    height=80
)
status_frame.pack(fill="x", pady=(15, 20))
status_frame.pack_propagate(False)

status_icon = ctk.CTkLabel(
    status_frame,
    text="💫",
    font=ctk.CTkFont(size=30)
)
status_icon.pack(side="left", padx=(20, 10))

status_label = ctk.CTkLabel(
    status_frame,
    text="Aguardando início da automação...",
    font=ctk.CTkFont(size=14),
    text_color=("#9575CD", "#CE93D8")
)
status_label.pack(side="left", padx=10, expand=True)

# ========== BOTÕES COM DESIGN MODERNO ==========
buttons_frame = ctk.CTkFrame(inner_container, fg_color="transparent")
buttons_frame.pack(pady=(10, 0))

def iniciar():
    resp = responsavel_var.get()
    cpf = RESPONSAVEIS[resp]
    progress.set(0)
    atualizar_status("🔄 Preparando automação...")
    btn_iniciar.configure(state="disabled", text="⏳ Processando...")
    
    def executar():
        try:
            automatizar(resp, cpf)
        finally:
            btn_iniciar.configure(state="normal", text="🚀 Iniciar Automação")
    
    threading.Thread(target=executar, daemon=True).start()

def abrir_logs():
    if os.path.exists(LOG_FILE):
        os.startfile(LOG_FILE)
    else:
        registrar_log("Arquivo de log criado.")
        os.startfile(LOG_FILE)

btn_iniciar = ctk.CTkButton(
    buttons_frame,
    text="🚀 Iniciar Automação",
    command=iniciar,
    width=200,
    height=50,
    corner_radius=15,
    font=ctk.CTkFont(size=15, weight="bold"),
    fg_color=("#9C27B0", "#BB86FC"),
    hover_color=("#7B1FA2", "#CE93D8"),
    text_color=("#FFFFFF", "#1A1A2E")
)
btn_iniciar.grid(row=0, column=0, padx=10, pady=5)

btn_logs = ctk.CTkButton(
    buttons_frame,
    text="📂 Visualizar Logs",
    command=abrir_logs,
    width=200,
    height=50,
    corner_radius=15,
    font=ctk.CTkFont(size=15, weight="bold"),
    fg_color=("#5E35B1", "#9575CD"),
    hover_color=("#4A148C", "#B39DDB"),
    text_color=("#FFFFFF", "#1A1A2E")
)
btn_logs.grid(row=0, column=1, padx=10, pady=5)

btn_sair = ctk.CTkButton(
    buttons_frame,
    text="⏹️ Encerrar Sistema",
    command=root.destroy,
    width=420,
    height=45,
    corner_radius=15,
    font=ctk.CTkFont(size=14, weight="bold"),
    fg_color=("#E91E63", "#C2185B"),
    hover_color=("#C2185B", "#AD1457"),
    text_color="#FFFFFF"
)
btn_sair.grid(row=1, column=0, columnspan=2, padx=10, pady=(10, 0))

# ========== RODAPÉ ==========
footer_label = ctk.CTkLabel(
    root,
    text="Desenvolvido com 💜 para facilitar seu dia a dia",
    font=ctk.CTkFont(size=11),
    text_color=("#9575CD", "#B39DDB")
)
footer_label.pack(pady=(0, 15))


root.mainloop()
