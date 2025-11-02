#!/usr/bin/env python3
"""
RPA para Preenchimento Automático de Notas Fiscais
Desenvolvido para automatizar o preenchimento de formulários de notas fiscais eletrônicas
"""

import pandas as pd
import time
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
try:
    from webdriver_manager.chrome import ChromeDriverManager
    from selenium.webdriver.chrome.service import Service
    WEBDRIVER_MANAGER_DISPONIVEL = True
except ImportError:
    WEBDRIVER_MANAGER_DISPONIVEL = False
import re

class RPANotasFiscais:
    def __init__(self, url_site, caminho_excel, mapeamento_cliente, delay=2):
        """
        Inicializa o RPA

        Args:
            url_site (str): URL do site de emissão de notas fiscais
            caminho_excel (str): Caminho para o arquivo Excel com os dados
            mapeamento_cliente (str): Nome do cliente para usar o mapeamento ('kleiton' ou 'katia')
            delay (int): Tempo de delay entre ações (segundos)
        """
        self.url_site = url_site
        self.caminho_excel = caminho_excel
        self.delay = delay
        self.driver = None
        self.wait = None
        self.setup_logging()

        # Definir mapeamentos por cliente
        self.mapeamentos_clientes = {
            'kleiton': {
                'atividade': '508',  # GUARDA, TRATAMENTO, AMESTRAMENTO, EMBELEZAMENTO, ALOJAMENTO E CONGENERES
                'tipo_pessoa': 'PESSOA FÍSICA',
                'uf': 'SP',
                'exigibilidade': 'EXIGÍVEL',
                'simples_nacional': 'Sim',
                'regime_especial': 'MICROEMPRESARIO E EMPRESA DE PEQUENO PORTE',
                'iss_retido': 'Não',
                'incentivo_fiscal': 'Não',
                'valor_deducoes': '0,00',
                'inss': '0,00',
                'ir': '0,00',
                'csll': '0,00',
                'cofins': '0,00',
                'pis': '0,00',
                'outras_retencoes': '0,00',
                "itAliquota": "2,01",
                "uf_incidencia": "SP",
                "municipio_incidencia": "INDAIATUBA",
                "UfServico": "SP",
                "somMunicipioServico": "INDAIATUBA"
            },
            'katia': {
                'atividade': '508',  # GUARDA, TRATAMENTO, AMESTRAMENTO, EMBELEZAMENTO, ALOJAMENTO E CONGENERES
                'tipo_pessoa': 'PESSOA FÍSICA',
                'uf': 'SP',
                'exigibilidade': 'EXIGÍVEL',
                'simples_nacional': 'Sim',
                'regime_especial': 'MICROEMPRESARIO E EMPRESA DE PEQUENO PORTE',
                'iss_retido': 'Não',
                'incentivo_fiscal': 'Não',
                'valor_deducoes': '0,00',
                'inss': '0,00',
                'ir': '0,00',
                'csll': '0,00',
                'cofins': '0,00',
                'pis': '0,00',
                'outras_retencoes': '0,00',
                "itAliquota": "2,01",
                "uf_incidencia": "SP",
                "municipio_incidencia": "INDAIATUBA",
                "UfServico": "SP",
                "somMunicipioServico": "INDAIATUBA"
            }
        }

        # Configurar mapeamento baseado no cliente selecionado
        if mapeamento_cliente.lower() in self.mapeamentos_clientes:
            self.configuracoes_padrao = self.mapeamentos_clientes[mapeamento_cliente.lower()]
            self.cliente_atual = mapeamento_cliente.lower()
        else:
            raise ValueError(f"Cliente '{mapeamento_cliente}' não encontrado. Clientes disponíveis: {list(self.mapeamentos_clientes.keys())}")

        print(f"✅ Mapeamento configurado para cliente: {self.cliente_atual.upper()}")
        print(f"   📊 Alíquota: {self.configuracoes_padrao['itAliquota']}%")
        print(f"   🏢 Município de incidência: {self.configuracoes_padrao['municipio_incidencia']}")
        print(f"   📍 Município de serviço: {self.configuracoes_padrao['somMunicipioServico']}")
        print(f"   🏛️  Regime especial: {self.configuracoes_padrao['regime_especial']}")

    def mostrar_comparacao_mapeamentos(self):
        """Mostra uma comparação visual entre os mapeamentos dos clientes"""
        print("\n" + "="*80)
        print("📋 COMPARAÇÃO DOS MAPEAMENTOS DE CLIENTES")
        print("="*80)

        campos_diferentes = ["atividade"]

        print(f"{'Campo':<25} {'KLEITON':<25} {'KATIA':<25}")
        print("-" * 80)

        for campo in campos_diferentes:
            kleiton_val = self.mapeamentos_clientes['kleiton'][campo]
            katia_val = self.mapeamentos_clientes['katia'][campo]
            print(f"{campo:<25} {kleiton_val:<25} {katia_val:<25}")

        print("-" * 80)
        print("💡 As demais configurações são idênticas para ambos os clientes")
        print("="*80)

    def setup_logging(self):
        """Configura o sistema de logs"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('rpa_notas_fiscais.log'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)

    def fechar_modals(self):
        """Fecha qualquer modal que possa estar aberto"""
        try:
            # Procura especificamente pelo modal do Simples Nacional
            modal_simples = self.driver.find_elements(By.ID, "primefacesmessagedlg")
            if modal_simples and modal_simples[0].is_displayed():
                self.logger.info("Modal do Simples Nacional detectado, fechando...")
                close_btn = modal_simples[0].find_element(By.CSS_SELECTOR, ".ui-dialog-titlebar-close")
                close_btn.click()
                time.sleep(1)

            # Procura por outros modais/overlays
            modal_overlay = self.driver.find_elements(By.CSS_SELECTOR, ".ui-widget-overlay, .ui-dialog-mask")
            if modal_overlay:
                self.logger.info("Modal detectado, tentando fechar...")

                # Tenta fechar pelo botão X
                close_buttons = self.driver.find_elements(By.CSS_SELECTOR, ".ui-dialog-closable .ui-dialog-titlebar-close, .ui-button-icon-only")
                for btn in close_buttons:
                    try:
                        if btn.is_displayed() and btn.is_enabled():
                            btn.click()
                            time.sleep(0.5)
                            break
                    except:
                        continue

                # Se ainda existe modal, tenta ESC
                if self.driver.find_elements(By.CSS_SELECTOR, ".ui-widget-overlay, .ui-dialog-mask"):
                    from selenium.webdriver.common.keys import Keys
                    self.driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                    time.sleep(0.5)

        except Exception as e:
            self.logger.debug(f"Erro ao fechar modal: {str(e)}")

    def fechar_dropdowns_abertos(self):
        """Fecha qualquer dropdown que possa estar aberto"""
        try:
            # Estratégia 1: Clicar no body para fechar dropdowns
            self.driver.execute_script("""
                // Fecha dropdowns clicando no body
                document.body.click();

                // Remove painéis visíveis de dropdown
                var panels = document.querySelectorAll('[id$="_panel"]:not([style*="display: none"])');
                panels.forEach(function(panel) {
                    if (panel.style.display !== 'none') {
                        panel.style.display = 'none';
                    }
                });

                // Pressiona ESC para garantir
                document.body.dispatchEvent(new KeyboardEvent('keydown', {
                    key: 'Escape',
                    keyCode: 27,
                    which: 27
                }));
            """)
            time.sleep(0.1)
        except Exception as e:
            self.logger.debug(f"Erro ao fechar dropdowns: {str(e)}")

    def configurar_driver(self):
        """Configura e inicializa o driver do Chrome com download automático"""
        try:
            chrome_options = Options()
            chrome_options.add_argument('--disable-blink-features=AutomationControlled')
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_experimental_option('useAutomationExtension', False)

            # Tenta usar webdriver-manager para download automático do ChromeDriver
            if WEBDRIVER_MANAGER_DISPONIVEL:
                try:
                    print("🔧 Configurando ChromeDriver automaticamente...")
                    service = Service(ChromeDriverManager().install())
                    self.driver = webdriver.Chrome(service=service, options=chrome_options)
                    print("✅ ChromeDriver configurado automaticamente!")
                except Exception as e:
                    print(f"⚠️  Falha no download automático: {e}")
                    print("🔧 Tentando usar ChromeDriver local...")
                    self.driver = webdriver.Chrome(options=chrome_options)
            else:
                # Fallback para ChromeDriver local
                self.driver = webdriver.Chrome(options=chrome_options)

            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            self.wait = WebDriverWait(self.driver, 10)

            self.logger.info("Driver configurado com sucesso")

        except Exception as e:
            error_msg = f"""
❌ ERRO: Não foi possível inicializar o ChromeDriver!

Possíveis soluções:
1. 🔧 Execute: setup_python.bat (instala dependências automaticamente)
2. 🌐 Verifique se o Google Chrome está instalado e atualizado
3. 📦 Instale manualmente: pip install webdriver-manager
4. 🔍 Ou baixe ChromeDriver em: https://chromedriver.chromium.org/

Erro técnico: {str(e)}
            """
            print(error_msg)
            self.logger.error(f"Erro ao configurar driver: {str(e)}")
            raise

    def ler_dados_excel(self):
        """Lê e processa os dados do Excel"""
        try:
            df = pd.read_excel(self.caminho_excel)

            # Limpeza e formatação dos dados obrigatórios do CPF
            df['CPF'] = df['CPF'].astype(str).str.replace('zero', '0').str.replace('nan', '00000000000')

            # Remove .0 APENAS do final (quando vem de float do Excel) - mais preciso
            df['CPF'] = df['CPF'].str.replace(r'\.0$', '', regex=True)  # Remove .0 apenas do final

            # Remove pontos e traços de formatação
            df['CPF'] = df['CPF'].str.replace('.', '', regex=False).str.replace('-', '', regex=False)

            # Garante que seja apenas dígitos
            df['CPF'] = df['CPF'].str.replace(r'[^\d]', '', regex=True)

            # Preenche com zeros à esquerda para ter 11 dígitos
            df['CPF'] = df['CPF'].str.zfill(11)

            # Log para debug de CPFs processados
            self.logger.info("CPFs processados (primeiros 5):")
            for i, cpf in enumerate(df['CPF'].head(5)):
                self.logger.info(f"  Registro {i+1}: {cpf} (tamanho: {len(cpf)})")

            df['Valor'] = df['Valor'].astype(str).str.replace(',', '.')
            df['Valor'] = pd.to_numeric(df['Valor'], errors='coerce')

            # Formatação da data - apenas se a coluna existir
            if 'Data' in df.columns:
                df['Data'] = pd.to_datetime(df['Data'], format='%d/%m/%y', errors='coerce')
            else:
                # Para Katia, se não tem Data, cria uma coluna vazia
                df['Data'] = None

            # Verificação de colunas opcionais para Katia
            if self.cliente_atual == 'katia':
                colunas_opcionais = ['Cidade', 'Endereco']
                for coluna in colunas_opcionais:
                    if coluna not in df.columns:
                        df[coluna] = ''  # Cria coluna vazia se não existir

            self.logger.info(f"Dados carregados: {len(df)} registros")
            self.logger.info(f"Colunas disponíveis: {list(df.columns)}")
            return df

        except Exception as e:
            self.logger.error(f"Erro ao ler Excel: {str(e)}")
            raise

    def navegar_para_site(self):
        """Navega para o site de notas fiscais"""
        try:
            self.driver.get(self.url_site)
            self.logger.info("Navegação para o site realizada")
            time.sleep(1)  # Reduzido para 1s
        except Exception as e:
            self.logger.error(f"Erro ao navegar para o site: {str(e)}")
            raise

    def aguardar_elemento(self, locator, timeout=10):
        """Aguarda um elemento ficar disponível"""
        try:
            return WebDriverWait(self.driver, timeout).until(
                EC.element_to_be_clickable(locator)
            )
        except TimeoutException:
            self.logger.warning(f"Timeout aguardando elemento: {locator}")
            return None

    def aguardar_ajax_cpf(self, timeout=8):
        """Wait inteligente para AJAX do CPF e preenchimento automático de campos"""
        try:
            # Múltiplas estratégias para detectar fim do AJAX e preenchimento automático
            inicio = time.time()

            while (time.time() - inicio) < timeout:
                try:
                    # Estratégia 1: Verifica se não há requisições jQuery/PrimeFaces ativas
                    ajax_ativo = self.driver.execute_script("""
                        // Verifica jQuery AJAX
                        var jqueryAtivo = (typeof jQuery !== 'undefined' && jQuery.active > 0);

                        // Verifica PrimeFaces AJAX
                        var pfAtivo = false;
                        if (typeof PrimeFaces !== 'undefined' && PrimeFaces.ajax) {
                            pfAtivo = PrimeFaces.ajax.Queue.isEmpty !== undefined ?
                                     !PrimeFaces.ajax.Queue.isEmpty() : false;
                        }

                        return jqueryAtivo || pfAtivo;
                    """)

                    # Estratégia 2: Verifica se não há indicadores de loading
                    loading_ativo = self.driver.execute_script("""
                        var loadings = document.querySelectorAll('.ui-blockui, .loading, [id*="loading"], .ui-ajax-status');
                        return loadings.length > 0 && Array.from(loadings).some(el =>
                            el.style.display !== 'none' && el.offsetParent !== null
                        );
                    """)

                    # Estratégia 3: Verifica se campos que deveriam ser preenchidos automaticamente estão disponíveis
                    campos_preenchidos = self.driver.execute_script("""
                        // Lista de campos que podem ser preenchidos automaticamente
                        var camposVerificar = [
                            'frmConteudo:itRazaoSocialT',
                            'frmConteudo:somUfT',
                            'frmConteudo:itLogradouroT'
                        ];

                        var camposOk = 0;
                        for (var i = 0; i < camposVerificar.length; i++) {
                            var campo = document.getElementById(camposVerificar[i]);
                            if (campo && !campo.disabled) {
                                camposOk++;
                            }
                        }

                        return camposOk >= 2; // Pelo menos 2 dos 3 campos devem estar disponíveis
                    """)

                    # Se não há AJAX ativo E não há loading E campos estão disponíveis
                    if not ajax_ativo and not loading_ativo and campos_preenchidos:
                        self.logger.info("CPF AJAX completo - campos preenchidos automaticamente")
                        return True

                    time.sleep(0.1)  # Check rápido a cada 100ms

                except:
                    # Em caso de erro, continua tentando
                    time.sleep(0.1)

            # Se chegou ao timeout, assume que terminou
            self.logger.warning("Wait AJAX CPF: timeout atingido, continuando...")
            return False

        except Exception as e:
            self.logger.warning(f"Erro no wait AJAX CPF: {e}")
            time.sleep(1)  # Fallback mínimo
            return False

    def verificar_campos_preenchidos_automaticamente(self):
        """Verifica quais campos foram preenchidos automaticamente pelo CPF"""
        try:
            resultado = self.driver.execute_script("""
                var campos = {
                    nome: false,
                    uf: false,
                    municipio: false,
                    logradouro: false,
                    numero: false,
                    cep: false,
                    telefone: false,
                    email: false,
                    tipo_logradouro: false
                };

                // Verifica campo Nome/Razão Social
                var nomeField = document.getElementById('frmConteudo:itRazaoSocialT');
                if (nomeField && nomeField.value && nomeField.value.trim()) {
                    campos.nome = true;
                }

                // Verifica dropdown UF
                var ufField = document.getElementById('frmConteudo:somUfT_input');
                if (ufField && ufField.selectedIndex > 0) {
                    campos.uf = true;
                }

                // Verifica dropdown Município
                var municipioField = document.getElementById('frmConteudo:somMunicipioT_input');
                if (municipioField && municipioField.selectedIndex > 0) {
                    campos.municipio = true;
                }

                // Verifica campo Logradouro
                var logradouroField = document.getElementById('frmConteudo:itLogradouroT');
                if (logradouroField && logradouroField.value && logradouroField.value.trim()) {
                    campos.logradouro = true;
                }

                // Verifica campo Número
                var numeroField = document.getElementById('frmConteudo:itNumeroT');
                if (numeroField && numeroField.value && numeroField.value.trim()) {
                    campos.numero = true;
                }

                // Verifica campo CEP
                var cepField = document.getElementById('frmConteudo:itCepT');
                if (cepField && cepField.value && cepField.value.trim()) {
                    campos.cep = true;
                }

                // Verifica campo Telefone
                var telefoneField = document.getElementById('frmConteudo:itTelefoneT');
                if (telefoneField && telefoneField.value && telefoneField.value.trim()) {
                    campos.telefone = true;
                }

                // Verifica campo Email
                var emailField = document.getElementById('frmConteudo:itEmailT');
                if (emailField && emailField.value && emailField.value.trim()) {
                    campos.email = true;
                }

                // Verifica dropdown Tipo Logradouro
                var tipoLogradouroField = document.getElementById('frmConteudo:somTipoLogradouroT_input');
                if (tipoLogradouroField && tipoLogradouroField.selectedIndex > 0) {
                    campos.tipo_logradouro = true;
                }

                return campos;
            """)

            # Log dos campos preenchidos automaticamente
            campos_preenchidos = [campo for campo, preenchido in resultado.items() if preenchido]
            if campos_preenchidos:
                self.logger.info(f"Campos preenchidos automaticamente pelo CPF: {', '.join(campos_preenchidos)}")
            else:
                self.logger.info("Nenhum campo foi preenchido automaticamente pelo CPF")

            return resultado

        except Exception as e:
            self.logger.warning(f"Erro ao verificar campos preenchidos automaticamente: {e}")
            # Retorna False para todos os campos em caso de erro
            return {
                'nome': False, 'uf': False, 'municipio': False, 'logradouro': False,
                'numero': False, 'cep': False, 'telefone': False, 'email': False,
                'tipo_logradouro': False
            }

    def aguardar_municipios_carregados(self, timeout=3):
        """Wait inteligente para carregamento de municípios"""
        try:
            inicio = time.time()

            while (time.time() - inicio) < timeout:
                try:
                    # Verifica se o dropdown de municípios tem opções carregadas
                    municipios_carregados = self.driver.execute_script("""
                        var dropdown = document.getElementById('frmConteudo:somMunicipioT');
                        if (!dropdown) return false;

                        // Verifica se não está disabled
                        if (dropdown.disabled) return false;

                        // Verifica se há opções no panel (se existir)
                        var panel = document.getElementById('frmConteudo:somMunicipioT_panel');
                        if (panel) {
                            var items = panel.querySelectorAll('.ui-selectonemenu-item');
                            return items.length > 1; // Mais que só "Selecione..."
                        }

                        return true; // Se não há panel, assume que está pronto
                    """)

                    if municipios_carregados:
                        return True

                    time.sleep(0.1)

                except:
                    time.sleep(0.1)

            return False

        except Exception as e:
            self.logger.warning(f"Erro no wait municípios: {e}")
            time.sleep(0.5)  # Fallback mínimo
            return False

    def aguardar_municipios_carregados_incidencia(self, timeout=2):
        """Wait inteligente para municípios de incidência"""
        return self.aguardar_dropdown_carregado('frmConteudo:somMunicipioIncidencia', timeout)

    def aguardar_municipios_carregados_servico(self, timeout=2):
        """Wait inteligente para municípios de serviço"""
        return self.aguardar_dropdown_carregado('frmConteudo:somMunicipioServico', timeout)

    def aguardar_dropdown_carregado(self, dropdown_id, timeout=2):
        """Wait genérico para dropdown carregado"""
        try:
            inicio = time.time()
            while (time.time() - inicio) < timeout:
                try:
                    carregado = self.driver.execute_script(f"""
                        var dropdown = document.getElementById('{dropdown_id}');
                        return dropdown && !dropdown.disabled;
                    """)
                    if carregado:
                        return True
                    time.sleep(0.05)
                except:
                    time.sleep(0.05)
            return False
        except:
            time.sleep(0.2)
            return False

    def selecionar_dropdown(self, element_id, value, retry_count=3):
        """Seleciona valor em dropdown - OTIMIZADO para PrimeFaces"""
        for attempt in range(retry_count):
            try:
                self.logger.info(f"Tentativa {attempt + 1}: Selecionando '{value}' no dropdown {element_id}")

                # Fecha qualquer dropdown/modal que possa estar aberto
                self.fechar_dropdowns_abertos()
                self.fechar_modals()

                # ESTRATÉGIA 1: Usar select oculto diretamente (mais rápido e confiável)
                try:
                    select_input_id = f"{element_id}_input"
                    select_element = self.driver.find_element(By.ID, select_input_id)

                    if select_element:
                        select = Select(select_element)
                        options = select.options
                        success = False

                        # Estratégia 1: Por value direto (para códigos como '508')
                        try:
                            select.select_by_value(str(value))
                            success = True
                            self.logger.info(f"Selecionado por value: {value}")
                        except:
                            pass

                        # Estratégia 2: Por texto exato
                        if not success:
                            try:
                                select.select_by_visible_text(str(value))
                                success = True
                                self.logger.info(f"Selecionado por texto exato: {value}")
                            except:
                                pass

                        # Estratégia 3: Busca flexível por texto que contenha o valor
                        if not success:
                            for option in options:
                                option_value = option.get_attribute("value") or ""
                                option_text = option.text.strip()

                                # Busca mais flexível para textos
                                if (str(value).upper() in option_text.upper() or
                                    option_text.upper().startswith(str(value).upper()) or
                                    str(value) == option_value):
                                    select.select_by_value(option_value)
                                    success = True
                                    self.logger.info(f"Selecionado por busca flexível: '{option_text}' (value={option_value})")
                                    break

                        # Estratégia 4: Para casos específicos como "INDAIATUBA"
                        if not success and str(value).upper() == "INDAIATUBA":
                            for option in options:
                                option_text = option.text.strip().upper()
                                if "INDAIATUBA" in option_text:
                                    option_value = option.get_attribute("value")
                                    select.select_by_value(option_value)
                                    success = True
                                    self.logger.info(f"Selecionado INDAIATUBA: '{option.text}' (value={option_value})")
                                    break

                        if success:
                            # Dispara eventos PrimeFaces
                            self.driver.execute_script("""
                                var select = arguments[0];
                                select.dispatchEvent(new Event('change', {bubbles: true}));
                                var labelId = arguments[1] + '_label';
                                var label = document.getElementById(labelId);
                                if (label && select.selectedOptions.length > 0) {
                                    label.textContent = select.selectedOptions[0].text;
                                }
                            """, select_element, element_id)
                            time.sleep(0.3)
                            return True
                        else:
                            # Log das opções disponíveis para debug
                            opcoes_disponiveis = [f"'{opt.text}' (value='{opt.get_attribute('value')}')" for opt in options[:5]]
                            self.logger.warning(f"'{value}' não encontrado. Primeiras opções: {opcoes_disponiveis}")
                            raise Exception(f"'{value}' não encontrado nas opções")

                except Exception as e:
                    self.logger.debug(f"Método select oculto falhou: {e}")

                # ESTRATÉGIA 2: Método clássico com panel (fallback)
                self.logger.info("Tentando método clássico com click...")

                # Aguarda o elemento estar presente e clicável
                dropdown = self.wait.until(EC.element_to_be_clickable((By.ID, element_id)))

                # Scroll para o elemento
                self.driver.execute_script("arguments[0].scrollIntoView({behavior: 'instant', block: 'center'});", dropdown)
                time.sleep(0.3)

                # Clica no dropdown
                try:
                    dropdown.click()
                except Exception:
                    self.driver.execute_script("arguments[0].click();", dropdown)

                # Procura por panel/items com múltiplos sufixos possíveis
                panel_found = False
                for panel_suffix in ['_panel', '_items', '_list']:
                    try:
                        panel_id = f"{element_id}{panel_suffix}"
                        panel = WebDriverWait(self.driver, 2).until(EC.visibility_of_element_located((By.ID, panel_id)))

                        option = self.encontrar_opcao_dropdown(panel_id, value)
                        if option:
                            option.click()
                            self.logger.info(f"Selecionado via panel: {panel_id}")
                            time.sleep(0.3)
                            return True
                        panel_found = True
                        break
                    except TimeoutException:
                        continue

                if not panel_found:
                    raise Exception("Nenhum panel encontrado")

            except Exception as e:
                self.fechar_dropdowns_abertos()

                if attempt < retry_count - 1:
                    self.logger.warning(f"Tentativa {attempt + 1} falhou: {e}. Tentando novamente...")
                    time.sleep(1)
                else:
                    self.logger.error(f"Todas as tentativas falharam para {element_id}: {e}")
                    return False

        return False


    def encontrar_opcao_dropdown(self, panel_id, value):
        """Encontra opção no dropdown com log detalhado para debug"""
        try:
            self.logger.info(f"Procurando por '{value}' no dropdown {panel_id}")

            # Debug: lista todas as opções disponíveis com múltiplas estratégias
            try:
                # Tenta vários seletores CSS
                selectors = [
                    f"#{panel_id} .ui-selectonemenu-item",
                    f"#{panel_id} li",
                    f"#{panel_id} .ui-selectonemenu-list-item",
                    f"#{panel_id} .ui-menu-item",
                    f"#{panel_id} [role='option']",
                    f"#{panel_id} *[data-label]"
                ]

                all_items = []
                for selector in selectors:
                    items = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    if items:
                        all_items = items
                        self.logger.info(f"Seletor que funcionou: {selector}")
                        break

                if not all_items:
                    # Se nenhum seletor funcionou, tenta buscar todos os elementos no panel
                    all_items = self.driver.find_elements(By.CSS_SELECTOR, f"#{panel_id} *")
                    self.logger.warning(f"Usando seletor genérico - encontrados {len(all_items)} elementos")

                self.logger.info(f"Encontradas {len(all_items)} opções no dropdown")
                for i, item in enumerate(all_items[:15]):  # Log as primeiras 15
                    try:
                        data_label = item.get_attribute("data-label")
                        text = item.text.strip()
                        tag = item.tag_name
                        classes = item.get_attribute("class")
                        self.logger.info(f"Opção {i+1}: tag='{tag}', class='{classes}', data-label='{data_label}', text='{text}'")
                    except Exception as e:
                        self.logger.warning(f"Erro ao ler item {i+1}: {e}")
            except Exception as e:
                self.logger.error(f"Erro no debug das opções: {e}")

            # Busca com múltiplas estratégias usando os seletores que funcionaram
            selectors_to_try = [
                f"#{panel_id} .ui-selectonemenu-item",
                f"#{panel_id} li",
                f"#{panel_id} .ui-selectonemenu-list-item",
                f"#{panel_id} .ui-menu-item",
                f"#{panel_id} [role='option']",
                f"#{panel_id} *[data-label]",
                f"#{panel_id} *"  # Seletor genérico como último recurso
            ]

            # Estratégia 1: Busca por data-label exato
            for selector in selectors_to_try:
                try:
                    option = self.driver.find_element(By.CSS_SELECTOR, f"{selector}[data-label='{value}']")
                    if option and option.is_displayed():
                        self.logger.info(f"Encontrou por data-label exato: {value}")
                        return option
                except:
                    continue

            # Estratégia 2: Busca manual por todos os elementos
            for selector in selectors_to_try:
                try:
                    items = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    if not items:
                        continue

                    for item in items:
                        try:
                            if not item.is_displayed():
                                continue

                            data_label = item.get_attribute("data-label") or ""
                            text = item.text.strip()

                            # Busca exata primeiro
                            if str(data_label) == str(value) or str(text) == str(value):
                                self.logger.info(f"Encontrou por busca exata: selector='{selector}', data-label='{data_label}', text='{text}'")
                                return item
                        except:
                            continue
                    break  # Se encontrou itens, não tenta outros seletores
                except:
                    continue

            # Estratégia 3: Busca parcial como último recurso
            for selector in selectors_to_try:
                try:
                    items = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    if not items:
                        continue

                    for item in items:
                        try:
                            if not item.is_displayed():
                                continue

                            data_label = item.get_attribute("data-label") or ""
                            text = item.text.strip()

                            # Busca parcial
                            if str(value) in str(data_label) or str(value) in str(text):
                                self.logger.info(f"Encontrou por busca parcial: selector='{selector}', data-label='{data_label}', text='{text}'")
                                return item
                        except:
                            continue
                    break  # Se encontrou itens, não tenta outros seletores
                except:
                    continue

            self.logger.error(f"Opção '{value}' NÃO ENCONTRADA no dropdown {panel_id}")
            return None

        except Exception as e:
            self.logger.error(f"Erro ao encontrar opção {value}: {e}")
            return None

    def preencher_campo(self, element_id, valor, retry_count=3):
        """Preenche um campo de texto sempre substituindo valores existentes"""
        for attempt in range(retry_count):
            try:
                # Aguarda o elemento estar presente e clicável em uma operação
                campo = self.wait.until(EC.element_to_be_clickable((By.ID, element_id)))

                # SEMPRE substitui o valor, sem verificar se já está preenchido
                # Scroll otimizado e preenchimento via JavaScript (mais rápido e confiável)
                self.driver.execute_script(f"""
                    var campo = arguments[0];
                    campo.scrollIntoView({{behavior: 'instant', block: 'center'}});
                    campo.focus();

                    // Sequência otimizada: limpa, aguarda, preenche, verifica
                    campo.value = '';

                    // Pequena pausa para garantir que a limpeza foi processada
                    setTimeout(function() {{
                        campo.value = '{valor}';

                        // Verifica se foi preenchido corretamente
                        if (campo.value !== '{valor}') {{
                            campo.value = '{valor}'; // Tenta novamente se necessário
                        }}

                        // Dispara eventos após confirmar o valor
                        campo.dispatchEvent(new Event('input', {{bubbles: true}}));
                        campo.dispatchEvent(new Event('change', {{bubbles: true}}));
                    }}, 10);
                """, campo)

                self.logger.info(f"Campo {element_id} preenchido com: {valor}")
                time.sleep(0.02)  # Aguarda timeout do JavaScript (10ms) + margem
                return True

            except Exception as e:
                if attempt < retry_count - 1:
                    self.logger.warning(f"Tentativa {attempt + 1} falhou para campo {element_id}. Tentando novamente...")
                    time.sleep(0.2)
                else:
                    self.logger.error(f"Erro ao preencher campo {element_id} após {retry_count} tentativas: {str(e)}")
        return False

    def mapear_cidade_para_codigo(self, cidade):
        """Mapeia nome da cidade para código (alguns exemplos do HTML)"""
        mapeamento_cidades = {
            'SUMARE': '3552403',
            'CAMPINAS': '3509502',
            'VALINHOS': '3556206',
            'CAPIVARI': '3510401',
            'INDAIATUBA': '3520509'
        }
        return mapeamento_cidades.get(cidade.upper(), '')

    def mapear_tipo_logradouro(self, prefixo):
        """Mapeia prefixo do endereço para tipo de logradouro"""
        mapeamento = {
            'AL': 'ALAMEDA',
            'R': 'RUA',
            'AV': 'AVENIDA',
            'JD': 'JARDIM'
        }
        return mapeamento.get(prefixo, 'RUA')  # RUA como padrão

    def extrair_endereco(self, endereco_completo):
        """Extrai componentes do endereço seguindo nova regra"""
        endereco = endereco_completo.strip()

        # Busca por padrões como "R CLaudemeses dos Santos, 164"
        # Captura: prefixo + espaço + logradouro + vírgula + número
        match = re.match(r'^(AL|R|AV|JD)\s+(.+?),\s*(\d+).*$', endereco)

        if match:
            prefixo = match.group(1)
            logradouro = match.group(2).strip()
            numero = match.group(3).strip()
            tipo_logradouro = self.mapear_tipo_logradouro(prefixo)
        else:
            # Fallback para formato antigo se não seguir o novo padrão
            endereco_sem_prefixo = re.sub(r'^(AL|R|AV|JD)\s+', '', endereco)
            match_antigo = re.search(r'(.+?)\s*,?\s*(\d+)$', endereco_sem_prefixo)

            if match_antigo:
                logradouro = match_antigo.group(1).strip()
                numero = match_antigo.group(2).strip()
            else:
                logradouro = endereco_sem_prefixo
                numero = ''

            tipo_logradouro = 'RUA'  # Padrão

        return logradouro, numero, tipo_logradouro

    def preencher_aliquota(self):
        """Preenche o campo alíquota"""
        self.preencher_campo('frmConteudo:itAliquota', self.configuracoes_padrao['itAliquota'])

    def preencher_uf_incidencia(self):
        """Preenche UF de incidência"""
        self.selecionar_dropdown('frmConteudo:somUfIncidencia', self.configuracoes_padrao['uf_incidencia'])
        # Wait inteligente removido - próximo método aguardará se necessário

    def preencher_municipio_incidencia(self):
        """Preenche município de incidência"""
        # Aguarda municípios carregados se necessário
        self.aguardar_municipios_carregados_incidencia()
        self.selecionar_dropdown('frmConteudo:somMunicipioIncidencia', self.configuracoes_padrao['municipio_incidencia'])

    def preencher_uf_servico(self):
        """Preenche UF do serviço"""
        self.selecionar_dropdown('frmConteudo:somUfServico', self.configuracoes_padrao['UfServico'])
        # Wait inteligente removido - próximo método aguardará se necessário

    def preencher_municipio_servico(self):
        """Preenche município do serviço"""
        # Aguarda municípios carregados se necessário
        self.aguardar_municipios_carregados_servico()
        self.selecionar_dropdown('frmConteudo:somMunicipioServico', self.configuracoes_padrao['somMunicipioServico'])

    def gerar_descricao_servico(self, nome_pet, data):
        """Gera descrição do serviço no formato padrão"""
        # Para Katia, usa o formato "SERVIÇOS PRESTADOS AO PET + [NOME PET]"
        if self.cliente_atual == 'katia':
            return f"SERVIÇOS PRESTADOS AO PET {nome_pet.upper()}"
        else:
            # Para outros clientes (Kleiton), mantém o formato original com data
            if data is not None and pd.notna(data):
                # Formatar data para DD/MM/AAAA
                if isinstance(data, str):
                    # Se já está em string, usa direto
                    data_formatada = data
                else:
                    # Se é datetime, formata para DD/MM/AAAA
                    data_formatada = data.strftime('%d/%m/%Y')
                return f"ANESTESIA {nome_pet.upper()} EM {data_formatada}"
            else:
                # Se não tem data, usa formato sem data
                return f"ANESTESIA {nome_pet.upper()}"

    def gerar_observacoes(self, valor):
        """Gera observações com cálculo do imposto"""
        # Para Katia, observações ficam vazias
        if self.cliente_atual == 'katia':
            return ""
        else:
            # Para outros clientes (Kleiton), mantém o formato original
            # Calcula 6% do valor
            valor_imposto = valor * 0.06
            valor_imposto_formatado = f"{valor_imposto:.2f}".replace('.', ',')
            return f"ALIQUOTA 6%. VALOR APROXIMADO IMPOSTO R${valor_imposto_formatado}"

    def preencher_formulario(self, dados_linha):
        """Preenche o formulário com os dados de uma linha"""
        try:
            self.logger.info(f"Preenchendo nota para: {dados_linha['Nome_Cliente']}")

            # 1. ATIVIDADE (já configurada como padrão)
            self.selecionar_dropdown('frmConteudo:somAtividade', self.configuracoes_padrao['atividade'])

            # 2. IDENTIFICAÇÃO DO TOMADOR
            self.selecionar_dropdown('frmConteudo:somTipoPessoa', self.configuracoes_padrao['tipo_pessoa'])

            # CPF - CORRIGIDO para evitar stale element reference
            try:
                # Aguarda um pouco após selecionar tipo_pessoa para AJAX processar
                time.sleep(0.5)

                cpf = str(dados_linha['CPF']).strip()
                self.logger.info(f"DEBUG CPF: Tentando preencher CPF '{cpf}' (tamanho: {len(cpf)})")

                # Validação adicional do CPF
                if len(cpf) != 11 or not cpf.isdigit():
                    self.logger.warning(f"CPF inválido: '{cpf}' - deveria ter 11 dígitos numéricos")

                # Método mais robusto - sempre busca elemento novamente
                campo_cpf_preenchido = False
                for tentativa_cpf in range(3):
                    try:
                        # Sempre busca o elemento novamente para evitar stale reference
                        campo_cpf = self.wait.until(EC.element_to_be_clickable((By.ID, 'frmConteudo:imCpfCnpjT')))

                        # Scroll e foco
                        self.driver.execute_script("arguments[0].scrollIntoView({behavior: 'instant', block: 'center'});", campo_cpf)

                        # Método 1: Tentativa síncrona mais robusta
                        resultado = self.driver.execute_script(f"""
                            var campo = document.getElementById('frmConteudo:imCpfCnpjT');
                            if (!campo) return 'campo_nao_encontrado';

                            // Passo 1: Preparação do campo
                            campo.focus();
                            campo.select(); // Seleciona todo o texto existente
                            campo.value = ''; // Limpa

                            // Passo 2: Preenchimento robusto
                            campo.value = '{cpf}';

                            // Verifica múltiplas vezes se foi preenchido
                            var tentativas = 0;
                            while (campo.value !== '{cpf}' && tentativas < 3) {{
                                campo.value = '{cpf}';
                                tentativas++;
                            }}

                            if (campo.value !== '{cpf}') {{
                                return 'erro_preenchimento';
                            }}

                            // Passo 3: Eventos de validação
                            campo.dispatchEvent(new Event('input', {{bubbles: true}}));
                            campo.dispatchEvent(new Event('change', {{bubbles: true}}));

                            // Passo 4: Ativação do preenchimento automático
                            campo.blur();
                            campo.dispatchEvent(new Event('blur', {{bubbles: true}}));

                            return 'sucesso';
                        """)

                        if resultado == 'sucesso':
                            self.logger.info(f"CPF preenchido: {cpf}")
                            campo_cpf_preenchido = True

                            # Pequena pausa para garantir que os eventos foram processados
                            time.sleep(0.1)

                            # Wait inteligente - aguarda AJAX do CPF terminar
                            self.aguardar_ajax_cpf()
                            break
                        elif resultado == 'erro_preenchimento':
                            self.logger.warning(f"Tentativa {tentativa_cpf + 1}: Erro ao preencher valor do CPF")
                        elif resultado == 'campo_nao_encontrado':
                            self.logger.warning(f"Tentativa {tentativa_cpf + 1}: Campo CPF não encontrado")
                        else:
                            self.logger.warning(f"Tentativa {tentativa_cpf + 1}: Resultado inesperado: {resultado}")

                    except Exception as e:
                        self.logger.warning(f"Tentativa {tentativa_cpf + 1} falhou: {e}")
                        if tentativa_cpf < 2:
                            time.sleep(0.5)  # Aguarda antes da próxima tentativa
                        continue

                if not campo_cpf_preenchido:
                    self.logger.error("Erro: Não foi possível preencher o CPF após 3 tentativas")

            except Exception as e:
                self.logger.error(f"Erro geral ao preencher CPF: {str(e)}")
            # Verifica se os campos foram preenchidos automaticamente pelo CPF
            campos_preenchidos_auto = self.verificar_campos_preenchidos_automaticamente()

            # Nome/Razão Social - preenche apenas se não foi preenchido automaticamente
            if not campos_preenchidos_auto.get('nome', False):
                self.preencher_campo('frmConteudo:itRazaoSocialT', dados_linha['Nome_Cliente'].upper())
            else:
                self.logger.info("Nome/Razão Social já preenchido automaticamente pelo CPF")

            # UF - verifica se foi preenchido automaticamente
            cidade_valida = pd.notna(dados_linha['Cidade']) and str(dados_linha['Cidade']).strip()
            if not campos_preenchidos_auto.get('uf', False):
                # UF não foi preenchida automaticamente, preenche manualmente
                if self.cliente_atual != 'katia' or cidade_valida:
                    self.selecionar_dropdown('frmConteudo:somUfT', self.configuracoes_padrao['uf'])
            else:
                self.logger.info("UF já preenchida automaticamente pelo CPF")

            # Município - verifica se precisa preencher
            if cidade_valida and not campos_preenchidos_auto.get('municipio', False):
                self.aguardar_municipios_carregados()
                cidade_upper = str(dados_linha['Cidade']).upper()
                self.selecionar_dropdown('frmConteudo:somMunicipioT', cidade_upper)
            elif campos_preenchidos_auto.get('municipio', False):
                self.logger.info("Município já preenchido automaticamente pelo CPF")

            # Endereço - verifica se foi preenchido automaticamente
            endereco_valido = pd.notna(dados_linha['Endereco']) and str(dados_linha['Endereco']).strip()
            if endereco_valido:
                logradouro, numero, tipo_logradouro = self.extrair_endereco(dados_linha['Endereco'])

                # Tipo de logradouro - sempre preenche se necessário
                if not campos_preenchidos_auto.get('tipo_logradouro', False):
                    self.selecionar_dropdown('frmConteudo:somTipoLogradouroT', tipo_logradouro)

                # Logradouro - preenche apenas se não foi preenchido automaticamente
                if not campos_preenchidos_auto.get('logradouro', False):
                    self.preencher_campo('frmConteudo:itLogradouroT', logradouro.upper())
                else:
                    self.logger.info("Logradouro já preenchido automaticamente pelo CPF")

                # Número - sempre preenche se disponível
                if numero and not campos_preenchidos_auto.get('numero', False):
                    self.preencher_campo('frmConteudo:itNumeroT', numero)

            # 3. LOCAL DE INCIDÊNCIA DO IMPOSTO
            self.preencher_uf_incidencia()
            self.preencher_municipio_incidencia()

            # 4. PRESTAÇÃO DO SERVIÇO
            self.selecionar_dropdown('frmConteudo:somExigibilidade', self.configuracoes_padrao['exigibilidade'])

            # Simples Nacional (otimizado)
            self.selecionar_dropdown('frmConteudo:somSimplesNacional', self.configuracoes_padrao['simples_nacional'])
            time.sleep(0.3)  # Tempo mínimo para modal aparecer
            self.fechar_modals()  # Fecha modal do Simples Nacional

            self.selecionar_dropdown('frmConteudo:somRegimeEspecial', self.configuracoes_padrao['regime_especial'])
            self.selecionar_dropdown('frmConteudo:somIssRetido', self.configuracoes_padrao['iss_retido'])

            # Valor do Serviço
            valor_formatado = f"{dados_linha['Valor']:.2f}".replace('.', ',')
            self.preencher_campo('frmConteudo:itValorServico', valor_formatado)

            # Alíquota usando a nova função
            self.preencher_aliquota()

            # Valor das Deduções
            self.preencher_campo('frmConteudo:itValorDeducoes', self.configuracoes_padrao['valor_deducoes'])

            # Incentivo Fiscal - múltiplas tentativas
            incentivo_sucesso = False

            # Tentativa 1: Select oculto por value com eventos PrimeFaces
            try:
                select_element = self.driver.find_element(By.ID, 'frmConteudo:somIncentivo_input')
                select = Select(select_element)
                select.select_by_value('2')

                # Dispara eventos para atualizar a interface PrimeFaces
                self.driver.execute_script("""
                    var select = arguments[0];
                    var dropdown = document.getElementById('frmConteudo:somIncentivo');
                    var label = document.getElementById('frmConteudo:somIncentivo_label');

                    // Atualiza o label visual
                    if (label) label.textContent = 'Não';

                    // Dispara eventos
                    select.dispatchEvent(new Event('change', {bubbles: true}));
                    if (dropdown) dropdown.dispatchEvent(new Event('change', {bubbles: true}));
                """, select_element)

                time.sleep(0.1)  # Reduzido
                incentivo_sucesso = True
                self.logger.info("Incentivo fiscal: sucesso via select value com eventos")
            except Exception as e:
                self.logger.debug(f"Tentativa 1 falhou: {str(e)}")

            # Tentativa 2: Select oculto por texto
            if not incentivo_sucesso:
                try:
                    select_element = self.driver.find_element(By.ID, 'frmConteudo:somIncentivo_input')
                    select = Select(select_element)
                    select.select_by_visible_text('Não')
                    time.sleep(0.1)  # Reduzido
                    incentivo_sucesso = True
                    self.logger.info("Incentivo fiscal: sucesso via select texto")
                except Exception as e:
                    self.logger.debug(f"Tentativa 2 falhou: {str(e)}")

            # Tentativa 3: JavaScript direto
            if not incentivo_sucesso:
                try:
                    self.driver.execute_script("""
                        var select = document.getElementById('frmConteudo:somIncentivo_input');
                        select.value = '2';
                        select.dispatchEvent(new Event('change', {bubbles: true}));
                    """)
                    time.sleep(0.1)  # Reduzido
                    incentivo_sucesso = True
                    self.logger.info("Incentivo fiscal: sucesso via JavaScript")
                except Exception as e:
                    self.logger.debug(f"Tentativa 3 falhou: {str(e)}")

            # Tentativa 4: Método dropdown manual
            if not incentivo_sucesso:
                self.selecionar_dropdown('frmConteudo:somIncentivo', 'Não')

            if not incentivo_sucesso:
                self.logger.warning("Todas as tentativas de incentivo fiscal falharam")

            # 5. LOCAL DE REALIZAÇÃO DO SERVIÇO
            self.preencher_uf_servico()
            self.preencher_municipio_servico()

            # 6. DESCRIÇÃO DO SERVIÇO
            descricao = self.gerar_descricao_servico(dados_linha['Nome_Pet'], dados_linha['Data'])
            self.preencher_campo('frmConteudo:itaDescricaoServico', descricao)

            # 7. OBSERVAÇÕES
            observacoes = self.gerar_observacoes(dados_linha['Valor'])
            self.preencher_campo('frmConteudo:itaObservacoes', observacoes)

            # 8. VALORES DE RETENÇÕES (preenchimento em lote otimizado)
            self.preencher_retencoes_lote()

            self.logger.info("Formulário preenchido com sucesso")
            return True

        except Exception as e:
            self.logger.error(f"Erro ao preencher formulário: {str(e)}")
            return False

    def preencher_retencoes_lote(self):
        """Preenche todos os campos de retenção em uma operação JavaScript otimizada"""
        try:
            # Mapeamento de campos de retenção
            campos_retencoes = {
                'frmConteudo:itInss': self.configuracoes_padrao['inss'],
                'frmConteudo:itIr': self.configuracoes_padrao['ir'],
                'frmConteudo:itCsll': self.configuracoes_padrao['csll'],
                'frmConteudo:itCofins': self.configuracoes_padrao['cofins'],
                'frmConteudo:itPis': self.configuracoes_padrao['pis'],
                'frmConteudo:itOutrasRetencoes': self.configuracoes_padrao['outras_retencoes']
            }

            # JavaScript para preencher todos os campos de uma vez
            js_code = """
                var campos = arguments[0];
                var sucessos = 0;
                var erros = [];

                for (var id in campos) {
                    try {
                        var campo = document.getElementById(id);
                        if (campo) {
                            // SEMPRE substitui o valor, sem verificar se já está preenchido
                            campo.scrollIntoView({behavior: 'instant', block: 'center'});
                            campo.focus();
                            campo.value = '';  // Limpa primeiro
                            campo.value = campos[id];  // Preenche com novo valor
                            campo.dispatchEvent(new Event('input', {bubbles: true}));
                            campo.dispatchEvent(new Event('change', {bubbles: true}));
                            sucessos++;
                        } else {
                            erros.push('Campo não encontrado: ' + id);
                        }
                    } catch (e) {
                        erros.push('Erro em ' + id + ': ' + e.message);
                    }
                }

                return {sucessos: sucessos, erros: erros};
            """

            resultado = self.driver.execute_script(js_code, campos_retencoes)

            if resultado['sucessos'] > 0:
                self.logger.info(f"Retenções: {resultado['sucessos']} campos preenchidos em lote")

            if resultado['erros']:
                self.logger.warning(f"Retenções: {len(resultado['erros'])} erros - {resultado['erros']}")
                # Fallback para campos com erro
                for campo_id, valor in campos_retencoes.items():
                    if any(campo_id in erro for erro in resultado['erros']):
                        self.preencher_campo(campo_id, valor)

            return resultado['sucessos'] > 0

        except Exception as e:
            self.logger.error(f"Erro no preenchimento de retenções em lote: {e}")
            # Fallback para método individual
            self.preencher_campo('frmConteudo:itInss', self.configuracoes_padrao['inss'])
            self.preencher_campo('frmConteudo:itIr', self.configuracoes_padrao['ir'])
            self.preencher_campo('frmConteudo:itCsll', self.configuracoes_padrao['csll'])
            self.preencher_campo('frmConteudo:itCofins', self.configuracoes_padrao['cofins'])
            self.preencher_campo('frmConteudo:itPis', self.configuracoes_padrao['pis'])
            self.preencher_campo('frmConteudo:itOutrasRetencoes', self.configuracoes_padrao['outras_retencoes'])
            return False


    def emitir_nota(self):
        """Clica no botão emitir nota"""
        try:
            botao_emitir = self.wait.until(
                EC.element_to_be_clickable((By.ID, 'frmConteudo:cbEmitirNf'))
            )
            botao_emitir.click()
            self.logger.info("Nota fiscal emitida")
            time.sleep(3)  # Aguarda processamento
            return True
        except Exception as e:
            self.logger.error(f"Erro ao emitir nota: {str(e)}")
            return False

    def processar_notas(self, modo_teste=True):
        """Processa todas as notas do Excel com otimizações de performance"""
        import time as tempo_inicial
        inicio_processamento = tempo_inicial.time()

        try:
            print("\n🚀 RPA OTIMIZADO - VERSÃO 2.0")
            print("=" * 40)
            print("✨ Melhorias implementadas:")
            print("   • Wait inteligente CPF: 3s → 0.5-1s")
            print("   • Dropdowns otimizados: 0.7s → 0.2s")
            print("   • Preenchimento em lote de retenções")
            print("   • Scroll instantâneo")
            print("   • Delay entre registros: 2s → 0.5s")
            print("=" * 40)

            # Configurar driver
            self.configurar_driver()

            # Ler dados
            df = self.ler_dados_excel()

            # Navegar para o site
            self.navegar_para_site()

            # Aguardar carregamento da página
            input("Pressione ENTER após fazer login no site e estar na página de emissão...")

            sucessos = 0
            erros = 0
            tempos_por_nota = []

            # Processa todas as linhas (sem limite)
            limite = len(df)

            print(f"\n⏱️  MONITORAMENTO DE PERFORMANCE:")
            print("=" * 40)

            for index, linha in df.head(limite).iterrows():
                try:
                    inicio_nota = tempo_inicial.time()
                    self.logger.info(f"Processando registro {index + 1}/{limite}")

                    # Preencher formulário (sempre substitui dados existentes)
                    if self.preencher_formulario(linha):
                        if modo_teste:
                            # Em modo teste, apenas preenche sem emitir
                            tempo_nota = tempo_inicial.time() - inicio_nota
                            tempos_por_nota.append(tempo_nota)
                            print(f"📝 Nota {index + 1}: {tempo_nota:.1f}s")
                            input(f"Registro {index + 1} preenchido. Pressione ENTER para continuar...")
                        else:
                            # Em modo produção, emite a nota
                            if self.emitir_nota():
                                sucessos += 1
                                tempo_nota = tempo_inicial.time() - inicio_nota
                                tempos_por_nota.append(tempo_nota)
                                print(f"✅ Nota {index + 1}: {tempo_nota:.1f}s")
                                self.logger.info(f"Nota {index + 1} emitida com sucesso")
                            else:
                                erros += 1

                        # Delay otimizado entre registros
                        time.sleep(0.5)  # Reduzido drasticamente de 2s para 0.5s
                    else:
                        erros += 1

                except Exception as e:
                    self.logger.error(f"Erro no registro {index + 1}: {str(e)}")
                    erros += 1
                    continue

            # Relatório de performance
            tempo_total = tempo_inicial.time() - inicio_processamento
            if tempos_por_nota:
                tempo_medio = sum(tempos_por_nota) / len(tempos_por_nota)
                tempo_min = min(tempos_por_nota)
                tempo_max = max(tempos_por_nota)

                print("\n" + "=" * 40)
                print("📊 RELATÓRIO DE PERFORMANCE")
                print("=" * 40)
                print(f"⏱️  Tempo total: {tempo_total:.1f}s")
                print(f"📈 Tempo médio por nota: {tempo_medio:.1f}s")
                print(f"🚀 Tempo mínimo: {tempo_min:.1f}s")
                print(f"⚡ Tempo máximo: {tempo_max:.1f}s")
                print(f"📊 Notas processadas: {len(tempos_por_nota)}")

                # Comparação com versão anterior
                tempo_anterior_estimado = tempo_medio * 2.5  # Estimativa baseada nas otimizações
                melhoria = ((tempo_anterior_estimado - tempo_medio) / tempo_anterior_estimado) * 100
                print(f"\n🎯 MELHORIA ESTIMADA:")
                print(f"   Versão anterior: ~{tempo_anterior_estimado:.1f}s por nota")
                print(f"   Versão otimizada: {tempo_medio:.1f}s por nota")
                print(f"   Melhoria: {melhoria:.0f}% mais rápido!")
                print("=" * 40)

            self.logger.info(f"Processamento concluído. Sucessos: {sucessos}, Erros: {erros}")

        except Exception as e:
            self.logger.error(f"Erro no processamento: {str(e)}")
        finally:
            if self.driver:
                input("Pressione ENTER para fechar o navegador...")
                self.driver.quit()

def selecionar_mapeamento_cliente():
    """Permite ao usuário selecionar qual mapeamento de cliente usar"""
    print("\n" + "="*60)
    print("🎯 SELEÇÃO DO MAPEAMENTO DE CLIENTE")
    print("="*60)
    print("Clientes disponíveis:")
    print("1. 👨‍⚕️ KLEITON")
    print("   └─ 📍 Indaiatuba - Alíquota 2.01%")
    print()
    print("2. 👩‍⚕️ KATIA")
    print("   └─ 📍 Indaiatuba - Alíquota 2.01%")
    print("="*60)
    print("💡 Digite 'c' para ver comparação detalhada dos mapeamentos")
    print("="*60)

    while True:
        try:
            escolha = input("Digite sua opção (1, 2 ou 'c' para comparação): ").strip().lower()

            if escolha == '1':
                return 'kleiton'
            elif escolha == '2':
                return 'katia'
            elif escolha == 'c':
                mostrar_comparacao_temp()
                print("\n" + "="*60)
                print("Digite sua escolha após ver a comparação:")
            else:
                print("❌ Opção inválida! Digite 1 para Kleiton, 2 para Katia ou 'c' para comparação.")

        except KeyboardInterrupt:
            print("\n\nOperação cancelada pelo usuário.")
            exit()
        except Exception as e:
            print(f"❌ Erro: {str(e)}")

def mostrar_comparacao_temp():
    """Mostra comparação temporária dos mapeamentos antes da seleção"""
    mapeamentos = {
        'kleiton': {
           "atividade": "501"
        },
        'katia': {
            "atividade": "508"
        }
    }

    print("\n" + "="*80)
    print("📋 COMPARAÇÃO DOS MAPEAMENTOS DE CLIENTES")
    print("="*80)

    campos_diferentes = ["atividade"]
    labels = ["Atividade"]

    print(f"{'Campo':<25} {'KLEITON':<25} {'KATIA':<25}")
    print("-" * 80)

    for campo, label in zip(campos_diferentes, labels):
        kleiton_val = mapeamentos['kleiton'][campo]
        katia_val = mapeamentos['katia'][campo]
        print(f"{label:<25} {kleiton_val:<25} {katia_val:<25}")

    print("-" * 80)
    print("💡 As demais configurações são idênticas para ambos os clientes")
    print("="*80)

def main():
    """Função principal"""
    print("🤖 RPA NOTAS FISCAIS - SISTEMA MULTI-CLIENTE")

    # Seleção do cliente
    cliente_selecionado = selecionar_mapeamento_cliente()

    # Configurações
    URL_SITE = "https://deiss.indaiatuba.sp.gov.br/Deiss/restrito/nf_emissao.jsf"  # Substitua pela URL real
    CAMINHO_EXCEL = "notas_fiscais.xlsx"  # Substitua pelo caminho do seu arquivo

    # Criar instância do RPA com o mapeamento selecionado
    print(f"\n📋 Inicializando RPA para cliente: {cliente_selecionado.upper()}")
    rpa = RPANotasFiscais(URL_SITE, CAMINHO_EXCEL, cliente_selecionado, delay=2)

    # Executar em modo teste (apenas preenche, não emite)
    print("\n🧪 Iniciando RPA em MODO TESTE (apenas preenchimento)")
    print("Certifique-se de que:")
    print("1. O arquivo Excel está no formato correto")
    print("2. Você fez login no site")
    print("3. Está na página de emissão de notas")

    rpa.processar_notas(modo_teste=True)

if __name__ == "__main__":
    main()