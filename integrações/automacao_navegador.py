from playwright.sync_api import sync_playwright
import time

# Classe para automação do navegador utilizando Playwright
class ConectRh():
    def __init__(self):
        self.pw = None
        self.navegador = None
        self.contexto = None
        self.pagina = None
        self.pop_up = None

    def abrir_navegador(self):
        """
        Inicializa o Playwright e abre o navegador Chromium
        
        CONFIGURAÇÕES:
        - navegador: Chromium
        - headless: False (janela visível)
        - viewport: 1366x768
        """
        try:
            print(f"📍 Inicializando Playwright...")
            self.pw = sync_playwright().start()
            print(f"✅ Playwright iniciado com sucesso")
            
            print(f"📍 Abrindo navegador Chromium...")
            self.navegador = self.pw.chromium.launch(
                headless = False,
            )
            print(f"✅ Navegador Chromium aberto")

            print(f"📍 Criando contexto do navegador...")
            self.contexto = self.navegador.new_context(
                viewport = {'width': 1366, 'height': 768},
                accept_downloads = True
            )
            print(f"✅ Contexto criado (viewport: 1366x768)")
            
            print(f"📍 Criando nova página...")
            self.pagina = self.contexto.new_page()
            print(f"🎉 Navegador e página prontos para usar!")
            
        except Exception as error:
            print(f"❌ Erro ao abrir navegador: {error}")

    def logar(self, email, senha):
        """
        Realiza login na plataforma ConectRH
        
        ELEMENTOS ENCONTRADOS:
        - input_email: Textbox com nome "Login"
        - input_senha: Textbox com nome "Senha"
        - botao_login: Button com nome "Login"
        """
        tentativas = 0
        verificacao = False
        while tentativas < 3 and not verificacao:
            try:
                print(f"📍 [TENTATIVA {tentativas + 1}/3] Acessando página de login...")
                self.pagina.goto('https://cieds.conectrh.com.br/index.html')

                # INPUT: Campo de login (email)
                input_email = self.pagina.get_by_role(
                    "textbox",
                    name="Login"
                )
                input_email.wait_for(state="attached", timeout=5000)
                input_email.fill(email)
                print(f"✅ Email preenchido com sucesso")
                
                # INPUT: Campo de senha
                input_senha = self.pagina.get_by_role(
                    "textbox",
                    name="Senha"
                )
                input_senha.wait_for(state="attached", timeout=5000)
                input_senha.fill(senha)
                print(f"✅ Senha preenchida com sucesso")

                # BOTÃO: Botão de login
                botao_login = self.pagina.get_by_role(
                    "button",
                    name="Login"
                )
                botao_login.click()
                print(f"✅ Clique no botão Login realizado")

                verificacao = True
                print(f"🎉 Login realizado com sucesso!")

            except Exception as error:
                tentativas += 1
                print(f"❌ Erro ao fazer login: {error}")
                print(f"⏳ Tentando novamente... ({tentativas}/3)")
                continue

    def pagina_relatorios(self):
        """
        Navega até a página de relatórios de frequência por estudante
        
        ELEMENTOS ENCONTRADOS:
        - botao_menu: Locator com ID "#mxui_widget_SidebarToggleButton_2" (botão de menu)
        - botao_relatorios: MenuItem com nome "Relatórios Jovem Aprendiz"
        - botao_frequencias: MenuItem com nome "Presença Por Estudante"
        """
        tentativas = 0
        verificacao = False
        while tentativas < 3 and not verificacao:
            try:
                print(f"📍 [TENTATIVA {tentativas + 1}/3] Acessando página de relatórios...")
                
                # BOTÃO: Menu lateral
                botao_menu = self.pagina.locator("[id^='mxui_widget_SidebarToggleButton']").first
                botao_menu.wait_for(state="attached", timeout=5000)
                botao_menu.click()
                print(f"✅ Menu lateral aberto")

                # BOTÃO: Relatórios Jovem Aprendiz
                botao_relatorios = self.pagina.get_by_role(
                    "menuitem",
                    name="Relatórios Jovem Aprendiz"
                )
                botao_relatorios.wait_for(state="attached", timeout=5000)
                botao_relatorios.click()
                print(f"✅ Menu 'Relatórios Jovem Aprendiz' clicado")

                # BOTÃO: Presença Por Estudante
                botao_frequencias = self.pagina.get_by_role("menuitem",
                    name="Presença Por Estudante"
                )
                botao_frequencias.wait_for(state="attached", timeout=5000)
                botao_frequencias.click()
                print(f"✅ Menu 'Presença Por Estudante' clicado")
                
                verificacao = True
                print(f"🎉 Navegação para página de relatórios concluída!")
            
            except Exception as error:
                tentativas += 1
                print(f"❌ Erro ao acessar página de relatórios: {error}")
                print(f"⏳ Tentando novamente... ({tentativas}/3)")
                continue
    

    def frequencia_aprendiz(self, nome, curso):
        """
        Preenche os campos de nome e curso para filtrar frequência
        
        ELEMENTOS ENCONTRADOS:
        - input_nome: Locator com classe ".select2-selection__rendered" (seletor de nome)
        - textbox: Textbox genérico para preenchimento de nome
        - opcao_nome: Opção de resultado na lista select2
        - input_curso: Combobox com nome "CURSO"
        - opcao_curso: Opção de resultado do curso
        """
        tentativas = 0
        verificacao_geral = False

        while tentativas < 3 and not verificacao_geral:
            try:
                print(f"📍 [TENTATIVA {tentativas + 1}/3] Preenchendo dados de frequência...")
                
                # INPUT: Seletor de nome (Select2)
                input_nome = self.pagina.locator(".select2-selection__rendered")
                input_nome.wait_for(state="attached", timeout=5000)
                input_nome.click()
                print(f"✅ Seletor de nome encontrado")
                
                # INPUT: Campo de busca por nome - COM RETRY
                nome_encontrado = False
                retry_fill = 0
                while retry_fill < 2 and not nome_encontrado:
                    try:
                        textbox = self.pagina.get_by_role("textbox")
                        textbox.wait_for(state="attached", timeout=3000)
                        textbox.clear()  # Limpa antes de preencher
                        textbox.fill(nome)
                        print(f"✅ Nome '{nome}' digitado no campo de busca")
                        
                        # Aguarda um pouco para a lista aparecer
                        time.sleep(0.5)
                        
                        # OPÇÃO: Clica na opção do nome encontrado (com tratamento para strict mode)
                        try:
                            opcao_nome = self.pagina.locator(".select2-results__option").get_by_text(nome, exact=False).first
                        except Exception as strict_error:
                            # Se há múltiplos elementos (strict mode), pega o primeiro
                            if "strict mode" in str(strict_error).lower():
                                print(f"⚠️ Múltiplos elementos encontrados, selecionando o primeiro...")
                                opcao_nome = self.pagina.locator(".select2-results__option").get_by_text(
                                    nome, 
                                    exact=False
                                ).first
                            else:
                                raise
                        
                        opcao_nome.wait_for(state="attached", timeout=3000)
                        opcao_nome.click()
                        print(f"✅ Nome '{nome}' selecionado da lista")
                        nome_encontrado = True
                        
                    except Exception as e:
                        retry_fill += 1
                        if retry_fill < 2:
                            print(f"⚠️ Tentando digitar novamente... ({retry_fill}/2)")
                            time.sleep(0.3)
                            input_nome.click()
                        else:
                            raise Exception(f"❌ Aluno '{nome}' não pode ser selecionado: {str(e)[:80]}")

                # INPUT: Campo de seleção de curso
                input_curso = self.pagina.get_by_role(
                    "combobox",
                    name="CURSO"
                )
                input_curso.click()
                input_curso.fill(curso)
                print(f"✅ Curso '{curso}' digitado no campo")

                # OPÇÃO: Clica na opção do curso
                opcao_curso = self.pagina.locator(".widget-combobox-menu").get_by_text(curso, exact=False).first
                opcao_curso.wait_for(state="visible", timeout=5000)
                opcao_curso.click()
                # opcao_curso.click(force=True)
                print(f"✅ Curso '{curso}' selecionado")

                verificacao_geral = True
                print(f"🎉 Dados de frequência preenchidos com sucesso!")

            except Exception as error:
                tentativas += 1
                print(f"❌ Erro ao preencher campos de frequência: {error}")
                if tentativas < 3:
                    print(f"⏳ Tentando novamente... ({tentativas}/3)")
                continue

    def pop_up_frequencias(self):
        """
        Clica no botão 'Gerar Relatório' e captura o pop-up aberto
        
        ELEMENTOS ENCONTRADOS:
        - botao_gerar: Button com nome "Gerar Relatório"
        """
        try:
            print(f"📍 Gerando relatório e capturando pop-up...")
            
            # BOTÃO: Gerar Relatório
            with self.pagina.context.expect_page() as popup_info:
                botao_gerar = self.pagina.get_by_role(
                    "button",
                    name="Gerar Relatório"
                )
                botao_gerar.click()
                print(f"✅ Botão 'Gerar Relatório' clicado")

            self.pop_up = popup_info.value
            print(f"🎉 Pop-up de relatório capturado com sucesso!")

        except Exception as error:
            print(f"❌ Erro ao gerar relatório: {error}")

    def pegar_valor_frequencia(self):
        """
        Extrai o valor de frequência (%) do pop-up
        
        ELEMENTOS ENCONTRADOS:
        - frequencia: Elemento de texto contendo "%"
        """
        try:
            print(f"📍 Extraindo valor de frequência...")
            
            # ELEMENTO: Texto com símbolo de percentual
            frequencia = self.pop_up.get_by_text("%").nth(1)
            frequencia.wait_for(state="visible", timeout=5000)
            valor_frequencia = frequencia.text_content()
            
            print(f"✅ Frequência extraída: {valor_frequencia}")
            return valor_frequencia

        except Exception as error:
            self.fechar_popup()
            print(f"❌ Erro ao extrair frequência: {error}")
            return None

    def pagina_faltas(self):
        """
        Navega até a página de presença por empresa parceira (faltas)
        
        ELEMENTOS ENCONTRADOS:
        - botao_menu: Locator com ID "#mxui_widget_SidebarToggleButton_2" (botão de menu)
        - botao_relatorio: MenuItem com nome "Relatórios Jovem Aprendiz"
        - presenca_empresa: MenuItem com nome "Presença Por Empresa Parceira"
        """
        tentativas = 0
        verificacao = False
        while tentativas < 3 and not verificacao:
            try:
                print(f"📍 [TENTATIVA {tentativas + 1}/3] Acessando página de faltas...")
                
                # BOTÃO: Menu lateral
                botao_menu = self.pagina.locator("#mxui_widget_SidebarToggleButton_2").first
                botao_menu.wait_for(state="attached", timeout=10000)
                botao_menu.click()
                print(f"✅ Menu lateral aberto")

                # BOTÃO: Relatórios Jovem Aprendiz
                botao_relatorio = self.pagina.get_by_role(
                    "menuitem",
                    name="Relatórios Jovem Aprendiz"
                )
                botao_relatorio.wait_for(state="attached", timeout=5000)
                botao_relatorio.click()
                print(f"✅ Menu 'Relatórios Jovem Aprendiz' clicado")

                # BOTÃO: Presença Por Empresa Parceira
                presenca_empresa = self.pagina.get_by_role(
                    "menuitem",
                    name="Presença Por Empresa Parceira"
                )
                presenca_empresa.wait_for(state="attached", timeout=5000)
                presenca_empresa.click()
                print(f"✅ Menu 'Presença Por Empresa Parceira' clicado")
                
                verificacao = True
                print(f"🎉 Navegação para página de faltas concluída!")
                
            except Exception as error:
                tentativas += 1
                print(f"❌ Erro ao acessar página de faltas: {error}")
                print(f"⏳ Tentando novamente... ({tentativas}/3)")
                continue

    def baixar_faltas(self, data):
        """
        Preenche as datas e gera relatório detalhado de faltas
        
        ARGUMENTOS:
        - data (str): Data no formato esperado (DD/MM/YYYY ou similar)
        
        ELEMENTOS ENCONTRADOS:
        - input_data1: Textbox com nome "Data De" (data inicial)
        - input_data2: Textbox com nome "Data Até" (data final)
        - botao_relatorio: Button com nome "Gerar Relatório Detalhado"
        """
        verificacao = False
        tentativas = 0
        while verificacao == False and tentativas < 3:
            try:
                print(f"📍 [TENTATIVA {tentativas + 1}/3] Gerando relatório de faltas...")
                
                # INPUT: Campo de data inicial
                input_data1 = self.pagina.get_by_role("textbox", name="Data De")
                input_data1.wait_for(state="attached", timeout=5000)
                input_data1.fill(data)
                if not input_data1.text_content() == data:
                    print(f"⚠️ Aviso: O valor digitado no campo de data inicial pode não ter sido aceito corretamente.")
                    input_data1.fill(data)
                print(f"✅ Data inicial preenchida: {data}")

                # INPUT: Campo de data final
                input_data2 = self.pagina.get_by_role("textbox", name="Data Até")
                input_data2.wait_for(state="attached", timeout=5000)
                input_data2.fill(data)
                if not input_data2.text_content() == data:
                    print(f"⚠️ Aviso: O valor digitado no campo de data final pode não ter sido aceito corretamente.")
                    input_data2.fill(data)
                print(f"✅ Data final preenchida: {data}")

                # BOTÃO: Gerar Relatório Detalhado
                print(f"✅ Botão 'Gerar Relatório Detalhado' clicado")
                with self.pagina.expect_download() as download_info:
                    botao_relatorio = self.pagina.get_by_role("button", name="Gerar Relatório Detalhado")
                    botao_relatorio.wait_for(state="attached", timeout=5000)
                    botao_relatorio.click()
                    print(f"✅ Botão 'Gerar Relatório Detalhado' clicado")
                    download = download_info.value
                    # Aqui você salva com o nome que o site sugeriu ou o que você escolher
                    path_0 = r'C:\Users\danielsampaio.rj\Desktop\Pasta de Teste de Scripts Python\cieds_bot_v02\planilhas'
                    download.save_as(f"{path_0}/{download.suggested_filename}")

                botao_ok = self.pagina.get_by_role("button", name="Ok")
                botao_ok.wait_for(state='attached', timeout=5000)
                botao_ok.click()
                
                verificacao = True
                print(f"🎉 Relatório de faltas gerado com sucesso!")
                
            except Exception as error:
                tentativas += 1
                print(f"❌ Erro ao gerar relatório de faltas: {error}")
                print(f"⏳ Tentando novamente... ({tentativas}/3)")
                continue

    def pagina_contratos(self):
        """
        Navega até a página de contratos de jovem aprendiz
        
        ELEMENTOS ENCONTRADOS:
        - botao_menu: Locator com ID "#mxui_widget_SidebarToggleButton_2" (botão de menu)
        - botao_aprendiz: MenuItem com nome "Jovem Aprendiz"
        - botao_contrato: MenuItem com nome "Contrato Jovem Aprendiz"
        """
        tentativas = 0
        verificacao = False
        while tentativas < 3 and not verificacao:
            try:
                print(f"📍 [TENTATIVA {tentativas + 1}/3] Acessando página de contratos...")
                self.pagina_inicial()
                # BOTÃO: Menu lateral
                time.sleep(1)
                botao_menu = self.pagina.locator("[id^='mxui_widget_SidebarToggleButton']").first
                botao_menu.wait_for(state="visible", timeout=5000)
                botao_menu.click()
                print(f"✅ Menu lateral aberto")
                
                # BOTÃO: Jovem Aprendiz
                botao_aprendiz = self.pagina.get_by_role(
                    "menuitem",
                    name="Jovem Aprendiz", exact=True
                )
                botao_aprendiz.wait_for(state="attached", timeout=10000)
                botao_aprendiz.click()
                print(f"✅ Menu 'Jovem Aprendiz' clicado")

                # BOTÃO: Contrato Jovem Aprendiz
                botao_contrato = self.pagina.get_by_role(
                    "menuitem", 
                    name="Contrato Jovem Aprendiz"
                )
                botao_contrato.wait_for(state="attached", timeout=5000)
                botao_contrato.click()
                print(f"✅ Menu 'Contrato Jovem Aprendiz' clicado")

                verificacao = True
                print(f"🎉 Navegação para página de contratos concluída!")

            except Exception as error:
                tentativas += 1
                print(f"❌ Erro ao acessar página de contratos: {error}")
                print(f"⏳ Tentando novamente... ({tentativas}/3)")
                continue

    def baixar_contratos(self):
        verificao = False
        tentativas = 0
        while tentativas < 3 and not verificao:
            try:
                botao_exportar2 = self.pagina.locator('button[data-button-id="p.Jovem_Aprendiz.ContratoJovemAprendiz_Listagem.actionButton8"]').filter(visible=True).first
                botao_exportar2.wait_for(state='attached', timeout=5000)
                botao_exportar2.click()
                print(f"📍 [TENTATIVA {tentativas + 1}/3] Baixando planilha de contratos...")
                with self.pagina.expect_download() as download_info:
                    exportar_1 = self.pagina.get_by_label("Exportar Contratos").get_by_role("button", name="Exportar")
                    exportar_1.wait_for(state='attached', timeout=5000)
                    exportar_1.click()
                    print(f"✅ Botão 'Exportar contratos' clicado")
                    download = download_info.value
                    # Aqui você salva com o nome que o site sugeriu ou o que você escolher
                    path_0 = r'C:\Users\danielsampaio.rj\Desktop\Pasta de Teste de Scripts Python\cieds_bot_v02\planilhas'
                    download.save_as(f"{path_0}/{download.suggested_filename}")

                print ('Planilha de contratos baixada! ')

                verificao = True

            except Exception as error:
                print (error)
                botao_fechar = self.pagina.get_by_role("button", name="Fechar")
                if botao_fechar.is_visible():
                    botao_fechar.click()
                else:
                    pass
                print ('Tentando novamente... ')
                tentativas += 1
                continue


    def fechar_popup(self):
        try:
            if self.pop_up and not self.pop_up.is_closed():
                self.pop_up.close()
                self.pop_up = None
        except Exception as e:
            self.pop_up = None
            pass

    def pagina_inicial(self):
        self.pagina.reload(wait_until='domcontentloaded')

    def fechar_navegador(self):
        if self.navegador:
            self.navegador.close()

        if self.pw:
            self.pw.gstop()