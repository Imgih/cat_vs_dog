from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys  # Para pressionar Enter
from selenium.webdriver.common.action_chains import ActionChains
from openpyxl import Workbook, load_workbook
import time
import os
from collections import defaultdict
from datetime import datetime
import notificacoes  # Importando o módulo de notificações
import gerar_relatorio_diario  # Importando o módulo para gerar o relatório diário

# Palavras-chave para filtrar tickets pelo nome exato da operação (ignorando maiúsculas/minúsculas)
PALAVRAS_CHAVE_OPERACAO = []  # Lista vazia para buscar todos os tickets

# Status que NÃO queremos monitorar
STATUS_FECHADOS = ["Fechado", "Resolvido", "Cancelado"]

# Nome do arquivo Excel
EXCEL_ARQUIVO = "tickets_monitorados.xlsx"

# Configuração do Selenium
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")  # Inicia com a janela maximizada
driver = webdriver.Chrome(options=options)

# URL de login (substitua com a URL correta do login)
url_login = "https://exemplo.com/login"  # Substitua com o URL do login da sua página
driver.get(url_login)

# Função para realizar login na página
def fazer_login():
    """Realiza o login na página de login"""
    time.sleep(2)  # Aguarda carregamento da página

    # Localiza os campos de login (substitua os seletores pelos corretos da sua página)
    campo_usuario = driver.find_element(By.ID, "campo_usuario")  # Substitua pelo ID correto do campo de usuário
    campo_senha = driver.find_element(By.ID, "campo_senha")  # Substitua pelo ID correto do campo de senha

    # Insere os dados no formulário de login
    campo_usuario.send_keys("seu_usuario")  # Substitua pelo seu nome de usuário
    campo_senha.send_keys("sua_senha")  # Substitua pela sua senha

    # Envia o formulário (clicando no botão de login ou pressionando Enter)
    botao_login = driver.find_element(By.ID, "botao_login")  # Substitua pelo ID correto do botão de login
    botao_login.click()  # Clica no botão de login

    # Aguarda a página carregar
    time.sleep(5)  # Aguarde o tempo necessário para o login ser processado

def capturar_tickets():
    """Captura tickets diretamente do HTML sem precisar atualizar a página manualmente."""
    tickets = driver.find_elements(By.CLASS_NAME, "ticket-row")
    tickets_relevantes = {}

    for ticket in tickets:
        try:
            ticket_id = ticket.get_attribute("data-id")  # Pegando o valor do atributo data-id

            # Buscando os outros dados dentro do ticket usando o ID de data-id
            titulo = ticket.find_element(By.CSS_SELECTOR, "[data-id='ticket-Assunto']").text.strip()  # Usando data-id
            operacao = ticket.find_element(By.CSS_SELECTOR, "[data-id='ticket-operacao']").text.strip()  # Usando data-id
            data = ticket.find_element(By.CSS_SELECTOR, "[data-id='ticket-data']").text.strip()  # Usando data-id
            descricao = ticket.find_element(By.CSS_SELECTOR, "[data-id='ticket-description']").text.strip()  # Usando data-id
            status = ticket.find_element(By.CSS_SELECTOR, "[data-id='ticket-status']").text.strip()  # Usando data-id

            url_element = ticket.find_element(By.CSS_SELECTOR, "[data-id='ticket-link']")  # Usando data-id
            url = url_element.get_attribute("href") if url_element else ""

            # Converte a data para o formato datetime (yyyy-mm-dd)
            data_formatada = datetime.strptime(data, "%d/%m/%Y").date()

            if not PALAVRAS_CHAVE_OPERACAO or any(palavra.strip().lower() == operacao.lower() for palavra in PALAVRAS_CHAVE_OPERACAO):
                if status not in STATUS_FECHADOS:
                    tickets_relevantes[ticket_id] = {
                        "titulo": titulo,
                        "operacao": operacao,
                        "data": data_formatada,
                        "status": status,
                        "descricao": descricao,
                        "url": url
                    }

        except Exception as e:
            print(f"Erro ao capturar ticket: {e}")
            continue
    
    return tickets_relevantes

def atualizar_excel(novos_tickets, tickets_atualizados):
    """Atualiza a planilha Excel com novos tickets e mudanças de status."""
    wb = load_workbook(EXCEL_ARQUIVO)
    ws = wb["Tickets"]

    ids_existentes = {}
    for i in range(2, ws.max_row + 1):
        id_ticket = str(ws.cell(row=i, column=1).value)
        ids_existentes[id_ticket] = i

    for ticket_id, dados in novos_tickets.items():
        if ticket_id in ids_existentes:
            continue

        # Adiciona o ticket à planilha
        ws.append([ticket_id,
                   dados["titulo"],
                   dados["operacao"],
                   dados["data"],
                   dados["status"],
                   dados.get("descricao", ""),
                   dados.get("url", "")])
        relatorio_diario[dados["data"]]["novos"] += 1
        print(f"🆕 Novo ticket registrado: {ticket_id} - {dados['titulo']}")

        # Chama a função de notificação para o ticket aberto
        notificacoes.notificar_abertura_ticket(ticket_id, dados["titulo"])

    for ticket_id, dados in tickets_atualizados.items():
        if ticket_id in ids_existentes:
            linha = ids_existentes[ticket_id]
            ws.cell(row=linha, column=5, value=dados["status"])
            relatorio_diario[dados["data"]]["alterados"] += 1
            print(f"🔄 Status atualizado: {ticket_id} - {dados['titulo']} -> {dados['status']}")

            # Chama a função de notificação para a alteração de status
            notificacoes.notificar_alteracao_status(ticket_id, dados["titulo"], dados["status"])

    wb.save(EXCEL_ARQUIVO)

def gerar_relatorio():
    """Gera um relatório diário com a quantidade de tickets criados e com status alterado."""
    global ultimo_relatorio_data
    data_atual = datetime.today().date()

    if data_atual != ultimo_relatorio_data:
        print(f"\n📝 Relatório Diário - {ultimo_relatorio_data}:")
        for data, stats in relatorio_diario.items():
            print(f"Data: {data}")
            print(f"  - Tickets Criados: {stats['novos']}")
            print(f"  - Status Alterados: {stats['alterados']}")

        # Salva o relatório no formato yyyy-mm-dd
        data_formatada_str = ultimo_relatorio_data.strftime("%Y-%m-%d")
        with open(f"relatorio_{data_formatada_str}.txt", "w") as file:
            for data, stats in relatorio_diario.items():
                file.write(f"Data: {data}\n")
                file.write(f"  - Tickets Criados: {stats['novos']}\n")
                file.write(f"  - Status Alterados: {stats['alterados']}\n")
                file.write("\n")

        relatorio_diario.clear()
        ultimo_relatorio_data = data_atual

def interagir_com_pagina():
    """Movimenta o mouse e recarrega a página para garantir atualização dos elementos."""
    action = ActionChains(driver)
    action.move_by_offset(5, 5).perform()
    time.sleep(0.5)
    action.move_by_offset(-5, -5).perform()
    time.sleep(0.5)
    driver.refresh()
    print("🔄 Página recarregada para atualização dos tickets.")
    time.sleep(2)

def menu_opcoes():
    """Menu de opções para o usuário escolher a ação desejada."""
    while True:
        print("\nEscolha uma opção:")
        print("1. Iniciar monitoramento e gerar relatório diário")
        print("2. Gerar relatório diário específico")
        print("3. Sair")
        
        opcao = input("Digite a opção desejada (1, 2 ou 3): ")
        
        if opcao == "1":
            iniciar_monitoramento()
        elif opcao == "2":
            gerar_relatorio_diario.gerar_relatorio_diario()  # Chama a função para gerar o relatório diário
        elif opcao == "3":
            print("Saindo...")
            break
        else:
            print("Opção inválida. Tente novamente.")

def iniciar_monitoramento():
    """Inicia o monitoramento e gera o relatório diário"""
    # Realiza o login primeiro
    fazer_login()

    print("📊 Iniciando monitoramento...")
    tickets_vistos = capturar_tickets()
    atualizar_excel(tickets_vistos, {})

    if not tickets_vistos:
        print("✅ Nenhum ticket relevante encontrado. Encerrando script.")
        driver.quit()
        return

    print("✅ Planilha inicializada com os tickets encontrados.")
    print("🔍 Monitorando mudanças em tempo real...")

    try:
        while True:
            interagir_com_pagina()
            novos_tickets = capturar_tickets()

            if not novos_tickets:
                print("✅ Nenhum novo ticket encontrado.")
                time.sleep(5)
                continue

            tickets_adicionados = {k: v for k, v in novos_tickets.items() if k not in tickets_vistos}
            tickets_atualizados = {
                k: v for k, v in novos_tickets.items()
                if k in tickets_vistos and v["status"] != tickets_vistos[k]["status"]
            }

            if tickets_adicionados or tickets_atualizados:
                atualizar_excel(tickets_adicionados, tickets_atualizados)
                tickets_vistos.update(novos_tickets)

            gerar_relatorio()
            time.sleep(5)

    except KeyboardInterrupt:
        print("\n🛑 Monitoramento encerrado manualmente.")
        driver.quit()

# Inicia o menu
menu_opcoes()
