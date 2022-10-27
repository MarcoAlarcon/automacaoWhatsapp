from multiprocessing import context
from sre_parse import State
import pandas as pd
import time
from playwright.sync_api import sync_playwright


#Informar entre aspas duplas o caminho inteiro do arquivo excel e sua extensão.
excel = pd.read_excel("C:\projetos\AMC\envioWhatsapp\envioWhatsapp.xlsx")

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://web.whatsapp.com")

    #Espera 15 segundos para validação do do QR code e acesso a conta
    time.sleep(20)

    for i, linhas in enumerate(excel['Telefone']):
        #Guarda o número de contato do index i da planilha
        telefone = excel.loc[i,"Telefone"]

        #Armazena a mensagem que foi digitada na celula abaixo do campo mensagem do excel.
        mensagem = excel.iloc[0,4]
        caminhoArquivo = excel.iloc[0,5]
        link = f'https://web.whatsapp.com/send?phone={telefone}&text={mensagem}'

        try:
            page.goto(link)
            time.sleep(6)
            mensagemAlerta = page.locator("._2Nr6U")
            popupErro = page.locator("//*[@id='app']/div/span[2]/div/span/div/div/div/div/div/div[2]/div/div/div/div", has_text='OK')
            if mensagemAlerta.is_visible():
                popupErro.click()
                continue
            # with page.expect_navigation():
            #     page.locator("div[role=/'textbox/']")

            page.locator("[data-testid='compose-btn-send']").click()

            if caminhoArquivo:
                anexarArquivo = page.locator("//*[@id='main']/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div")
                anexarArquivo.click()
                caixaArquivo = page.locator("//*[@id='main']/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/div/ul/li[1]/button")
                with page.expect_file_chooser() as fc_info:
                    caixaArquivo.click()
                file_chooser = fc_info.value
                file_chooser.set_files(caminhoArquivo)
                botaoEnviar = page.locator("//*[@id='app']/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div")
                botaoEnviar.click()

            time.sleep(5)

        except:
            continue
