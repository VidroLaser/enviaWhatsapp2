import time

from buscaArquivo import *
from whatsApp import *
from buscaExcel import *
from patterns import *


def main():
    try:
        arquivos = []
        mensagem_enviada = []

        pastas, phones_num, enderecos, instaladores, recorrente, dates, ids_obra = buscar_dados_excel()

        if not any(phones_num):
            print("Telefones nulos")
            return

        if pastas:
            driver, wait = config_driver()
            open_whatsapp(driver, wait)

            for pasta, phone, endereco, instalador, is_recorrente, date, obra in zip(pastas, phones_num, enderecos,
                                                                                     instaladores, recorrente, dates,
                                                                                     ids_obra):

                limpa_contato(driver)
                buscar_contato(driver, phone)
                print(f"Enviando para: {instalador} - Telefone: {phone} - Pasta: {pasta}")

                if instalador not in mensagem_enviada:
                    time.sleep(1)
                    saudacao = f"Bpm dia, {instalador}! Tudo bem?"
                    enviar_mensagem(driver, saudacao)

                time.sleep(0.5)

                ficha_name = f"{pasta}_{date}_{obra}.pdf"
                diretorio_ficha = r"\\servidor\OBRAS\FICHA SERVICO"
                ficha_pdf = buscar_arquivo(diretorio_ficha, ficha_name)

                if ficha_pdf:

                    endereco = f"Endereço da Obra: {endereco}"
                    enviar_mensagem(driver, endereco)
                    time.sleep(1)
                    mensagem = f"Segue a ficha:"
                    enviar_mensagem(driver, mensagem)
                    append_file_click(driver)
                    anexa_arquivos(driver, ficha_pdf)
                    time.sleep(2)
                    enviar(driver)

                else:
                    print(f"Ficha {ficha_name} não encontrada")

                pasta_diaria = buscar_informacao(str(obra))

                if pasta_diaria:

                    pasta = remover_letras(str(pasta))
                    dir_path = set_dir(pasta)
                    diretorio_intermediario = buscar_pasta_com_palavra(dir_path, pasta)
                    pasta_instalacao = 'INSTAL'
                    diretorio_busca = buscar_pasta_com_palavra(diretorio_intermediario, pasta_instalacao)

                    if diretorio_busca:
                        itens_diario = filter_words(pasta_diaria)

                        itens = [item_pasta for item_pasta in itens_diario]

                        for item in itens:
                            arquivo_busca = item + " - CAPA.pdf"
                            arquivo = buscar_arquivo(diretorio_busca, arquivo_busca)
                            arquivos.append(arquivo)

                            if arquivo is None and contains_ic_r(item):
                                item_limpo = remove_after_ic_r(item) + " - CAPA.pdf"
                                dir_item_limpo = buscar_arquivo(diretorio_busca, item_limpo)
                                arquivos.append(dir_item_limpo)

                        arquivos_filtrados = list(filter(lambda pdfs: pdfs is not None, arquivos))

                        if arquivos_filtrados and is_recorrente == 'N':

                            time.sleep(2)
                            mensagem = f"Segue os projetos:"
                            enviar_mensagem(driver, mensagem)
                            append_file_click(driver)

                            for paths in arquivos_filtrados:
                                anexa_arquivos(driver, paths)

                            time.sleep(5)
                            enviar(driver)

                        else:
                            print(f"Sem arquivo para enviar da pasta {pasta}")

                        if instalador not in mensagem_enviada:
                            time.sleep(2)
                            mensagem = (f"*ESTE CONTATO NÃO RESPONDE MENSAGENS*\n"
                                        f"No caso de dúvidas entrar em contato com Obras:\n"
                                        f"https://wa.me/5511953400003")

                            enviar_mensagem(driver, mensagem)

                        arquivos.clear()

                    else:
                        print("Não há itens para enviar")
                else:
                    print(f"Pasta {pasta} não encontrada no Diário de saídas")
                    limpa_contato(driver)

            time.sleep(2)
            print("Arquivos enviados - fechando em 30 segundos...")

            time.sleep(30)
            close_driver(driver)

        else:
            print("Não há dados")

    except Exception as e:
        print(f"Erro inesperado: {e}")
